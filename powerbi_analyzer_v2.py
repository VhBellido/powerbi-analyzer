#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Power BI Analyzer V2 - Deep Mode
--------------------------------
Conecta diretamente ao modelo aberto no Power BI Desktop através do
Analysis Services local (MSOLAP/OLE DB), coleta TMSCHEMA + DMVs VertiPaq,
executa regras de boas práticas e gera PDF + JSON + CSV.

Uso principal:
  Executar pelo menu Ferramentas Externas do Power BI Desktop.

Uso manual:
  python powerbi_analyzer_v2.py --server "localhost:12345" --database "GUID"

Versão: 2.0.0
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import sys
import traceback
from collections import defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

APP_VERSION = "2.0.0"

try:
    import win32com.client
except ImportError:
    print("Dependência ausente: pywin32. Execute INSTALAR_V2.ps1.")
    sys.exit(2)

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        PageBreak, KeepTogether
    )
except ImportError:
    print("Dependência ausente: reportlab. Execute INSTALAR_V2.ps1.")
    sys.exit(2)


CATEGORIES = [
    "Modelo de dados",
    "DAX",
    "Nomenclatura e organização",
    "Performance / VertiPaq",
    "Consultas M / Fontes",
]

SEVERITY_WEIGHT = {
    "Crítica": 12,
    "Alta": 7,
    "Média": 4,
    "Baixa": 1,
    "Informativa": 0,
}

TECHNICAL_NAME_RE = re.compile(
    r"(^id($|_)|_id$|(^|_)key$|(^|_)sk$|(^|_)fk$|^cod($|_)|_cod$|"
    r"surrogate|rowguid|guid$|hash$|timestamp$|^chave|chave$)",
    re.I
)
CALENDAR_NAME_RE = re.compile(
    r"(calendar|calend[aá]rio|dim[_ ]?date|dim[_ ]?data|d[_ ]?data|datas?)",
    re.I
)
FACT_NAME_RE = re.compile(
    r"(^f[_ ]|^fact|fato|movimento|vendas|transa[cç]|detalhado|lan[cç]amento)",
    re.I
)


@dataclass
class Finding:
    id: str
    category: str
    severity: str
    title: str
    object_type: str
    object_name: str
    evidence: str
    impact: str
    recommendation: str
    confidence: str = "Alta"
    metric: Optional[float] = None


def clean(v: Any) -> Any:
    if v is None:
        return None
    try:
        # COM DBNull/DateTime/etc.
        if hasattr(v, "isoformat"):
            return v.isoformat()
    except Exception:
        pass
    if isinstance(v, (str, int, float, bool)):
        return v
    try:
        return int(v)
    except Exception:
        pass
    try:
        return float(v)
    except Exception:
        pass
    return str(v)


def key_upper(row: Dict[str, Any]) -> Dict[str, Any]:
    return {str(k).upper(): v for k, v in row.items()}


def pick(row: Dict[str, Any], *names: str, default=None):
    r = key_upper(row)
    for n in names:
        if n.upper() in r and r[n.upper()] not in (None, ""):
            return r[n.upper()]
    return default


def s(v: Any) -> str:
    return "" if v is None else str(v)


def num(v: Any, default: float = 0.0) -> float:
    if v is None or v == "":
        return default
    try:
        return float(v)
    except Exception:
        try:
            return float(str(v).replace(",", "."))
        except Exception:
            return default


def as_int(v: Any, default: int = 0) -> int:
    try:
        return int(float(v))
    except Exception:
        return default


def truth(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in {"1", "true", "yes", "sim", "-1"}


def mb(v: Any) -> float:
    return num(v) / 1024.0 / 1024.0


def fmt_int(v: Any) -> str:
    try:
        return f"{int(float(v)):,}".replace(",", ".")
    except Exception:
        return "0"


def fmt_mb(v: Any) -> str:
    return f"{mb(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def html(sv: Any) -> str:
    x = s(sv)
    return x.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def add(findings: List[Finding], **kwargs):
    findings.append(Finding(**kwargs))


class AdoAnalysisServices:
    """Conexão read-only via ADO + provider MSOLAP instalado pelo Power BI Desktop."""

    def __init__(self, server: str, database: str, timeout: int = 180):
        self.server = server
        self.database = database
        self.timeout = timeout
        self.conn = None

    def open(self):
        self.conn = win32com.client.Dispatch("ADODB.Connection")
        try:
            self.conn.ConnectionTimeout = 30
        except Exception:
            pass
        try:
            self.conn.CommandTimeout = self.timeout
        except Exception:
            pass
        conn_str = (
            f"Provider=MSOLAP;"
            f"Data Source={self.server};"
            f"Initial Catalog={self.database};"
            f"Persist Security Info=True;"
        )
        self.conn.Open(conn_str)

    def close(self):
        try:
            if self.conn is not None:
                self.conn.Close()
        except Exception:
            pass

    def query(self, query: str, timeout: Optional[int] = None) -> List[Dict[str, Any]]:
        if self.conn is None:
            self.open()

        cmd = win32com.client.Dispatch("ADODB.Command")
        cmd.ActiveConnection = self.conn
        cmd.CommandText = query
        try:
            cmd.CommandTimeout = int(timeout or self.timeout)
        except Exception:
            pass

        result = cmd.Execute()
        rs = result[0] if isinstance(result, tuple) else result

        if rs is None:
            return []

        fields = []
        try:
            count = rs.Fields.Count
            fields = [s(rs.Fields.Item(i).Name) for i in range(count)]
        except Exception:
            return []

        rows = []
        try:
            while not rs.EOF:
                row = {}
                for i, name in enumerate(fields):
                    try:
                        row[name] = clean(rs.Fields.Item(i).Value)
                    except Exception:
                        row[name] = None
                rows.append(row)
                rs.MoveNext()
        finally:
            try:
                rs.Close()
            except Exception:
                pass
        return rows

    def safe_query(self, query: str, warnings: List[str], label: str,
                   timeout: Optional[int] = None) -> List[Dict[str, Any]]:
        try:
            return self.query(query, timeout=timeout)
        except Exception as exc:
            warnings.append(f"{label}: {exc}")
            return []


DMV_QUERIES = {
    "tables": "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES",
    "columns": "SELECT * FROM $SYSTEM.TMSCHEMA_COLUMNS",
    "measures": "SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES",
    "relationships": "SELECT * FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS",
    "partitions": "SELECT * FROM $SYSTEM.TMSCHEMA_PARTITIONS",
    "expressions": "SELECT * FROM $SYSTEM.TMSCHEMA_EXPRESSIONS",
    "data_sources": "SELECT * FROM $SYSTEM.TMSCHEMA_DATA_SOURCES",
    "calc_dependency": "SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY",
    "storage_tables": "SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES",
    "storage_columns": "SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS",
    "storage_segments": "SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS",
}


def load_config(path: Optional[Path]) -> Dict[str, Any]:
    default = {
        "thresholds": {
            "wide_table_columns": 40,
            "calculated_columns_per_table": 5,
            "large_column_mb": 25,
            "very_large_column_mb": 100,
            "large_dictionary_mb": 15,
            "very_large_dictionary_mb": 60,
            "large_table_mb": 250,
            "very_large_table_mb": 1000,
            "high_cardinality_min_rows": 100000,
            "high_cardinality_ratio": 0.80,
            "exact_cardinality_top_n": 25,
            "exact_cardinality_timeout_seconds": 90,
            "dax_long_chars": 2500
        }
    }
    p = path or (Path(__file__).resolve().parent / "rules_config.json")
    if p.exists():
        try:
            user = json.loads(p.read_text(encoding="utf-8-sig"))
            if "thresholds" in user:
                default["thresholds"].update(user["thresholds"])
        except Exception:
            pass
    return default


def collect_live(server: str, database: str, config: Dict[str, Any]) -> Dict[str, Any]:
    warnings: List[str] = []
    conn = AdoAnalysisServices(server, database)
    data: Dict[str, Any] = {"warnings": warnings}
    try:
        print("[1/7] Conectando ao modelo aberto...")
        conn.open()

        print("[2/7] Coletando metadados TMSCHEMA...")
        for k in ("tables", "columns", "measures", "relationships",
                  "partitions", "expressions", "data_sources", "calc_dependency"):
            data[k] = conn.safe_query(DMV_QUERIES[k], warnings, k)

        print("[3/7] Coletando métricas VertiPaq...")
        for k in ("storage_tables", "storage_columns", "storage_segments"):
            data[k] = conn.safe_query(DMV_QUERIES[k], warnings, k, timeout=300)

        print("[4/7] Calculando memória, linhas e candidatos de cardinalidade...")
        metrics = build_deep_metrics(data)
        data["metrics"] = metrics

        top_n = int(config["thresholds"].get("exact_cardinality_top_n", 25))
        card_timeout = int(config["thresholds"].get("exact_cardinality_timeout_seconds", 90))
        candidates = [
            c for c in metrics["columns"]
            if c.get("table") and c.get("column") and c.get("is_basic_data", False)
        ]
        candidates.sort(key=lambda x: x.get("total_bytes", 0), reverse=True)
        candidates = candidates[:max(0, top_n)]

        print(f"[5/7] Cardinalidade exata nas {len(candidates)} colunas mais pesadas...")
        cardinality = {}
        table_row_cache = {}
        for idx, c in enumerate(candidates, 1):
            tn, cn = c["table"], c["column"]
            dax = (
                'EVALUATE ROW('
                '"__Cardinality", COUNTROWS(VALUES('
                + dax_table(tn) + "[" + dax_column(cn) + "]"
                + ')))'
            )
            try:
                rows = conn.query(dax, timeout=card_timeout)
                if rows:
                    val = pick(rows[0], "__Cardinality", "[__Cardinality]", default=None)
                    if val is None:
                        # OLE DB pode devolver nome qualificado.
                        for k, v in rows[0].items():
                            if "CARDINALITY" in str(k).upper():
                                val = v
                                break
                    if val is not None:
                        cardinality[(tn, cn)] = as_int(val)
            except Exception as exc:
                warnings.append(f"Cardinalidade {tn}[{cn}]: {exc}")

        for c in metrics["columns"]:
            key = (c.get("table"), c.get("column"))
            if key in cardinality:
                c["cardinality"] = cardinality[key]
                rows = c.get("row_count", 0)
                c["cardinality_ratio"] = (
                    cardinality[key] / rows if rows and rows > 0 else None
                )

        return data
    finally:
        conn.close()


def dax_table(name: str) -> str:
    return "'" + s(name).replace("'", "''") + "'"


def dax_column(name: str) -> str:
    return s(name).replace("]", "]]")


def table_name_map(data: Dict[str, Any]) -> Dict[str, str]:
    out = {}
    for r in data.get("tables", []):
        rid = s(pick(r, "ID"))
        name = s(pick(r, "Name", "ExplicitName", "InferredName"))
        if rid:
            out[rid] = name
    return out


def column_meta_map(data: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    out = {}
    for r in data.get("columns", []):
        rid = s(pick(r, "ID"))
        if rid:
            out[rid] = r
    return out


def build_deep_metrics(data: Dict[str, Any]) -> Dict[str, Any]:
    # Row count by model table. Storage DMV can return helper rows, relationships and hierarchies.
    row_counts: Dict[str, int] = defaultdict(int)
    row_seen = set()
    for r in data.get("storage_tables", []):
        dim = s(pick(r, "DIMENSION_NAME"))
        tid = s(pick(r, "TABLE_ID"))
        part = s(pick(r, "PARTITION_NAME"))
        rc = as_int(pick(r, "ROWS_COUNT"))
        if not dim or tid.upper().startswith(("H$", "R$")):
            continue
        # Deduplica entradas idênticas de partição.
        dedupe = (dim, tid, part, rc)
        if dedupe in row_seen:
            continue
        row_seen.add(dedupe)
        # Alguns engines retornam uma linha total e linhas auxiliares.
        # Conservador: guardamos o maior total observado para a dimensão.
        row_counts[dim] = max(row_counts[dim], rc)

    # Dicionário e metadados físicos.
    storage_col = {}
    for r in data.get("storage_columns", []):
        ctype = s(pick(r, "COLUMN_TYPE")).upper()
        dim = s(pick(r, "DIMENSION_NAME"))
        col = s(pick(r, "ATTRIBUTE_NAME", "COLUMN_ID"))
        col_id = s(pick(r, "COLUMN_ID"))
        if not dim or not col:
            continue
        key = (dim, col_id or col)
        storage_col[key] = {
            "table": dim,
            "column": col,
            "column_id": col_id or col,
            "dictionary_bytes": as_int(pick(r, "DICTIONARY_SIZE")),
            "data_type_physical": s(pick(r, "DATATYPE", "DATA_TYPE")),
            "encoding": pick(r, "COLUMN_ENCODING"),
            "is_key": truth(pick(r, "ISKEY")),
            "is_unique": truth(pick(r, "ISUNIQUE")),
            "is_nullable": truth(pick(r, "ISNULLABLE")),
            "is_row_number": truth(pick(r, "ISROWNUMBER")),
            "is_basic_data": ctype in ("", "BASIC_DATA"),
        }

    # Segmentos: somamos Used/Allocated por coluna.
    seg = defaultdict(lambda: {
        "used_bytes": 0, "allocated_bytes": 0, "records_count": 0,
        "segments": 0, "bits_max": 0
    })
    for r in data.get("storage_segments", []):
        dim = s(pick(r, "DIMENSION_NAME"))
        col_id = s(pick(r, "COLUMN_ID"))
        if not dim or not col_id:
            continue
        key = (dim, col_id)
        seg[key]["used_bytes"] += as_int(pick(r, "USED_SIZE"))
        seg[key]["allocated_bytes"] += as_int(pick(r, "ALLOCATED_SIZE"))
        seg[key]["records_count"] += as_int(pick(r, "RECORDS_COUNT"))
        seg[key]["segments"] += 1
        seg[key]["bits_max"] = max(seg[key]["bits_max"], as_int(pick(r, "BITS_COUNT")))

    # Une por (DIMENSION_NAME, COLUMN_ID). COLUMN_ID normalmente é o nome da coluna no VertiPaq.
    column_metrics = []
    all_keys = set(storage_col) | set(seg)
    for key in all_keys:
        sc = storage_col.get(key, {})
        sg = seg.get(key, {})
        table = sc.get("table") or key[0]
        column = sc.get("column") or key[1]
        dict_b = as_int(sc.get("dictionary_bytes"))
        used_b = as_int(sg.get("used_bytes"))
        alloc_b = as_int(sg.get("allocated_bytes"))
        total = dict_b + used_b
        row_count = row_counts.get(table, 0)
        column_metrics.append({
            **sc,
            "table": table,
            "column": column,
            "is_basic_data": sc.get("is_basic_data", False),
            "segment_used_bytes": used_b,
            "segment_allocated_bytes": alloc_b,
            "dictionary_bytes": dict_b,
            "total_bytes": total,
            "row_count": row_count,
            "segment_records_count": as_int(sg.get("records_count")),
            "segments": as_int(sg.get("segments")),
            "bits_max": as_int(sg.get("bits_max")),
            "cardinality": None,
            "cardinality_ratio": None,
        })

    # Totais por tabela.
    table_mem = defaultdict(int)
    table_dict = defaultdict(int)
    for c in column_metrics:
        if c.get("is_basic_data", False):
            table_mem[c["table"]] += as_int(c.get("total_bytes"))
            table_dict[c["table"]] += as_int(c.get("dictionary_bytes"))

    tables_metrics = []
    all_tables = set(row_counts) | set(table_mem)
    for t in all_tables:
        tables_metrics.append({
            "table": t,
            "row_count": row_counts.get(t, 0),
            "total_bytes": table_mem.get(t, 0),
            "dictionary_bytes": table_dict.get(t, 0),
        })

    tables_metrics.sort(key=lambda x: x["total_bytes"], reverse=True)
    column_metrics.sort(key=lambda x: x["total_bytes"], reverse=True)

    return {
        "tables": tables_metrics,
        "columns": column_metrics,
        "model_total_bytes": sum(x["total_bytes"] for x in tables_metrics),
        "model_rows_sum": sum(x["row_count"] for x in tables_metrics),
    }


def metadata_index(data: Dict[str, Any]) -> Dict[str, Any]:
    tmap = table_name_map(data)
    cmap = column_meta_map(data)

    columns_by_table = defaultdict(list)
    for r in data.get("columns", []):
        tid = s(pick(r, "TableID", "TABLE_ID"))
        name = s(pick(r, "Name", "ExplicitName", "InferredName"))
        columns_by_table[tmap.get(tid, tid)].append(r)

    measures_by_table = defaultdict(list)
    for r in data.get("measures", []):
        tid = s(pick(r, "TableID", "TABLE_ID"))
        measures_by_table[tmap.get(tid, tid)].append(r)

    return {
        "table_by_id": tmap,
        "column_by_id": cmap,
        "columns_by_table": columns_by_table,
        "measures_by_table": measures_by_table,
    }


def relation_label(r: Dict[str, Any], idx: Dict[str, Any]) -> str:
    tmap = idx["table_by_id"]
    cmap = idx["column_by_id"]

    ftid = s(pick(r, "FromTableID"))
    ttid = s(pick(r, "ToTableID"))
    fcid = s(pick(r, "FromColumnID"))
    tcid = s(pick(r, "ToColumnID"))

    ft = tmap.get(ftid, ftid)
    tt = tmap.get(ttid, ttid)
    fc = s(pick(cmap.get(fcid, {}), "Name", "ExplicitName", "InferredName", default=fcid))
    tc = s(pick(cmap.get(tcid, {}), "Name", "ExplicitName", "InferredName", default=tcid))
    return f"{ft}[{fc}] -> {tt}[{tc}]"


def decode_cardinality(v: Any) -> str:
    n = as_int(v, -1)
    return {0: "None", 1: "One", 2: "Many"}.get(n, s(v))


def decode_crossfilter(v: Any) -> str:
    n = as_int(v, -1)
    return {1: "OneDirection", 2: "BothDirections", 3: "Automatic"}.get(n, s(v))


def analyze_model(data: Dict[str, Any], findings: List[Finding], config: Dict[str, Any]) -> Dict[str, Any]:
    idx = metadata_index(data)
    tables = data.get("tables", [])
    columns = data.get("columns", [])
    measures = data.get("measures", [])
    relationships = data.get("relationships", [])

    # Relacionamentos
    for r in relationships:
        label = relation_label(r, idx)
        cross = decode_crossfilter(pick(r, "CrossFilteringBehavior"))
        fc = decode_cardinality(pick(r, "FromCardinality"))
        tc = decode_cardinality(pick(r, "ToCardinality"))

        if cross == "BothDirections":
            add(findings,
                id="MODEL-001", category="Modelo de dados", severity="Alta",
                title="Relacionamento bidirecional",
                object_type="Relacionamento", object_name=label,
                evidence="CrossFilteringBehavior = BothDirections.",
                impact="Pode criar caminhos ambíguos de filtro e elevar o custo de consultas.",
                recommendation="Preferir OneDirection. Manter bidirecional somente quando houver justificativa funcional explícita.",
                confidence="Alta")

        if fc == "Many" and tc == "Many":
            add(findings,
                id="MODEL-002", category="Modelo de dados", severity="Alta",
                title="Relacionamento many-to-many",
                object_type="Relacionamento", object_name=label,
                evidence="FromCardinality = Many e ToCardinality = Many.",
                impact="Pode gerar propagação de filtro complexa, resultados inesperados e pior performance.",
                recommendation="Reavaliar a modelagem. Quando possível, criar dimensão/tabela ponte com chave única e relacionamentos 1:N.",
                confidence="Alta")

    # Calendário - heurística live
    table_names = [s(pick(t, "Name", "ExplicitName", "InferredName")) for t in tables]
    calendar = [x for x in table_names if CALENDAR_NAME_RE.search(x)]
    if not calendar:
        add(findings,
            id="MODEL-003", category="Modelo de dados", severity="Média",
            title="Tabela calendário não identificada",
            object_type="Modelo", object_name="Modelo semântico",
            evidence="Nenhum nome de tabela compatível com padrão de calendário/data foi encontrado.",
            impact="Pode dificultar inteligência de tempo, filtros temporais consistentes e reutilização de medidas.",
            recommendation="Validar se existe dimensão calendário dedicada, contínua e corretamente relacionada.",
            confidence="Média")

    # Fato-fato por nomes
    for r in relationships:
        label = relation_label(r, idx)
        left = label.split("[", 1)[0]
        right = label.split("->", 1)[1].strip().split("[", 1)[0] if "->" in label else ""
        if FACT_NAME_RE.search(left) and FACT_NAME_RE.search(right):
            add(findings,
                id="MODEL-004", category="Modelo de dados", severity="Média",
                title="Possível relacionamento fato-fato",
                object_type="Relacionamento", object_name=label,
                evidence="Os nomes das duas tabelas sugerem tabelas transacionais/fato.",
                impact="Pode indicar granularidades distintas conectadas diretamente e filtros difíceis de controlar.",
                recommendation="Validar granularidade e avaliar dimensões conformadas ou tabela ponte.",
                confidence="Baixa")

    # RI violations do storage engine
    ri_total = 0
    for r in data.get("storage_tables", []):
        ri_total += max(0, as_int(pick(r, "RIVIOLATION_COUNT")))
    if ri_total > 0:
        add(findings,
            id="MODEL-005", category="Modelo de dados", severity="Alta",
            title="Violações de integridade referencial detectadas",
            object_type="Modelo", object_name="Relacionamentos VertiPaq",
            evidence=f"O storage engine reportou {fmt_int(ri_total)} violações de integridade referencial.",
            impact="Valores de chave sem correspondência podem produzir membro em branco e resultados inesperados.",
            recommendation="Localizar chaves órfãs entre fatos e dimensões e corrigir a integridade na origem/ETL.",
            confidence="Alta")

    calc_cols = sum(1 for c in columns if s(pick(c, "Expression")).strip())
    return {
        "tables": len(tables),
        "columns": len(columns),
        "measures": len(measures),
        "relationships": len(relationships),
        "calculated_columns": calc_cols,
    }


def strip_dax_comments(expr: str) -> str:
    expr = re.sub(r"/\*.*?\*/", " ", expr, flags=re.S)
    expr = re.sub(r"//.*?$", " ", expr, flags=re.M)
    expr = re.sub(r"--.*?$", " ", expr, flags=re.M)
    return expr


def dax_items(data: Dict[str, Any]) -> List[Tuple[str, str, str]]:
    idx = metadata_index(data)
    tmap = idx["table_by_id"]
    out = []
    for m in data.get("measures", []):
        tn = tmap.get(s(pick(m, "TableID")), s(pick(m, "TableID")))
        name = s(pick(m, "Name"))
        expr = s(pick(m, "Expression"))
        if expr:
            out.append(("Medida", f"{tn}[{name}]", expr))
    for c in data.get("columns", []):
        expr = s(pick(c, "Expression"))
        if expr:
            tn = tmap.get(s(pick(c, "TableID")), s(pick(c, "TableID")))
            name = s(pick(c, "Name", "ExplicitName", "InferredName"))
            out.append(("Coluna calculada", f"{tn}[{name}]", expr))
    return out


def contains_direct_division(expr: str) -> bool:
    e = strip_dax_comments(expr)
    e = re.sub(r'"(?:[^"]|"")*"', '""', e)
    return "/" in e


def analyze_dax(data: Dict[str, Any], findings: List[Finding], config: Dict[str, Any]) -> None:
    long_limit = int(config["thresholds"].get("dax_long_chars", 2500))
    for typ, obj, expr in dax_items(data):
        upper = strip_dax_comments(expr).upper()

        if "CALCULATE" in upper and "FILTER" in upper:
            add(findings,
                id="DAX-001", category="DAX", severity="Média",
                title="FILTER em expressão com CALCULATE",
                object_type=typ, object_name=obj,
                evidence="CALCULATE e FILTER aparecem na mesma expressão.",
                impact="FILTER pode iterar/materializar mais linhas quando um predicado direto seria suficiente.",
                recommendation="Verificar se FILTER pode ser substituído por filtro booleano direto no CALCULATE. Manter FILTER quando a lógica realmente exigir tabela/iteração.",
                confidence="Média")

        ncalc = len(re.findall(r"\bCALCULATE\s*\(", upper))
        if ncalc >= 2:
            add(findings,
                id="DAX-002", category="DAX", severity="Média",
                title="CALCULATE repetido/aninhado",
                object_type=typ, object_name=obj,
                evidence=f"{ncalc} ocorrências de CALCULATE.",
                impact="Pode indicar transições de contexto redundantes e aumentar complexidade da expressão.",
                recommendation="Revisar se os contextos/filtros podem ser consolidados sem alterar a semântica.",
                confidence="Média")

        if contains_direct_division(expr) and "DIVIDE" not in upper:
            add(findings,
                id="DAX-003", category="DAX", severity="Baixa",
                title="Divisão direta para revisão",
                object_type=typ, object_name=obj,
                evidence="Operador '/' encontrado sem DIVIDE.",
                impact="Pode exigir tratamento manual para zero/BLANK.",
                recommendation="Quando o denominador puder ser zero/BLANK, preferir DIVIDE(numerador, denominador, alternativa).",
                confidence="Média")

        iterators = [
            f for f in ("SUMX", "AVERAGEX", "MINX", "MAXX", "COUNTX", "RANKX", "FILTER")
            if re.search(rf"\b{f}\s*\(", upper)
        ]
        if len(iterators) >= 2:
            add(findings,
                id="DAX-004", category="DAX", severity="Baixa",
                title="Múltiplos iteradores DAX",
                object_type=typ, object_name=obj,
                evidence="Iteradores identificados: " + ", ".join(iterators),
                impact="Iterações sobre grandes tabelas podem elevar o custo no Formula Engine.",
                recommendation="Validar tamanho da tabela iterada, filtros prévios e possibilidade de usar agregações do Storage Engine.",
                confidence="Média")

        if len(expr) >= long_limit:
            add(findings,
                id="DAX-005", category="DAX", severity="Baixa",
                title="Expressão DAX extensa",
                object_type=typ, object_name=obj,
                evidence=f"{len(expr)} caracteres; limite configurado = {long_limit}.",
                impact="Eleva custo de manutenção e pode esconder lógica duplicada.",
                recommendation="Usar VAR, medidas auxiliares reutilizáveis e eliminar cálculos repetidos.",
                confidence="Alta")

        if typ == "Coluna calculada":
            add(findings,
                id="DAX-006", category="DAX", severity="Baixa",
                title="Coluna calculada para revisão",
                object_type=typ, object_name=obj,
                evidence="Expressão calculada armazenada no modelo.",
                impact="Ocupa memória e é processada no refresh.",
                recommendation="Validar se precisa ser coluna; quando possível, preferir medida ou cálculo na origem/Power Query.",
                confidence="Alta")


def analyze_naming(data: Dict[str, Any], findings: List[Finding], config: Dict[str, Any]) -> None:
    idx = metadata_index(data)
    tmap = idx["table_by_id"]

    for t in data.get("tables", []):
        name = s(pick(t, "Name", "ExplicitName", "InferredName"))
        hidden = truth(pick(t, "IsHidden"))
        desc = s(pick(t, "Description"))
        if name and not hidden and not desc:
            add(findings,
                id="NAME-001", category="Nomenclatura e organização", severity="Baixa",
                title="Tabela visível sem descrição",
                object_type="Tabela", object_name=name,
                evidence="Description está vazia.",
                impact="Dificulta manutenção e entendimento da granularidade/papel da tabela.",
                recommendation="Adicionar descrição funcional curta com propósito e granularidade.",
                confidence="Alta")

    for m in data.get("measures", []):
        tn = tmap.get(s(pick(m, "TableID")), s(pick(m, "TableID")))
        name = s(pick(m, "Name"))
        obj = f"{tn}[{name}]"
        if not s(pick(m, "DisplayFolder")).strip():
            add(findings,
                id="NAME-002", category="Nomenclatura e organização", severity="Baixa",
                title="Medida sem pasta de exibição",
                object_type="Medida", object_name=obj,
                evidence="DisplayFolder vazio.",
                impact="Modelos com muitas medidas ficam mais difíceis de navegar.",
                recommendation="Organizar medidas por assunto/área em Display Folders.",
                confidence="Alta")
        if not truth(pick(m, "IsHidden")) and not s(pick(m, "Description")).strip():
            add(findings,
                id="NAME-003", category="Nomenclatura e organização", severity="Baixa",
                title="Medida visível sem descrição",
                object_type="Medida", object_name=obj,
                evidence="Description vazia.",
                impact="A regra de negócio da medida fica implícita.",
                recommendation="Documentar cálculo, unidade e premissas relevantes.",
                confidence="Alta")

    for c in data.get("columns", []):
        tn = tmap.get(s(pick(c, "TableID")), s(pick(c, "TableID")))
        name = s(pick(c, "Name", "ExplicitName", "InferredName"))
        hidden = truth(pick(c, "IsHidden"))
        if name and not hidden and TECHNICAL_NAME_RE.search(name):
            add(findings,
                id="NAME-004", category="Nomenclatura e organização", severity="Média",
                title="Campo técnico visível",
                object_type="Coluna", object_name=f"{tn}[{name}]",
                evidence="Nome aparenta ser ID/chave/código técnico e IsHidden não está ativo.",
                impact="Polui a lista de campos e aumenta o risco de uso incorreto em visuais.",
                recommendation="Ocultar chaves técnicas que não precisam ser consumidas diretamente.",
                confidence="Média")


def analyze_performance(data: Dict[str, Any], findings: List[Finding], config: Dict[str, Any]) -> None:
    th = config["thresholds"]
    metrics = data.get("metrics", {})
    idx = metadata_index(data)

    # Largura de tabelas e calculadas
    for tn, cols in idx["columns_by_table"].items():
        if len(cols) >= int(th["wide_table_columns"]):
            add(findings,
                id="PERF-001", category="Performance / VertiPaq", severity="Média",
                title="Tabela muito larga",
                object_type="Tabela", object_name=tn,
                evidence=f"{len(cols)} colunas; limite configurado = {th['wide_table_columns']}.",
                impact="Mais colunas aumentam memória, custo de refresh e superfície de manutenção.",
                recommendation="Remover atributos sem uso e redundâncias; manter somente colunas necessárias.",
                confidence="Alta")
        calc = [c for c in cols if s(pick(c, "Expression")).strip()]
        if len(calc) >= int(th["calculated_columns_per_table"]):
            add(findings,
                id="PERF-002", category="Performance / VertiPaq", severity="Média",
                title="Muitas colunas calculadas",
                object_type="Tabela", object_name=tn,
                evidence=f"{len(calc)} colunas calculadas; limite = {th['calculated_columns_per_table']}.",
                impact="Colunas calculadas aumentam memória e tempo de processamento.",
                recommendation="Migrar o que for adequado para medida, origem ou Power Query.",
                confidence="Alta")

    # Tabelas pesadas
    for t in metrics.get("tables", []):
        size_mb = mb(t["total_bytes"])
        if size_mb >= float(th["very_large_table_mb"]):
            sev = "Alta"
        elif size_mb >= float(th["large_table_mb"]):
            sev = "Média"
        else:
            continue
        add(findings,
            id="PERF-003", category="Performance / VertiPaq", severity=sev,
            title="Tabela com alto consumo de memória",
            object_type="Tabela", object_name=t["table"],
            evidence=f"VertiPaq estimado: {fmt_mb(t['total_bytes'])} MB; linhas: {fmt_int(t['row_count'])}.",
            impact="Tabelas grandes aumentam memória, tempo de refresh e pressão sobre o Storage Engine.",
            recommendation="Priorizar remoção de colunas, redução de cardinalidade, tipos mais compactos e redução de granularidade quando funcionalmente possível.",
            confidence="Alta", metric=size_mb)

    # Colunas pesadas / dicionários / cardinalidade
    for c in metrics.get("columns", []):
        if not c.get("is_basic_data", False):
            continue
        obj = f"{c.get('table')}[{c.get('column')}]"
        size_mb = mb(c.get("total_bytes"))
        dict_mb = mb(c.get("dictionary_bytes"))

        if size_mb >= float(th["very_large_column_mb"]):
            sev = "Alta"
        elif size_mb >= float(th["large_column_mb"]):
            sev = "Média"
        else:
            sev = None
        if sev:
            add(findings,
                id="PERF-004", category="Performance / VertiPaq", severity=sev,
                title="Coluna com alto consumo VertiPaq",
                object_type="Coluna", object_name=obj,
                evidence=(
                    f"Total estimado: {fmt_mb(c.get('total_bytes'))} MB "
                    f"(dicionário {fmt_mb(c.get('dictionary_bytes'))} MB + "
                    f"segmentos {fmt_mb(c.get('segment_used_bytes'))} MB)."
                ),
                impact="Colunas pesadas têm impacto direto no tamanho do modelo e no custo de scan.",
                recommendation="Verificar se a coluna é necessária; reduzir cardinalidade; substituir texto por chave inteira; separar DateTime; arredondar precisão quando permitido.",
                confidence="Alta", metric=size_mb)

        if dict_mb >= float(th["very_large_dictionary_mb"]):
            sev = "Alta"
        elif dict_mb >= float(th["large_dictionary_mb"]):
            sev = "Média"
        else:
            sev = None
        if sev:
            add(findings,
                id="PERF-005", category="Performance / VertiPaq", severity=sev,
                title="Dicionário de coluna muito grande",
                object_type="Coluna", object_name=obj,
                evidence=f"DICTIONARY_SIZE = {fmt_mb(c.get('dictionary_bytes'))} MB.",
                impact="Dicionário grande costuma indicar muitos valores distintos, textos longos ou ambos.",
                recommendation="Avaliar remoção, normalização, chave substituta inteira, redução de precisão ou agrupamento de valores.",
                confidence="Alta", metric=dict_mb)

        card = c.get("cardinality")
        rows = c.get("row_count") or 0
        ratio = c.get("cardinality_ratio")
        if card is not None and rows >= int(th["high_cardinality_min_rows"]) and ratio is not None:
            if ratio >= float(th["high_cardinality_ratio"]):
                dtype = s(c.get("data_type_physical")).upper()
                sev = "Alta" if ("WSTR" in dtype or "STR" in dtype or "GUID" in dtype) else "Média"
                add(findings,
                    id="PERF-006", category="Performance / VertiPaq", severity=sev,
                    title="Alta cardinalidade",
                    object_type="Coluna", object_name=obj,
                    evidence=(
                        f"Cardinalidade exata = {fmt_int(card)} em {fmt_int(rows)} linhas "
                        f"({ratio*100:.1f}% de valores distintos)."
                    ),
                    impact="Alta cardinalidade reduz compressão e pode aumentar significativamente o tamanho do modelo.",
                    recommendation="Reavaliar necessidade e granularidade; para texto/GUID, preferir chave inteira quando possível; evitar DateTime em precisão desnecessária.",
                    confidence="Alta", metric=ratio)

    # Chaves textuais em relacionamento: cruza modelo lógico com físico
    storage_lookup = {
        (s(c.get("table")).lower(), s(c.get("column")).lower()): c
        for c in metrics.get("columns", [])
    }
    for r in data.get("relationships", []):
        label = relation_label(r, idx)
        # tenta descobrir nomes de origem/destino
        tmap = idx["table_by_id"]
        cmap = idx["column_by_id"]
        pairs = []
        for side in ("From", "To"):
            tid = s(pick(r, f"{side}TableID"))
            cid = s(pick(r, f"{side}ColumnID"))
            tn = tmap.get(tid, tid)
            cn = s(pick(cmap.get(cid, {}), "Name", "ExplicitName", "InferredName", default=cid))
            pairs.append((tn, cn))
        for tn, cn in pairs:
            phys = storage_lookup.get((tn.lower(), cn.lower()))
            if not phys:
                continue
            dtype = s(phys.get("data_type_physical")).upper()
            if "STR" in dtype or "GUID" in dtype:
                add(findings,
                    id="PERF-007", category="Performance / VertiPaq", severity="Média",
                    title="Chave textual/GUID em relacionamento",
                    object_type="Relacionamento", object_name=label,
                    evidence=f"{tn}[{cn}] usa tipo físico {dtype or 'texto/GUID'}.",
                    impact="Chaves largas tendem a consumir mais memória e ser menos eficientes que chaves inteiras.",
                    recommendation="Quando possível, usar surrogate key inteira consistente entre dimensão e fato.",
                    confidence="Alta")
                break


def m_code_items(data: Dict[str, Any]) -> List[Tuple[str, str]]:
    items = []
    for collection_name in ("expressions", "partitions"):
        for r in data.get(collection_name, []):
            name = s(pick(r, "Name", "PartitionName", "ID", default=collection_name))
            candidates = []
            upper = key_upper(r)
            for field in ("EXPRESSION", "QUERY", "SOURCE", "QUERYDEFINITION", "MEXPRESSION"):
                if field in upper and upper[field] not in (None, ""):
                    candidates.append(s(upper[field]))
            # Alguns providers retornam o M serializado em outra coluna textual.
            for code in candidates:
                if "let" in code.lower() or any(x in code for x in ("Sql.Database", "Odbc.", "Excel.Workbook", "Web.Contents", "SharePoint.")):
                    items.append((name, code))
    # dedupe
    seen = set()
    out = []
    for n, c in items:
        key = (n, c)
        if key not in seen:
            seen.add(key)
            out.append((n, c))
    return out


def analyze_m(data: Dict[str, Any], findings: List[Finding], config: Dict[str, Any]) -> Dict[str, Any]:
    queries = m_code_items(data)
    sources = set()

    source_patterns = [
        ("ODBC", r"\bOdbc\.(DataSource|Query)\b"),
        ("SQL Server", r"\bSql\.Database\b"),
        ("Excel", r"\bExcel\.Workbook\b"),
        ("CSV", r"\bCsv\.Document\b"),
        ("SharePoint", r"\bSharePoint\.(Files|Contents)\b"),
        ("Web/API", r"\bWeb\.(Contents|BrowserContents)\b"),
        ("Folder", r"\bFolder\.(Files|Contents)\b"),
        ("Oracle", r"\bOracle\.Database\b"),
        ("PostgreSQL", r"\bPostgreSQL\.Database\b"),
        ("MySQL", r"\bMySQL\.Database\b"),
    ]

    for name, code in queries:
        for label, pattern in source_patterns:
            if re.search(pattern, code, flags=re.I):
                sources.add(label)

        if re.search(r"\bTable\.Buffer\s*\(", code, flags=re.I):
            add(findings,
                id="M-001", category="Consultas M / Fontes", severity="Alta",
                title="Uso de Table.Buffer",
                object_type="Consulta M", object_name=name,
                evidence="Table.Buffer localizado.",
                impact="Pode interromper query folding e forçar materialização local.",
                recommendation="Remover salvo quando teste controlado comprovar benefício; priorizar filtros/projeções na fonte.",
                confidence="Alta")

        if re.search(r"\bOdbc\.Query\s*\(", code, flags=re.I):
            add(findings,
                id="M-002", category="Consultas M / Fontes", severity="Baixa",
                title="SQL nativo via Odbc.Query",
                object_type="Consulta M", object_name=name,
                evidence="Odbc.Query identificado.",
                impact="Pode ser eficiente por pushdown, mas os passos posteriores precisam ser avaliados e o SQL deve estar otimizado.",
                recommendation="Validar índices/plano no banco e incorporar ao SQL transformações pesadas que não façam folding.",
                confidence="Alta")

        p_sort = code.upper().find("TABLE.SORT")
        p_filter = code.upper().find("TABLE.SELECTROWS")
        if p_sort >= 0 and p_filter >= 0 and p_sort < p_filter:
            add(findings,
                id="M-003", category="Consultas M / Fontes", severity="Média",
                title="Ordenação antes de filtro",
                object_type="Consulta M", object_name=name,
                evidence="Table.Sort aparece antes de Table.SelectRows.",
                impact="Pode ordenar um volume maior do que o necessário.",
                recommendation="Filtrar e remover colunas cedo; ordenar somente quando necessário.",
                confidence="Média")

        breaker = []
        for fn in ("Table.Buffer", "Table.AddIndexColumn", "Table.Distinct", "Table.Group"):
            p = code.upper().find(fn.upper())
            if p >= 0:
                breaker.append((p, fn))
        if breaker and p_filter >= 0 and min(x[0] for x in breaker) < p_filter:
            fn = min(breaker)[1]
            add(findings,
                id="M-004", category="Consultas M / Fontes", severity="Média",
                title="Possível quebra de folding antes de filtro",
                object_type="Consulta M", object_name=name,
                evidence=f"{fn} aparece antes de Table.SelectRows.",
                impact="Se a etapa não for delegável, mais dados podem ser transferidos/processados localmente.",
                recommendation="Mover filtros e seleção de colunas para o início e validar View Native Query/Query Folding.",
                confidence="Média")

        if re.search(r"\b(Excel\.Workbook|Csv\.Document)\s*\(", code, flags=re.I):
            add(findings,
                id="M-005", category="Consultas M / Fontes", severity="Baixa",
                title="Fonte baseada em arquivo",
                object_type="Consulta M", object_name=name,
                evidence="Excel/CSV identificado.",
                impact="Arquivos têm menos capacidade de pushdown e podem escalar mal em volume/refresh.",
                recommendation="Para cargas grandes ou críticas, avaliar banco, dataflow, lakehouse ou outra fonte gerenciada.",
                confidence="Alta")

        if re.search(r"\bWeb\.Contents\s*\(", code, flags=re.I):
            add(findings,
                id="M-006", category="Consultas M / Fontes", severity="Baixa",
                title="Fonte Web/API",
                object_type="Consulta M", object_name=name,
                evidence="Web.Contents identificado.",
                impact="APIs podem sofrer paginação, throttling e alto custo por refresh.",
                recommendation="Revisar paginação, parâmetros, cache e redução do payload na origem.",
                confidence="Alta")

    return {"m_queries": len(queries), "sources": sorted(sources)}


def score_findings(findings: List[Finding]) -> Tuple[int, Dict[str, int]]:
    penalty = {c: 0 for c in CATEGORIES}
    for f in findings:
        penalty[f.category] = penalty.get(f.category, 0) + SEVERITY_WEIGHT.get(f.severity, 0)
    cat = {c: max(0, 100 - penalty.get(c, 0)) for c in CATEGORIES}
    return round(sum(cat.values()) / len(cat)), cat


def write_csv(path: Path, rows: List[Dict[str, Any]], fields: List[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k) for k in fields})


def export_csvs(outdir: Path, findings: List[Finding], metrics: Dict[str, Any]) -> None:
    write_csv(
        outdir / "ACHADOS.csv",
        [asdict(x) for x in findings],
        ["id", "category", "severity", "title", "object_type", "object_name",
         "evidence", "impact", "recommendation", "confidence", "metric"]
    )
    write_csv(
        outdir / "VERTIPAQ_TABELAS.csv",
        metrics.get("tables", []),
        ["table", "row_count", "total_bytes", "dictionary_bytes"]
    )
    write_csv(
        outdir / "VERTIPAQ_COLUNAS.csv",
        metrics.get("columns", []),
        ["table", "column", "row_count", "cardinality", "cardinality_ratio",
         "data_type_physical", "dictionary_bytes", "segment_used_bytes",
         "segment_allocated_bytes", "total_bytes", "encoding", "is_key",
         "is_unique", "segments", "bits_max"]
    )


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.drawString(15*mm, 9*mm, f"Power BI Analyzer V2 Deep — v{APP_VERSION}")
    canvas.drawRightString(A4[0]-15*mm, 9*mm, f"Página {doc.page}")
    canvas.restoreState()


def pdf_table(rows, widths, header=True, font=7.5):
    t = Table(rows, colWidths=widths, repeatRows=1 if header else 0)
    commands = [
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,0), (-1,-1), font),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]
    if header:
        commands += [
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ]
    t.setStyle(TableStyle(commands))
    return t


def build_pdf(path: Path, database: str, server: str, data: Dict[str, Any],
              findings: List[Finding], model_stats: Dict[str, Any], m_stats: Dict[str, Any]):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleCenter", parent=styles["Title"], alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="Small2", parent=styles["BodyText"], fontSize=8, leading=10))
    styles.add(ParagraphStyle(name="Tiny", parent=styles["BodyText"], fontSize=7, leading=9))
    styles.add(ParagraphStyle(name="Finding", parent=styles["Heading3"], fontSize=11, leading=13, spaceBefore=6, spaceAfter=3))

    doc = SimpleDocTemplate(
        str(path), pagesize=A4,
        leftMargin=14*mm, rightMargin=14*mm,
        topMargin=15*mm, bottomMargin=15*mm
    )
    story = []

    overall, cat_scores = score_findings(findings)
    sev = defaultdict(int)
    for f in findings:
        sev[f.severity] += 1

    metrics = data.get("metrics", {})
    total_mb = mb(metrics.get("model_total_bytes", 0))

    story += [
        Spacer(1, 18*mm),
        Paragraph("POWER BI ANALYZER", styles["TitleCenter"]),
        Paragraph("Relatório de Boas Práticas — Modo Profundo", styles["Heading2"]),
        Spacer(1, 5*mm),
        Paragraph(
            f"<b>Modelo:</b> {html(database)}<br/>"
            f"<b>Servidor local:</b> {html(server)}<br/>"
            f"<b>Gerado em:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>"
            f"<b>VertiPaq estimado:</b> {total_mb:,.2f} MB".replace(",", "X").replace(".", ",").replace("X", "."),
            styles["BodyText"]
        ),
        Spacer(1, 8*mm),
    ]

    score_rows = [
        ["Score geral", str(overall)],
        ["Críticas", str(sev["Crítica"])],
        ["Altas", str(sev["Alta"])],
        ["Médias", str(sev["Média"])],
        ["Baixas", str(sev["Baixa"])],
        ["Achados", str(len(findings))],
    ]
    story.append(pdf_table(score_rows, [55*mm, 28*mm], header=False, font=9))
    story += [
        Spacer(1, 7*mm),
        Paragraph(
            "O score é um índice de priorização. Um achado não significa automaticamente erro: "
            "o contexto de negócio deve ser validado antes de qualquer mudança.",
            styles["Small2"]
        ),
        PageBreak(),
    ]

    # Executive summary
    story.append(Paragraph("Resumo executivo", styles["Heading1"]))
    summary = [
        ["Indicador", "Valor"],
        ["Tabelas", fmt_int(model_stats.get("tables"))],
        ["Colunas", fmt_int(model_stats.get("columns"))],
        ["Medidas", fmt_int(model_stats.get("measures"))],
        ["Relacionamentos", fmt_int(model_stats.get("relationships"))],
        ["Colunas calculadas", fmt_int(model_stats.get("calculated_columns"))],
        ["Consultas M identificadas", fmt_int(m_stats.get("m_queries"))],
        ["Fontes M", ", ".join(m_stats.get("sources", [])) or "Não identificadas via metadados"],
        ["VertiPaq estimado", f"{total_mb:.2f} MB"],
    ]
    story.append(pdf_table(summary, [65*mm, 105*mm], font=8.5))
    story.append(Spacer(1, 6*mm))

    score_rows = [["Categoria", "Score", "Achados"]]
    for c in CATEGORIES:
        score_rows.append([c, str(cat_scores[c]), str(sum(1 for f in findings if f.category == c))])
    story.append(Paragraph("Score por categoria", styles["Heading2"]))
    story.append(pdf_table(score_rows, [105*mm, 25*mm, 25*mm], font=8))
    story.append(PageBreak())

    # Top tables
    story.append(Paragraph("VertiPaq — tabelas mais pesadas", styles["Heading1"]))
    rows = [["Tabela", "Linhas", "Memória MB", "Dicionário MB"]]
    for t in metrics.get("tables", [])[:20]:
        rows.append([
            html(t["table"]), fmt_int(t["row_count"]),
            fmt_mb(t["total_bytes"]), fmt_mb(t["dictionary_bytes"])
        ])
    if len(rows) > 1:
        story.append(pdf_table(rows, [75*mm, 32*mm, 32*mm, 32*mm], font=7.5))
    else:
        story.append(Paragraph("DMV de storage não retornou métricas de tabelas.", styles["BodyText"]))
    story.append(PageBreak())

    # Top columns
    story.append(Paragraph("VertiPaq — colunas mais pesadas", styles["Heading1"]))
    rows = [["Tabela", "Coluna", "Tipo", "MB", "Dict MB", "Cardinalidade"]]
    for c in metrics.get("columns", [])[:30]:
        if not c.get("is_basic_data", False):
            continue
        rows.append([
            html(c.get("table")), html(c.get("column")),
            html(c.get("data_type_physical")),
            fmt_mb(c.get("total_bytes")), fmt_mb(c.get("dictionary_bytes")),
            fmt_int(c.get("cardinality")) if c.get("cardinality") is not None else "n/c",
        ])
    if len(rows) > 1:
        story.append(pdf_table(rows, [38*mm, 45*mm, 24*mm, 18*mm, 19*mm, 28*mm], font=6.8))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "n/c = cardinalidade não coletada. O V2 calcula cardinalidade exata somente nas colunas mais pesadas, "
        "conforme limite configurável, para evitar varreduras excessivas.",
        styles["Tiny"]
    ))
    story.append(PageBreak())

    # Priorities
    story.append(Paragraph("Prioridades de otimização", styles["Heading1"]))
    priority = sorted(
        findings,
        key=lambda f: (SEVERITY_WEIGHT.get(f.severity, 0), f.metric or 0),
        reverse=True
    )[:25]
    rows = [["Sev.", "Categoria", "Objeto", "Achado"]]
    for f in priority:
        rows.append([
            f.severity, f.category, html(f.object_name[:58]), html(f.title)
        ])
    if len(rows) > 1:
        story.append(pdf_table(rows, [22*mm, 43*mm, 64*mm, 45*mm], font=6.8))
    story.append(PageBreak())

    # Findings by category
    for category in CATEGORIES:
        story.append(Paragraph(category, styles["Heading1"]))
        cat = [f for f in findings if f.category == category]
        if not cat:
            story.append(Paragraph("Nenhum achado pelas regras atuais.", styles["BodyText"]))
            story.append(PageBreak())
            continue
        for f in cat:
            block = [
                Paragraph(f"{html(f.severity)} | {html(f.id)} — {html(f.title)}", styles["Finding"]),
                Paragraph(f"<b>Objeto:</b> {html(f.object_type)} — {html(f.object_name)}", styles["Small2"]),
                Paragraph(f"<b>Evidência:</b> {html(f.evidence)}", styles["Small2"]),
                Paragraph(f"<b>Impacto:</b> {html(f.impact)}", styles["Small2"]),
                Paragraph(f"<b>Recomendação:</b> {html(f.recommendation)}", styles["Small2"]),
                Paragraph(f"<b>Confiança:</b> {html(f.confidence)}", styles["Tiny"]),
                Spacer(1, 2.5*mm),
            ]
            story.append(KeepTogether(block))
        story.append(PageBreak())

    # Warnings
    story.append(Paragraph("Cobertura e avisos técnicos", styles["Heading1"]))
    story.append(Paragraph(
        "O modo profundo consulta apenas metadados, DMVs e pequenas consultas DAX de cardinalidade. "
        "Ele não altera o modelo. Memória de coluna é estimada por DICTIONARY_SIZE + USED_SIZE dos segmentos "
        "de dados básicos do VertiPaq; estruturas auxiliares internas podem não estar integralmente atribuídas "
        "à coluna lógica.",
        styles["BodyText"]
    ))
    warnings = data.get("warnings", [])
    if warnings:
        story.append(Spacer(1, 4*mm))
        story.append(Paragraph("Consultas não disponíveis nesta instância:", styles["Heading3"]))
        for w in warnings[:50]:
            story.append(Paragraph("• " + html(w), styles["Tiny"]))

    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def safe_filename(x: str) -> str:
    x = re.sub(r'[<>:"/\\|?*]+', "_", s(x)).strip(" .")
    return x[:120] or "Modelo"


def main() -> int:
    ap = argparse.ArgumentParser(description="Power BI Analyzer V2 Deep")
    ap.add_argument("--server", required=True, help="Servidor local enviado pelo Power BI, ex.: localhost:12345")
    ap.add_argument("--database", required=True, help="Nome/GUID do modelo enviado pelo Power BI")
    ap.add_argument("--out", help="Diretório de saída")
    ap.add_argument("--config", help="rules_config.json alternativo")
    ap.add_argument("--no-open", action="store_true", help="Não abre o PDF ao terminar")
    args = ap.parse_args()

    cfg = load_config(Path(args.config).resolve() if args.config else None)
    default_out = Path(__file__).resolve().parent / "Resultados"
    base_out = Path(args.out).expanduser().resolve() if args.out else default_out
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = base_out / f"{stamp}_{safe_filename(args.database)}"
    outdir.mkdir(parents=True, exist_ok=True)

    try:
        data = collect_live(args.server, args.database, cfg)

        print("[6/7] Aplicando regras...")
        findings: List[Finding] = []
        model_stats = analyze_model(data, findings, cfg)
        analyze_dax(data, findings, cfg)
        analyze_naming(data, findings, cfg)
        analyze_performance(data, findings, cfg)
        m_stats = analyze_m(data, findings, cfg)

        overall, cat_scores = score_findings(findings)

        result = {
            "analyzer_version": APP_VERSION,
            "mode": "deep_live",
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "server": args.server,
            "database": args.database,
            "score": overall,
            "category_scores": cat_scores,
            "model_stats": model_stats,
            "m_stats": m_stats,
            "metrics": data.get("metrics", {}),
            "warnings": data.get("warnings", []),
            "findings": [asdict(f) for f in findings],
            "rules_config": cfg,
        }

        json_path = outdir / "ANALISE_COMPLETA.json"
        json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        export_csvs(outdir, findings, data.get("metrics", {}))

        print("[7/7] Gerando PDF...")
        pdf_path = outdir / f"PowerBI_Analyzer_{safe_filename(args.database)}.pdf"
        build_pdf(pdf_path, args.database, args.server, data, findings, model_stats, m_stats)

        summary_txt = (
            f"Power BI Analyzer V2 concluído.\n"
            f"Score: {overall}/100\n"
            f"Achados: {len(findings)}\n"
            f"VertiPaq estimado: {fmt_mb(data.get('metrics',{}).get('model_total_bytes',0))} MB\n"
            f"PDF: {pdf_path}\n"
        )
        (outdir / "RESUMO.txt").write_text(summary_txt, encoding="utf-8")
        print(summary_txt)

        if not args.no_open and os.name == "nt":
            try:
                os.startfile(str(pdf_path))
            except Exception:
                pass
        return 0

    except Exception as exc:
        err = outdir / "ERRO_ANALISE.txt"
        err.write_text(
            f"{exc}\n\n{traceback.format_exc()}",
            encoding="utf-8"
        )
        print(f"\nERRO: {exc}\nDetalhes: {err}")
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
