# Power BI Analyzer

Auditoria automática de modelos Power BI, direto do Power BI Desktop: um botão em **Ferramentas Externas** lê o modelo aberto (somente leitura) e gera um relatório em PDF com o que melhorar em DAX, organização, memória (VertiPaq) e Power Query.

**Por que existe:** quem mantém muitos painéis não consegue revisar cada modelo à mão. Medidas sem pasta, `FILTER` dentro de `CALCULATE`, colunas com cardinalidade alta e tabelas pesadas passam despercebidas até o relatório ficar lento.

```
Power BI Desktop ──(%server% / %database%)──▶ Analyzer ──▶ DMVs + TMSCHEMA ──▶ regras ──▶ PDF + JSON + CSV
```

| O que ele verifica | Exemplos |
| --- | --- |
| Modelo | relacionamentos, cardinalidade, direção de filtro, colunas calculadas, integridade referencial |
| DAX | `FILTER` com `CALCULATE`, `CALCULATE` repetido, divisão com `/`, iteradores, expressões longas |
| Organização | medidas sem Display Folder, objetos sem descrição, campos técnicos visíveis |
| VertiPaq | memória por tabela e coluna, dicionários grandes, cardinalidade exata das colunas mais pesadas |
| Power Query | fontes, `Table.Buffer`, ordenação antes do filtro, possíveis quebras de folding |

Stack: Python · ADO/MSOLAP · DMVs do Analysis Services · ReportLab · PowerShell

---

## Detalhes da V2


## O que mudou no V2

O V2 não depende de "adivinhar" a porta do Power BI Desktop.

Ele é registrado como **Ferramenta Externa**. Quando você clica no botão dentro do Power BI Desktop, o próprio Power BI envia:

- `%server%` — instância local do Analysis Services + porta;
- `%database%` — nome/GUID do modelo atual.

A ferramenta conecta em modo somente leitura usando o provider **MSOLAP**, coleta TMSCHEMA + DMVs e gera a documentação.

---

# Instalação

## 1. Pré-requisitos

- Windows 10/11 64 bits.
- Power BI Desktop.
- Python 3.11+ 64 bits.
- Internet somente para instalar os pacotes Python na primeira execução.

O Power BI Desktop instala as bibliotecas cliente do Analysis Services, incluindo o provider MSOLAP utilizado pelo Analyzer.

## 2. Execute

Clique com botão direito em:

`INSTALAR_V2.ps1`

e escolha **Executar com PowerShell**.

O instalador:

1. cria `.venv`;
2. instala `pywin32` e `reportlab`;
3. testa ADO;
4. pede elevação de administrador apenas para registrar o botão em External Tools.

## 3. Reinicie o Power BI Desktop

Abra um PBIX com modelo local e clique:

**Ferramentas Externas > Power BI Analyzer V2**

---

# Saída

O Analyzer cria:

`Resultados\AAAAmmdd_HHMMSS_<modelo>\`

Arquivos:

- `PowerBI_Analyzer_<modelo>.pdf` — relatório executivo/técnico;
- `ANALISE_COMPLETA.json` — achados + métricas;
- `ACHADOS.csv`;
- `VERTIPAQ_TABELAS.csv`;
- `VERTIPAQ_COLUNAS.csv`;
- `RESUMO.txt`.

O PDF abre automaticamente ao terminar.

---

# O que o modo profundo coleta

## Modelo
- tabelas;
- colunas;
- medidas;
- relacionamentos;
- cardinalidade dos relacionamentos;
- direção de filtro;
- colunas calculadas;
- metadados de partições/fontes;
- violações de integridade referencial reportadas pelo storage engine.

## DAX
- `FILTER` com `CALCULATE`;
- `CALCULATE` repetido;
- divisão direta `/`;
- iteradores;
- expressões extensas;
- colunas calculadas.

## Organização
- medidas sem Display Folder;
- medidas/tabelas sem descrição;
- campos técnicos visíveis.

## VertiPaq / performance
- row count;
- memória por tabela;
- memória por coluna;
- `DICTIONARY_SIZE`;
- `USED_SIZE`;
- `ALLOCATED_SIZE`;
- quantidade de segmentos;
- bits por segmento;
- cardinalidade exata das colunas mais pesadas;
- razão cardinalidade/linhas;
- tabelas largas;
- tabelas/colunas muito pesadas;
- dicionários grandes;
- chaves textuais/GUID em relacionamentos.

## Power Query / fontes
Quando as expressões M estão expostas nos metadados do modelo:
- ODBC;
- SQL Server;
- Excel/CSV;
- SharePoint;
- Web/API;
- `Table.Buffer`;
- ordenação antes do filtro;
- possíveis quebras de folding antes de filtros.

---

# Cardinalidade exata

Calcular `DISTINCTCOUNT` de toda coluna de um modelo grande seria caro.

Por padrão o V2 calcula cardinalidade exata somente para as **25 colunas físicas mais pesadas**.

Você pode alterar em `rules_config.json`:

```json
"exact_cardinality_top_n": 25
```

Se colocar `0`, desativa as consultas de cardinalidade exata.

---

# Thresholds

Edite `rules_config.json` para adaptar ao padrão da sua equipe:

- tabela larga;
- quantidade de colunas calculadas;
- tamanho de coluna;
- tamanho de dicionário;
- tamanho de tabela;
- percentual de alta cardinalidade;
- quantidade de colunas para cardinalidade exata.

---

# Segurança

O Analyzer V2:

- abre conexão somente para leitura;
- executa DMVs/TMSCHEMA;
- executa consultas DAX `EVALUATE` somente para cardinalidade;
- não executa `ALTER`, `CREATE`, `DELETE`, `UPDATE` ou comandos de alteração;
- não modifica o PBIX.

---

# Observações

## Relatórios thin / live connection
Se o PBIX for apenas um relatório fino conectado a um modelo remoto, a instância local pode não conter o modelo importado/VertiPaq completo. Nesse caso algumas DMVs podem retornar vazias ou indisponíveis. O relatório PDF registra esses avisos.

## DirectQuery
Tabelas puramente DirectQuery podem não ter as mesmas métricas físicas de VertiPaq de uma tabela Import. O Analyzer trabalha com o que o engine expõe.

## Memória
A memória atribuída a uma coluna é estimada como:

`DICTIONARY_SIZE + soma(USED_SIZE dos segmentos físicos de BASIC_DATA)`

Estruturas auxiliares internas (hierarquias/relacionamentos) não são necessariamente atribuídas integralmente a uma coluna lógica.

---

# Uso manual

Se você souber `server` e `database`:

```powershell
.\.venv\Scripts\python.exe .\powerbi_analyzer_v2.py `
  --server "localhost:51234" `
  --database "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

Também existe `ANALISAR_MANUAL.bat`.

---

# Desinstalar botão do Power BI

Execute:

`DESINSTALAR_FERRAMENTA_EXTERNA.ps1`

A pasta do Analyzer pode ser apagada depois, se desejado.
