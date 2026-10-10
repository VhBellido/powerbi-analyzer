# Arquitetura do V2

```text
Power BI Desktop
      |
      | External Tools
      | %server% / %database%
      v
Power BI Analyzer V2
      |
      +--> TMSCHEMA_TABLES
      +--> TMSCHEMA_COLUMNS
      +--> TMSCHEMA_MEASURES
      +--> TMSCHEMA_RELATIONSHIPS
      +--> TMSCHEMA_PARTITIONS
      +--> TMSCHEMA_EXPRESSIONS
      +--> DISCOVER_CALC_DEPENDENCY
      |
      +--> DISCOVER_STORAGE_TABLES
      |       +--> ROWS_COUNT
      |       +--> RIVIOLATION_COUNT
      |
      +--> DISCOVER_STORAGE_TABLE_COLUMNS
      |       +--> DICTIONARY_SIZE
      |       +--> COLUMN_ENCODING
      |       +--> DATATYPE
      |
      +--> DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS
              +--> USED_SIZE
              +--> ALLOCATED_SIZE
              +--> RECORDS_COUNT
              +--> BITS_COUNT
              +--> SEGMENTS
      |
      +--> DAX cardinality probes (top N)
      |
      v
Rule Engine
      |
      +--> Modelo
      +--> DAX
      +--> Naming
      +--> VertiPaq
      +--> Power Query
      |
      v
PDF + JSON + CSV
```

## Próximas evoluções possíveis

- parser DAX real em vez de regras regex;
- dependências de medidas e objetos não usados;
- captura de Server Timings;
- Benchmark de medidas;
- geração de scripts TMDL/Tabular Editor para correções assistidas;
- histórico dos scores por relatório;
- painel consolidado com todos os PBIX analisados;
- regras corporativas por área/modelo.
