from pathlib import Path
import json
import powerbi_analyzer_v2 as a

sample = {
 "tables":[{"ID":1,"Name":"F_Vendas"},{"ID":2,"Name":"D_Produto"}],
 "columns":[
   {"ID":11,"TableID":1,"Name":"ProdutoId","IsHidden":False},
   {"ID":12,"TableID":1,"Name":"Valor","IsHidden":False},
   {"ID":21,"TableID":2,"Name":"ProdutoId","IsHidden":False},
   {"ID":22,"TableID":2,"Name":"Produto","IsHidden":False}
 ],
 "measures":[{"TableID":1,"Name":"Receita","Expression":"SUM(F_Vendas[Valor])"}],
 "relationships":[{"FromTableID":1,"FromColumnID":11,"ToTableID":2,"ToColumnID":21,
                   "FromCardinality":2,"ToCardinality":1,"CrossFilteringBehavior":1}],
 "partitions":[], "expressions":[], "data_sources":[], "calc_dependency":[],
 "storage_tables":[
   {"DIMENSION_NAME":"F_Vendas","TABLE_ID":"F_Vendas","PARTITION_NAME":"P","ROWS_COUNT":1000000,"RIVIOLATION_COUNT":0},
   {"DIMENSION_NAME":"D_Produto","TABLE_ID":"D_Produto","PARTITION_NAME":"P","ROWS_COUNT":10000,"RIVIOLATION_COUNT":0}
 ],
 "storage_columns":[
   {"DIMENSION_NAME":"F_Vendas","ATTRIBUTE_NAME":"ProdutoId","COLUMN_ID":"ProdutoId","COLUMN_TYPE":"BASIC_DATA","DATATYPE":"DBTYPE_I8","DICTIONARY_SIZE":10000},
   {"DIMENSION_NAME":"F_Vendas","ATTRIBUTE_NAME":"Valor","COLUMN_ID":"Valor","COLUMN_TYPE":"BASIC_DATA","DATATYPE":"DBTYPE_R8","DICTIONARY_SIZE":1000},
   {"DIMENSION_NAME":"D_Produto","ATTRIBUTE_NAME":"Produto","COLUMN_ID":"Produto","COLUMN_TYPE":"BASIC_DATA","DATATYPE":"DBTYPE_WSTR","DICTIONARY_SIZE":500000}
 ],
 "storage_segments":[
   {"DIMENSION_NAME":"F_Vendas","COLUMN_ID":"ProdutoId","USED_SIZE":2000000,"ALLOCATED_SIZE":2200000,"RECORDS_COUNT":1000000,"BITS_COUNT":14},
   {"DIMENSION_NAME":"F_Vendas","COLUMN_ID":"Valor","USED_SIZE":5000000,"ALLOCATED_SIZE":5200000,"RECORDS_COUNT":1000000,"BITS_COUNT":32},
   {"DIMENSION_NAME":"D_Produto","COLUMN_ID":"Produto","USED_SIZE":300000,"ALLOCATED_SIZE":350000,"RECORDS_COUNT":10000,"BITS_COUNT":14}
 ],
 "warnings":[]
}
sample["metrics"] = a.build_deep_metrics(sample)
cfg = a.load_config(Path("rules_config.json"))
findings=[]
stats=a.analyze_model(sample,findings,cfg)
a.analyze_dax(sample,findings,cfg)
a.analyze_naming(sample,findings,cfg)
a.analyze_performance(sample,findings,cfg)
mstats=a.analyze_m(sample,findings,cfg)
print("SELFTEST OK")
print("tables", len(sample["metrics"]["tables"]))
print("columns", len(sample["metrics"]["columns"]))
print("findings", len(findings))
