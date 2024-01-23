import pandas as pd

df = pd.read_fwf('SQL/TESTE.rpt', delimeter='|')
print(df)