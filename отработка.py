import pandas as pd

df = pd.read_excel('entry.xlsx', sheet_name='model')
df = df.dropna(subset='наименование', axis=0)
df['item'] = df['item'].astype('int64')
df.info()
