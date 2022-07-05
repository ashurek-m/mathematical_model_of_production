import pandas as pd

df_rest = pd.read_excel('restrictions.xlsx')
df_rest.info()
print(df_rest.loc[0, 'delta_smen_hours'])
