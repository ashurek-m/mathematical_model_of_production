from pulp import LpMaximize, LpProblem, LpStatus, lpSum, LpVariable
import pandas as pd


def save_excel(path, data_frame, columns=True):
    writer_on = pd.ExcelWriter(path, engine='xlsxwriter')
    data_frame.to_excel(writer_on, index=False, header=columns)
    writer_on.save()


# Подготовка исходных данных для расчетов
df = pd.read_excel('entry.xlsx', sheet_name='model')
df_2 = df.dropna(subset='наименование', axis=0)
df_2.info()
# величина смены в часах
b = 8 * 0.8
# кол-во смен в месяце
c = 6
# Создание модели
model = LpProblem('prod', LpMaximize)
# Генерация словарей с значениями переменных
names_det = list(df_2['item'])
costs = dict(zip(names_det, df_2['стоимость']))
cn = dict(zip(names_det, df_2['cn']))
cn_4 = dict(zip(names_det, df_2['4_ось']))
dart = dict(zip(names_det, df_2['dart']))
tour = dict(zip(names_det, df_2['tour']))

vat = LpVariable.dicts('item', names_det, cat='Binary')

costs_1 = []
for n in vat:
    costs_1.append(n)
cost_s1 = pd.Series(data=costs_1, name='item')
cost_df = cost_s1.to_frame()
cost_df.info()
cost_2 = []
for k in vat:
    cost_2.append(vat[k])
cost_df = cost_df.assign(re_name=cost_2)

model += lpSum(costs[i] * vat[i] for i in names_det)
t_cn = 1600
t_dart = 765
t_tour = 1080
'''
количество центров 18шт
количество дартов 16шт
количество токарных 9 шт
'''
count_cn = 16
count_dart = 14
count_tour = 8
model += lpSum(costs[i] * vat[i] for i in names_det)
model += lpSum(cn[i] * vat[i] for i in names_det) <= t_cn
model += lpSum(dart[i] * vat[i] for i in names_det) <= t_dart
model += lpSum(tour[i] * vat[i] for i in names_det) <= t_tour
for i in names_det:
    model += ((cn[i] + dart[i] + tour[i]) * vat[i]) <= c * b
model += lpSum(cn[i] * vat[i] for i in names_det) <= c * b * count_cn
model += lpSum(dart[i] * vat[i] for i in names_det) <= c * b * count_dart
model += lpSum(tour[i] * vat[i] for i in names_det) <= c * b * count_tour
model.solve()
var_list = []
print("Status:", LpStatus[model.status])

for v in model.variables():
    var_list_2 = [v.name, v.varValue]
    var_list.append(var_list_2)

df_var = pd.DataFrame(data=var_list, columns=['re_name', 'var'])
# print(df_var.head())
# print(cost_df.head())
#df_out = df_var.merge(cost_df, on='re_name', how='outer')
df_out = pd.concat([df_var, cost_df], axis=1)

save_excel('result.xlsx', df_out)

'''https://proglib.io/p/lineynoe-programmirovanie-praktika-resheniya-zadach-optimizacii-na-python-2020-11-26 
https://www.machinelearningmastery.ru/linear-programming-and-discrete-optimization-with-python-using-pulp
-449f3c5f6e99/ '''
