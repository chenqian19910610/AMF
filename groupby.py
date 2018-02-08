import numpy as np
import pandas as pd


# """"wall schedule"""
# df_wal = pd.read_excel('walls schedule.xlsx', sheetname='walls',skiprows=None,
#                        converters={'WBS code':str,'Geschoss':str,'Type':str,'Width_m':float,
#                                    'Volume_m3':float, 'Length_m':float, 'height':float})
# df_wal14 = df_wal.loc[df_wal['Geschoss']== '14OG']
#
# print(type(df_wal14))
#
# def subtotal(attr):
#     numcount = df_wal14.groupby(attr).count()
#     subtotals = df_wal14.groupby(attr).agg({'Volume_m3': sum}).dropna()
#     uniquenumber = df_wal14.groupby(attr).nunique()
#     uniquename = df_wal14.groupby(attr).groups.keys()
#     # numcount = pd.DataFrame(numcount)
#     # numcount.to_csv(str(attr)+'.txt', sep=",", index=True, header=True)
#     return subtotals
#
# a = subtotal(['Type'])
#
# print(a, type(a))

# """"prefab col schedule"""
# df_col = pd.read_excel('Structural Column Schedule IDs LOD 400.xlsx', sheetname='information',skiprows=None,
#                        converters={'WBS code':str,'Geschoss':str,'Family and Type':str, 'design load':int, 'Length':int})
# df_col14 = df_col.loc[df_col['Geschoss']== '14OG']
#
# print(type(df_col14))
#
# def subtotal(attr):
#     numcount = df_col14.groupby(attr).count()
#     uniquenumber = df_col14.groupby(attr).nunique()
#     uniquename = df_col14.groupby(attr).groups.keys()
#     # numcount = pd.DataFrame(numcount)
#     # numcount.to_csv(str(attr)+'.txt', sep=",", index=True, header=True)
#     return numcount
#
# b = pd.DataFrame(subtotal(['Family and Type']))
# # b.to_csv('a.csv',sep=',')
#
# print(b, type(b))

""""prefab col schedule"""
df_sla = pd.read_excel('slabs schedule.xlsx', sheetname='slabs',skiprows=None,
                       converters={'WBS code':str,'Geschoss':str,'Type':str,'Volume_m3':float,
                                   'Area_m2':float, 'height':float, 'Perimeter_m':float})
df_sla14 = df_sla.loc[df_sla['Geschoss']== '14OG']

print(type(df_sla14))

def subtotal(attr):
    numcount = df_sla14.groupby(attr).count()
    subtotals = df_sla14.groupby(attr).agg({'Volume_m3': sum}).dropna()
    uniquenumber = df_sla14.groupby(attr).nunique()
    uniquename = df_sla14.groupby(attr).groups.keys()
    # numcount = pd.DataFrame(numcount)
    # numcount.to_csv(str(attr)+'.txt', sep=",", index=True, header=True)
    return numcount,subtotals

c = subtotal(['Type'])

print(c, type(c))

