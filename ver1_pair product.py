import numpy as np
import pandas as pd


""""Step 1: define product category from supplier"""
df_col1 = pd.read_excel('Structural Column Schedule_cal.xlsx', sheetname='delivery',skiprows=None,
                       converters={'element ID':str,'Geschoss':str,'Bauabschnitt':str,'Family and Type':str,
                                   'WBS code':str, 'Length':int, 'design load':int})
df_col2 = pd.read_excel('ordering subtotal.xlsx', sheetname='supplier product',skiprows=None,
                       converters={'Col Type':str,'Profile':str,'Length':int, 'Nd':int, 'total':int, 'estimate total':int})

len1=len(df_col1.index)
len2=len(df_col2.index)
print(len1,len2)

totallist=[]
for num in range(len2):
    totallist.append(0)

# print(str(df_col2.loc[1, 'Profile']) in str(df_col1.loc[1,'Family and Type']))
# print(df_col2['Nd'][2] in range(int(df_col1['design load'][23]*0.9),int(df_col1['design load'][23]*1.1)))

for col in range(len1):
    for item in range(len2):
        if str(df_col2['Profile'][item]) in str(df_col1['Family and Type'][col]) \
                and df_col1['design load'][col] in range(int(df_col2['Nd'][item]*0.85),int(df_col2['Nd'][item]*1.001))\
                and df_col1['Length'][col] in range(int(df_col2['Length'][item]*0.85),int(df_col2['Length'][item]*1.001)):
            totallist[item] = totallist[item]+1
        else:
            continue

listdiff = [n-m for n,m in zip(totallist,df_col2['estimate total'])]

dftotal = pd.DataFrame({'total':totallist, 'Type NO.': df_col2['Type NO.'], 'Implenia': df_col2['estimate total'], 'difference':listdiff})
writer = pd.ExcelWriter('PO num.xlsx')
dftotal.to_excel(writer,sheet_name='supplier product',startcol=0,startrow=0, index=False)

writer.save

print(sum(totallist))
