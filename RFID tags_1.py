import numpy as np
import pandas as pd
import uuid
"""Step1: creating the IDs"""
df_col = pd.read_excel('Structural Column Schedule.xlsx', sheetname='Stru.Col',skiprows=None,
                       converters={'element ID':str,'Geschoss':str,'Bauabschnitt':str,'Family and Type':str,
                                   'WBS code':str})
"""1st line in excel file is counted as the header in pandas"""
# print(len(df_col.index))

"""generate rfid tags for each of the items"""
df_col['Delivery ID'] = pd.Series([uuid.uuid4() for i in range(len(df_col))],index=df_col.index)
df_col['Delivery ID'] = df_col['Delivery ID'].apply(lambda x: str(x))


"""generate WBS code for each of the items, avoid chain indexing"""
for i in range(len(df_col.index)):
    df_col.loc[i,'WBS code']="AS_OE_B1_ST_COL"+"_"+df_col.loc[i,'Geschoss']+"_"+df_col.loc[i,'Bauabschnitt']
# for i in range(len(df_col.index)):
#     df_col['WBS code'][i]= "AS_OE_B1_ST_COL"+"_"+df_col['Geschoss'][i]+"_"+df_col['Bauabschnitt'][i]
#
#
""""generate PO ID for each of the items"""
for i in range(len(df_col.index)):
    if "Rechteck Stützen - Vorfabriziert" in str(df_col.loc[i,'Family and Type']):
        df_col.loc[i,'PO ID'] = str(df_col.loc[i,'WBS code'])+"_" + "RSV" + "_" + str(df_col.loc[i,'element ID'])
    if "Rund Stütze" in str(df_col.loc[i,'Family and Type']):
        df_col.loc[i,'PO ID'] = str(df_col.loc[i,'WBS code']) + "_" + "OSV" + "_" + str(df_col.loc[i,'element ID'])
    if "ortbeton" in str(df_col.loc[i,'Family and Type']):
        df_col.loc[i,'PO ID'] = str(df_col.loc[i,'WBS code']) + "_" + "INS" + "_" + str(df_col.loc[i,'element ID'])



df_col.to_excel(pd.ExcelWriter('Structural Column Schedule_all.xlsx',engine='xlsxwriter'), sheet_name='delivery', index=False)



""""Step 2: group product into categories"""
df_col1 = pd.read_excel('Structural Column Schedule_cal.xlsx', sheetname='delivery',skiprows=None,
                       converters={'element ID':str,'Geschoss':str,'Bauabschnitt':str,'Family and Type':str,
                                   'WBS code':str, 'Length':int, 'design load':int})
print(df_col1.shape)

def subtotal(attr):
    numcount = df_col1.groupby(attr).count()
    uniquenumber = df_col1.groupby(attr).nunique()
    uniquename = df_col1.groupby(attr).groups.keys()
    numcount = pd.DataFrame(numcount)
    numcount.to_csv(str(attr)+'.txt', sep=",", index=True, header=True)
    return uniquename, uniquenumber

""""total number of family and types, pair the BIM elements with product categories"""
subtotal(['Family and Type', 'Length'])
# print(type(subtotal(['Family and Type', 'Length'])))

""""total number of family and types to be delivered on the same time"""
subtotal(['Family and Type','WBS code'])

""""check if the total number of BIM elements are the same as the number from the CAD"""
def totaleachfloor(floor):
    totalnum=0
    for i in range(len(df_col1['Geschoss'])):
        if floor == df_col1.loc[i,'Geschoss']:
            totalnum=totalnum+1
        else:
            continue
    return totalnum

floors = ['4UG','3UG','2UG','1UG','00EG','01OG','02OG','03OG','04OG','05OG','06OG','07OG','08OG','09OG','10OG','11OG','12OG'
          ,'13OG','14OG','15OG','16OG','17OG','18OG','19OG','20OG','21OG']
listsumCAD=[30,54, 31, 46, 43, 35, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27]
listsum = []
for i in range(len(floors)):
    a = totaleachfloor(floors[i])
    listsum.append(a)
missingcol = [m-n for m,n in zip(listsumCAD,listsum)]


