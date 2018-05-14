import numpy as np
import pandas as pd
from datetime import datetime, timedelta

"""Step1: read BOQ_extended to find the material categories and quantity
=================================================================="""
Con_slab={'element ID':str,'Delivery ID':str,'PO ID':str,'WBS code':str,'Family':str,'Type':str,'Family and Type':str,
'Volume':float,'Area':float,'Perimeter':str,'Bauabschnitt':str,'Betoneigenschaften':str,'Core Thickness':str,
	'Estimated Reinforcement Volume':float,'Geschoss':str,	'Manufacturer':str, 'Start_Date': pd.to_datetime,'Finish_Date':pd.to_datetime,'Duration':float}
Con_shearwall={'element ID':str,'Delivery ID':str,'PO ID':str, 'WBS code':str,'Family':str,	'Type':str,	'Family and Type':str,
	'Area':float,	'Length':str,	'Width':str,'Volume':float,	'design load':float,	'Base Constraint':str,
    'Estimated Reinforcement Volume':float,	'Bauabschnitt':str,	'Geschoss':str,	'Manufacturer':str,'Start_Date': pd.to_datetime,'Finish_Date':pd.to_datetime,'Duration':float }
Con_column={'element ID':str,	'Delivery ID':str,	'PO ID':str,	'WBS code':str,	'Family':str,	'Type':str,
	'Family and Type':str,	'design load':int,	'Length':float,	'Volume':float,
    'Estimated Reinforcement Volume':float,	'Bauabschnitt':str,	'Geschoss':str,	'Column Location Mark':str,
    'Manufacturer':str, 'Start_Date': pd.to_datetime,'Finish_Date':pd.to_datetime,'Duration':float}

df_slab_element=pd.read_excel('BIM_QTOs_extended.xlsx',sheetname='Slab_exe',skiprows=None, converters=Con_slab)
df_shearwall_element=pd.read_excel('BIM_QTOs_extended.xlsx',sheetname='CoreShearWall_exe',skiprows=None, converters=Con_shearwall)
df_column_element=pd.read_excel('BIM_QTOs_extended.xlsx',sheetname='Col_exe',skiprows=None, converters=Con_column)

df_slab_po=df_slab_element.copy()
df_shearwall_po=df_shearwall_element.copy()
df_column_po=df_column_element.copy()

df_slab_po['Construction_method']=['insitu' for i in range(df_slab_po.shape[0])]
df_shearwall_po['Construction_method']=['insitu' for i in range(df_shearwall_po.shape[0])]
df_column_po['Construction_method']='prefabrication'

# rebar types and concrete types, for 'insitu'
# http://buildipedia.com/knowledgebase/division-03-concrete/03-20-00-concrete-reinforcing/03-21-00-reinforcing-steel/03-21-00-reinforcing-steel
df_slab_po['Material_Rebar']='ASTM A-615, Type S, Grade 60'
df_shearwall_po['Material_Rebar']='ASTM A-615, Type S, Grade 60'
def Material_Concrete(dfs):
    for i in range(dfs.shape[0]):
        dfs.loc[i,'Material_Concrete']=dfs.loc[i,'Family and Type'][-6:]
    return dfs
df_slab_po=Material_Concrete(df_slab_po)
df_shearwall_po=Material_Concrete(df_shearwall_po)

df_slab_po.rename(columns={'Estimated Reinforcement Volume':'Material_Rebar_Quantity'},inplace=True)
df_slab_po.rename(columns={'Volume':'Material_Concrete_Quantity'}, inplace=True)

df_shearwall_po.rename(columns={'Estimated Reinforcement Volume':'Material_Rebar_Quantity'},inplace=True)
df_shearwall_po.rename(columns={'Volume':'Material_Concrete_Quantity'}, inplace=True)

# prefab types, for 'prefabrication'
df_column_po['Material_assmeblies']=df_column_po['Family and Type']
df_column_po['Material_assmeblies_Quantity']=1


"""Step2: link material with WBS, to drop irrelevant elements, assign PO ID to the same day delivered element
=================================================================="""
df_slab_po.drop(['Start_Date','Finish_Date','Duration','Manufacturer'],axis=1,inplace=True)
df_shearwall_po.drop(['Start_Date','Finish_Date','Duration','Manufacturer'],axis=1,inplace=True)
df_column_po.drop(['Start_Date','Finish_Date','Duration'],axis=1,inplace=True)

Con_schedule={'WBS_code':str,'Start_Date':str,'Finish_Date':pd.to_datetime,'Duration':str}
df_schedule=pd.read_excel('TAKT schedule floor 04OG-10OG.xlsx',sheetname='Task_Table1',skiprows=None, converters=Con_schedule)
df_schedule.rename(columns={'WBS_code':'WBS code'},inplace=True)
df_schedule['Start_Date']=pd.to_datetime(df_schedule['Start_Date']).dt.date
df_schedule['Flow Maturity Index']="{:.1%}".format(1)
df_schedule['Constraints']='No constraints'
df_schedule['Material Status']='transporting'
df_schedule['Coordinator']='Supplier'



df_join_slab=pd.merge(df_slab_po,df_schedule,how='inner',on=['WBS code'])
df_join_shearwall=pd.merge(df_shearwall_po,df_schedule,how='inner',on=['WBS code'])
df_join_column=pd.merge(df_column_po,df_schedule,how='inner',on=['WBS code'])



def PO_ID(dfs,Material_name):
    PO_name = list(dfs.groupby('Start_Date').groups.keys())
    for i in range(dfs.shape[0]):
        for time in range(len(PO_name)):
            if str(dfs.loc[i, 'Start_Date'])[0:10]==str(PO_name[time])[0:10]:
                dfs.loc[i,'PO ID']=str(PO_name[time])+Material_name
            else:
                pass
    return dfs

PO_ID(df_join_slab,'Slab')
PO_ID(df_join_shearwall,'Coreshearwall')
PO_ID(df_join_column,'Column')


"""Step3: Calcualte total quantity of each PO on the same onsite day
=================================================================="""
df_join_slab_PO1=df_join_slab.groupby(['Material_Rebar','PO ID']).agg({'Material_Rebar_Quantity':sum})
df_join_slab_PO2=df_join_slab.groupby(['Material_Concrete','PO ID']).agg({'Material_Concrete_Quantity':sum})

df_join_shearwall_PO1=df_join_shearwall.groupby(['Material_Rebar','PO ID']).agg({'Material_Rebar_Quantity':sum})
df_join_shearwall_PO2=df_join_shearwall.groupby(['Material_Concrete','PO ID']).agg({'Material_Concrete_Quantity':sum})

df_join_column_PO1=df_join_column.groupby(['Material_assmeblies','PO ID']).agg({'Material_assmeblies_Quantity':sum})

df_rebar=pd.concat([df_join_slab_PO1,df_join_shearwall_PO1],keys=['slab','shearwall'])
df_concrete=pd.concat([df_join_slab_PO2,df_join_shearwall_PO2],keys=['slab','shearwall'])
df_prefab_column=df_join_column_PO1
df_rebar.reset_index(level=['PO ID','Material_Rebar'],inplace=True)
df_concrete.reset_index(level=['PO ID','Material_Concrete'],inplace=True)
df_prefab_column.reset_index(level=['PO ID','Material_assmeblies'],inplace=True)

df_rebar.reset_index(inplace=True)
df_rebar.rename(columns={'index':'From Building Element'},inplace=True)
df_concrete.reset_index(inplace=True)
df_concrete.rename(columns={'index':'From Building Element'},inplace=True)
df_prefab_column.reset_index(inplace=True)
df_prefab_column.rename(columns={'index':'From Building Element'},inplace=True)

"""Step4: assign Material Manufacturer to "prefab column" to correspond to BIM element ID
calculate the release date
=================================================================="""
def Manufacturer(dfs,name):
    if 'Manufacturer' in dfs.columns:
        dfs['Manufacturer']=name
    else:
        print('assign manufacturer on POs manually')
    return dfs

Manufacturer(df_join_column,'SACAC AG')
Manufacturer(df_join_shearwall,'N/A')
Manufacturer(df_join_slab,'N/A')

df_rebar['Manufacturer']='REBAR AG'
df_concrete['Manufacturer']='CONCRETE AG'
df_prefab_column['Manufacturer']=np.unique(df_join_column['Manufacturer'].values)[0]

df_rebar['transport_leadtime']=1
df_concrete['transport_leadtime']=1
df_prefab_column['transport_leadtime']=1
df_rebar['manu_leadtime']=2
df_concrete['manu_leadtime']=2
df_prefab_column['manu_leadtime']=2

# calculate the release time for POs
def Releasetime(dfs):
    for i in range(dfs.shape[0]):
        needdate=dfs.loc[i,'PO ID'][0:10]
        dfs.loc[i, 'Material_needdate']=pd.to_datetime(needdate)
        releasedate=dfs.loc[i,'Material_needdate']-timedelta(days=1)
        dfs.loc[i,'Material_releasedate']=datetime.date(releasedate)
        manudate=dfs.loc[i,'Material_needdate']-timedelta(days=2)-timedelta(days=1)
        dfs.loc[i,'Material_manudate']=datetime.date(manudate)
    return dfs
Releasetime(df_rebar)
Releasetime(df_concrete)
Releasetime(df_prefab_column)

def converttodate(dfs):
    dfs.drop(['Material_needdate'],axis=1,inplace=True)
    for i in range(dfs.shape[0]):
        needdate = dfs.loc[i, 'PO ID'][0:10]
        dfs.loc[i, 'Material_needdate'] = (pd.to_datetime(needdate)).date()
    return dfs
converttodate(df_rebar)
converttodate(df_concrete)
converttodate(df_prefab_column)

# baseline delivery schedule, alert action
dumb_duration1=3
def baseline(dfs):
    df_copy=dfs.copy()
    for i in range(df_copy.shape[0]):
        df_copy.loc[i,'Baseline_release']=pd.to_datetime(dfs.loc[i,'Material_needdate'])-timedelta(days=3)
        dfs.loc[i,'Baseline_releasedate']=df_copy.loc[i,'Baseline_release']
        if pd.to_datetime(dfs.loc[i,'Material_releasedate'])<df_copy.loc[i,'Baseline_release']:
            dfs.loc[i,'Alert Action']='Expedite'
        if pd.to_datetime(dfs.loc[i,'Material_releasedate'])>df_copy.loc[i,'Baseline_release']:
            dfs.loc[i, 'Alert Action'] = 'Postpone'
        if pd.to_datetime(dfs.loc[i,'Material_releasedate'])==df_copy.loc[i,'Baseline_release']:
            dfs.loc[i, 'Alert Action'] = 'As planned'
        else:
            pass
    return dfs
baseline(df_rebar)
baseline(df_concrete)
baseline(df_prefab_column)

writer_PO=pd.ExcelWriter('PO_IDs_WBS.xlsx',engine='xlsxwriter')
df_join_slab.to_excel(writer_PO,sheet_name='PO_slab', index=False)
df_join_shearwall.to_excel(writer_PO,sheet_name='PO_coreshearwall', index=False)
df_join_column.to_excel(writer_PO,sheet_name='PO_column', index=False)

df_rebar.to_excel(writer_PO,sheet_name='sum_rebar', index=True)
df_concrete.to_excel(writer_PO,sheet_name='sum_concrete', index=True)
df_prefab_column.to_excel(writer_PO,sheet_name='sum_prefabcol', index=True)

df_schedule.to_excel(writer_PO,sheet_name='Look_Ahead', index=False)

writer_PO.save()

"""Step5: write back to BIM_extend file to be populated into BIM dynamo
=================================================================="""
def writeback_PO(df1,df2, construction_method):
    df1['PO ID']='to be assigned'
    df1['Construct_Method']=construction_method
    for ele in range (df1.shape[0]):
        for i in range(df2.shape[0]):
            if df1.loc[ele,'element ID']==df2.loc[i,'element ID']:
                df1.loc[ele,'PO ID']=df2.loc[i,'PO ID']
            else:
                pass
            if df1.loc[ele, 'Construct_Method'] == 'insitu':
                df1.loc[ele, 'Delivery ID']= 'palleted PO ID'
            else:
                pass
    return df1

df_slab=writeback_PO(df_slab_element,df_join_slab,'insitu')
df_shearwall=writeback_PO(df_shearwall_element,df_join_shearwall, 'insitu')
df_col=writeback_PO(df_column_element,df_join_column, 'prefabrication')

writer_dynamo=pd.ExcelWriter('BIM_Dynamo.xlsx',engine='xlsxwriter')
df_slab.to_excel(writer_dynamo,sheet_name='slab', index=False)
df_shearwall.to_excel(writer_dynamo,sheet_name='shearwall', index=False)
df_col.to_excel(writer_dynamo,sheet_name='column', index=False)
writer_dynamo.save()


