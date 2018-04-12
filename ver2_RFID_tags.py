import numpy as np
import pandas as pd
import uuid

"""Step1: read excel data
=================================================================="""
Con_slab={'element ID':str,'Delivery ID':str,'PO ID':str,'WBS code':str,'Family':str,'Type':str,'Family and Type':str,
'Volume':str,'Area':str,'Perimeter':str,'Bauabschnitt':str,'Betoneigenschaften':str,'Core Thickness':str,
	'Estimated Reinforcement Volume':str,'Geschoss':str,	'Manufacturer':str}
Con_shearwall={'element ID':str,'Delivery ID':str,'PO ID':str, 'WBS code':str,'Family':str,	'Type':str,	'Family and Type':str,
	'Area':str,	'Length':str,	'Width':str,'Volume':str,	'design load':str,	'Base Constraint':str,
    'Estimated Reinforcement Volume':str,	'Bauabschnitt':str,	'Geschoss':str,	'Manufacturer':str}
Con_column={'element ID':str,	'Delivery ID':str,	'PO ID':str,	'WBS code':str,	'Family':str,	'Type':str,
	'Family and Type':str,	'design load':str,	'Length':str,	'Volume':str,
    'Estimated Reinforcement Volume':str,	'Bauabschnitt':str,	'Geschoss':str,	'Column Location Mark':str,
    'Manufacturer':str}

def df_generator(sheet_name,converter):
	dfs=pd.read_excel('QTOs.xlsx', sheetname=sheet_name,skiprows=None,
	                       converters=converter)
	return dfs

df_slab = df_generator('Slab',Con_slab)
df_shearwall = df_generator('ShearCoreWall',Con_shearwall)
df_col = df_generator('Column',Con_column)

"""read schedule from excel file that is exported from .mpp file using Filter1 in MS project"""
Con_schedule={'WBS_code':str,'Start_Date':pd.to_datetime,'Finish_Date':pd.to_datetime,'Duration':str}
df_schedule = pd.read_excel('TAKT schedule floor 04OG-10OG.xlsx',sheetname='Task_Table1',skiprows=None,converters=Con_schedule)
# print(df_schedule)

"""1st line in excel file is counted as the header in pandas"""
# print(len(df_slab.index), df_slab.shape)
# print(len(df_shearwall.index), df_shearwall.shape)
# print(len(df_col.index), df_col.shape)

def convertonum(dfs, attr, a):
	for i in range(dfs.shape[0]):
		if isinstance(dfs.loc[i,attr],float):
			pass
		else:
			dfs.loc[i,attr]=float(dfs.loc[i,attr][:a].strip().replace(' ',''))
convertonum(df_slab,'Volume',-3)
convertonum(df_shearwall,'Volume',-3)
convertonum(df_col,'Volume',-3)
convertonum(df_slab,'Area',-3)
convertonum(df_shearwall,'Area',-3)
convertonum(df_col,'Length',-2)

"""Step 2:generate rfid tags for each element
===================================================================="""
def rfid_generator(dfs):
	dfs['Delivery ID'] = pd.Series([uuid.uuid4() for i in range(len(dfs))],index=dfs.index)
	dfs['Delivery ID'] = dfs['Delivery ID'].apply(lambda x: str(x))
	return dfs
rfid_generator(df_slab)
rfid_generator(df_shearwall)
rfid_generator(df_col)


"""Step3: pair the element with the task start-finish dates via WBS code
========================================================================="""
def pair(dfs,df2):
	for ele in range(dfs.shape[0]):
		for i in range(df2.shape[0]):
			if dfs.loc[ele,'WBS code']==df2.loc[i,'WBS_code']:
				dfs.loc[ele,'Start_Date']=df2.loc[i,'Start_Date']
				dfs.loc[ele,'Finish_Date']=df2.loc[i,'Finish_Date']
				dfs.loc[ele,'Duration']=float(df2.loc[i,'Duration'].replace(' dys','').replace(' dy',''))
			else:
				pass
	return dfs
pair(df_slab,df_schedule)
pair(df_shearwall,df_schedule)
pair(df_col,df_schedule)

# print(len(df_slab.Start_Date.isnull()))
"""Step4: find subtotal and combine, generate the PO ID for the order that happen on the same day
========================================================================="""
# scenario 1: insitu, rebar is proportional to days, concrete is one-day delviery for the whole floor
def rebarQ(dfs,attr1,attr2,rebarcoef):
	dfs[attr2]=dfs[attr1]*rebarcoef/1000
	#steel reinforcement factor from SIA 264,  (slab:120kg/cum, column:150kg/cum, concretewall: 150kg/cum, beam: 110kg/cum)
	return dfs
rebarQ(df_slab,'Volume','Estimated Reinforcement Volume', 120)
rebarQ(df_shearwall,'Volume','Estimated Reinforcement Volume', 150)

from pandas import groupby
def subtotal(dfs, groupbystart, groupbyfinish, aggattr1, aggattr2):
	for i in range(dfs.shape[0]):
		if dfs.loc[i,'Start_Date'] is None:
			pass
		else:
			df_group=dfs.groupby([groupbystart,groupbyfinish]).agg({aggattr1:sum,aggattr2:sum})
			df_group.reset_index(inplace=True)
	return df_group
slab_group=subtotal(df_slab,'Start_Date','Finish_Date','Estimated Reinforcement Volume', 'Volume')
shearwall_group=subtotal(df_shearwall,'Start_Date','Finish_Date','Estimated Reinforcement Volume', 'Volume')


# scenario 2: prefab, all the elements are delivered proportional to days
def subtotal(dfs, groupbystart, groupbyfinish):
	for i in range(dfs.shape[0]):
		if dfs.loc[i,'Start_Date'] is None:
			pass
		else:
			df_group=dfs.groupby([groupbystart,groupbyfinish]).size().reset_index(name='counts')
	return df_group
column_group=subtotal(df_col,'Start_Date','Finish_Date')

# print(df_col.loc[df_col['Start_Date']==column_group.loc[0,'Start_Date']])

# pieces={'s':slab_group,'w':shearwall_group,'c':column_group}
# total_the_day=pd.concat(pieces)

total_the_day=pd.concat([column_group,shearwall_group,slab_group],join='outer',join_axes=None)
new_index=[i for i in range(total_the_day.shape[0])]
total_the_day.index=new_index

total_the_day['duration']=total_the_day['Finish_Date']-total_the_day['Start_Date']
total_the_day['duration']=pd.to_timedelta(total_the_day.duration).dt.days
for i in range(total_the_day.shape[0]):
	if total_the_day.loc[i,'duration']==0:
		total_the_day.loc[i,'duration']=1

total_the_day['Volume_amountperday']=0
total_the_day['rebar_amountperday']=0
total_the_day['prefcol_amountperday']=0
for i in range(total_the_day.shape[0]):
	if total_the_day.loc[i,'Volume'] is None:
		pass
	else:
		total_the_day.loc[i,'Volume_amountperday']=total_the_day.loc[i,'Volume']/total_the_day.loc[i,'duration']

	if total_the_day.loc[i, 'Estimated Reinforcement Volume'] is None:
		pass
	else:
		total_the_day.loc[i, 'rebar_amountperday'] = total_the_day.loc[i, 'Estimated Reinforcement Volume'] / total_the_day.loc[i, 'duration']

	if pd.isnull(total_the_day.loc[i,'counts']):
		pass
	else:
		total_the_day.loc[i,'prefcol_amountperday']=int(total_the_day.loc[i,'counts']/total_the_day.loc[i,'duration'])+1

total_the_day['WBS_code'] = 0
for i in range(df_schedule.shape[0]):
	for task in range(total_the_day.shape[0]):
		if total_the_day.loc[task,'Start_Date']==df_schedule.loc[i,'Start_Date'] and total_the_day.loc[task,'Finish_Date']==df_schedule.loc[i,'Finish_Date']:
			total_the_day.loc[task,'WBS_code']=df_schedule.loc[i,'WBS_code']
		else:
			pass


writer_total = pd.ExcelWriter('summarize_days.xlsx', engine='xlsxwriter')
total_the_day.to_excel(writer_total, sheet_name='total_theday', index=True)
writer_total.save()

# material concrete volume every day in a year, assume PO order ID is based on daily needs, total=365 IDs
from datetime import date,timedelta
# schedu_start_date=min(total_the_day['Start_Date'])
# schedu_finish_date=max(total_the_day['Finish_Date'])
# print(schedu_finish_date-schedu_start_date)

from datetime import timedelta,date
# the total number of different task (WBS) and their start-finish dates, different task may need the same material
task_collect=total_the_day.index
start_date=date(2017,1,1)
eachday=[0 for i in range(365)]
for i in range(365):
    eachday[i]=str(start_date+timedelta(i))

df_PO = pd.DataFrame(np.zeros((len(task_collect),len(eachday))),index=task_collect, columns=eachday)
df_PO.columns = df_PO.columns.astype(str)


def PO_material_type(material_name):
	for WBS in range(len(task_collect)):
		for day in range(len(eachday)):
			eachdaytime=pd.to_datetime(eachday[day])
			if eachdaytime<=total_the_day.loc[WBS,'Finish_Date'] and eachdaytime>=total_the_day.loc[WBS,'Start_Date']:
				df_PO.loc[WBS,eachday[day]]=total_the_day.loc[WBS,material_name]
			else:
				df_PO.loc[WBS, eachday[day]] = 0
	return df_PO
df_PO_concrete=PO_material_type('Volume_amountperday')
df_PO_rebar=PO_material_type('rebar_amountperday')
df_PO_col=PO_material_type('prefcol_amountperday')

# df_PO_concrete_twoproject=df_PO_concrete.add(df_PO_concrete,fill_value=0)
# print(df_PO_concrete_twoproject,df_PO_concrete.shape)

def sum_eachday(df_POs):
	for day in range(len(eachday)):
		df_POs.at['Total',eachday[day]]=df_POs[eachday[day]].sum()
		df_POs.at['Largest', eachday[day]] = df_POs[eachday[day]].max()
		df_POs.at['Average', eachday[day]] = df_POs[eachday[day]].mean()
	PO_totaldemand_per_day=list(df_POs.loc['Total'])
	PO_largest_per_day=list(df_POs.loc['Largest'])
	PO_average_per_day=list(df_POs.loc['Average'])
	data=[PO_totaldemand_per_day,PO_largest_per_day,PO_average_per_day]
	PO_per_day=(pd.DataFrame(data)).transpose()
	PO_per_day.columns=['total','largest','average']
	return PO_per_day

writer_PO=pd.ExcelWriter('PO_per_day.xlsx',engine='xlsxwriter')
sum_eachday(df_PO_concrete).to_excel(writer_PO,sheet_name='col_PO_per_day',index=False)
sum_eachday(df_PO_rebar).to_excel(writer_PO,sheet_name='rebar_PO_per_day',index=False)
sum_eachday(df_PO_col).to_excel(writer_PO,sheet_name='concrete_PO_per_day',index=False)
writer_PO.save()

#assign PO IDs


"""Step5: save dataframes to excel files
========================================================================"""
writer_QTO=pd.ExcelWriter('QTOs_extended.xlsx',engine='xlsxwriter')
df_slab.to_excel(writer_QTO,sheet_name='Slab_exe', index=False)
df_shearwall.to_excel(writer_QTO,sheet_name='Shearcorewall_exe', index=False)
df_col.to_excel(writer_QTO,sheet_name='Col_exe', index=False)
writer_QTO.save()



""""Step6: plotting the consumption lines"""

