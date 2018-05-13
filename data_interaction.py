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
	dfs=pd.read_excel('BIM_QTOs.xlsx', sheetname=sheet_name,skiprows=None,
	                       converters=converter)
	return dfs

df_slab = df_generator('Slab',Con_slab)
df_shearwall = df_generator('CoreShearWall',Con_shearwall)
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
# scenario 1: insitu, rebar is proportional to days, concrete is one-day delivery for the whole floor
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
total_the_day.to_excel(writer_total, sheet_name='total_the_day', index=True)
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
	PO_totaldemand_per_day=list(df_POs.loc['Total'])
	PO_largest_per_day=list(df_POs.loc['Largest'])
	data=[PO_totaldemand_per_day,PO_largest_per_day]
	PO_per_day=(pd.DataFrame(data)).transpose()
	PO_per_day.columns=['total','largest_WBS']
	return PO_per_day

writer_PO=pd.ExcelWriter('PO_per_day.xlsx',engine='xlsxwriter')
sum_eachday(df_PO_concrete).to_excel(writer_PO,sheet_name='col_PO_per_day',index=False)
sum_eachday(df_PO_rebar).to_excel(writer_PO,sheet_name='rebar_PO_per_day',index=False)
sum_eachday(df_PO_col).to_excel(writer_PO,sheet_name='concrete_PO_per_day',index=False)
writer_PO.save()

"""Step5: save dataframes to excel files
========================================================================"""
writer_QTO=pd.ExcelWriter('BIM_QTOs_extended.xlsx',engine='xlsxwriter')
df_slab.to_excel(writer_QTO,sheet_name='Slab_exe', index=False)
df_shearwall.to_excel(writer_QTO,sheet_name='CoreShearWall_exe', index=False)
df_col.to_excel(writer_QTO,sheet_name='Col_exe', index=False)
writer_QTO.save()


""""Step6: find the material consumption date (needed date), more detailed with sequencing from WBS code start/finish date"""
#first identify the material consumption is in-situ or prefab. if in-situ, needed for sequencing, otherwise the same as BIM elements
#suppose the sequencing has four sub-tasks for insitu: 35%(of duration)for formwork, 35% for rebar, 15% for concreting and 15% removal of formwork
index_dfmaterial=[i for i in range(df_schedule.shape[0]*4)]
column_dfmaterial=['WBS_code','Duration','Sequencing','Start_Date','Finish_Date', 'Quantity']
df_materialneeds = pd.DataFrame(np.zeros((df_schedule.shape[0]*4,6)),index=index_dfmaterial, columns=column_dfmaterial)
df_materialneeds['Quantity']=0
for task in range(df_schedule.shape[0]):
	df_schedule.loc[task, 'Duration'] = float(df_schedule.loc[task, 'Duration'].replace(' dys', '').replace(' dy', ''))
	if df_schedule.loc[task, 'WBS_code'][12]=='P':
		df_materialneeds.loc[4 * task, 'WBS_code'] = df_schedule.loc[task, 'WBS_code']
		df_materialneeds.loc[4 * task, 'Duration'] = float(df_schedule.loc[task, 'Duration'])
		df_materialneeds.loc[4 * task, 'Sequencing'] = df_schedule.loc[task, 'WBS_code'] + 'NS'
		df_materialneeds.loc[4 * task, 'Start_Date'] = pd.to_datetime(df_schedule.loc[task, 'Start_Date'])
		df_materialneeds.loc[4 * task, 'Finish_Date'] = pd.to_datetime(df_schedule.loc[task, 'Finish_Date'])
	else:
		df_materialneeds.loc[4*task,'WBS_code']=df_schedule.loc[task,'WBS_code']
		df_materialneeds.loc[4*task,'Duration']=float(df_schedule.loc[task,'Duration'])*0.35
		df_materialneeds.loc[4*task,'Sequencing']=df_schedule.loc[task,'WBS_code']+'S1'
		df_materialneeds.loc[4*task,'Start_Date']=pd.to_datetime(df_schedule.loc[task,'Start_Date'])
		df_materialneeds.loc[4*task,'Finish_Date']=pd.to_datetime(df_schedule.loc[task,'Start_Date'])+timedelta(float(df_schedule.loc[task,'Duration'])*0.35)

		df_materialneeds.loc[4*task+1,'WBS_code']=df_schedule.loc[task,'WBS_code']
		df_materialneeds.loc[4*task+1,'Duration']=float(df_schedule.loc[task,'Duration'])*0.35
		df_materialneeds.loc[4*task+1,'Sequencing']=df_schedule.loc[task,'WBS_code']+'S2'
		df_materialneeds.loc[4*task+1,'Start_Date']=df_materialneeds.loc[4*task,'Finish_Date']
		df_materialneeds.loc[4*task+1,'Finish_Date']=pd.to_datetime(df_schedule.loc[task,'Start_Date'])+timedelta(float(df_schedule.loc[task,'Duration'])*0.7)

		df_materialneeds.loc[4*task+2,'WBS_code']=df_schedule.loc[task,'WBS_code']
		df_materialneeds.loc[4*task+2,'Duration']=float(df_schedule.loc[task,'Duration'])*0.15
		df_materialneeds.loc[4*task+2,'Sequencing']=df_schedule.loc[task,'WBS_code']+'S3'
		df_materialneeds.loc[4*task+2,'Start_Date']=df_materialneeds.loc[4*task+1,'Finish_Date']
		df_materialneeds.loc[4*task+2,'Finish_Date']=pd.to_datetime(df_schedule.loc[task,'Start_Date'])+timedelta(float(df_schedule.loc[task,'Duration'])*0.85)

		df_materialneeds.loc[4*task+3,'WBS_code']=df_schedule.loc[task,'WBS_code']
		df_materialneeds.loc[4*task+3,'Duration']=float(df_schedule.loc[task,'Duration'])*0.15
		df_materialneeds.loc[4*task+3,'Sequencing']=df_schedule.loc[task,'WBS_code']+'S4'
		df_materialneeds.loc[4*task+3,'Start_Date']=df_materialneeds.loc[4*task+2,'Finish_Date']
		df_materialneeds.loc[4*task+3,'Finish_Date']=pd.to_datetime(df_schedule.loc[task,'Start_Date'])+timedelta(float(df_schedule.loc[task,'Duration']))

#link material quantity with material need date, S2 corresponds to rebar, S3 corresponds to concrete, differentiate prefab element and insitu raw material
for i in range(total_the_day.shape[0]):
	for task in range(df_materialneeds.shape[0]):
		if df_materialneeds.loc[task,'WBS_code']==total_the_day.loc[i,'WBS_code'] and 'S2' in df_materialneeds.loc[task,'Sequencing']:
			df_materialneeds.loc[task,'Quantity']=total_the_day.loc[i,'Estimated Reinforcement Volume']
		if df_materialneeds.loc[task,'WBS_code']==total_the_day.loc[i,'WBS_code'] and 'S3' in df_materialneeds.loc[task,'Sequencing']:
			df_materialneeds.loc[task,'Quantity']=total_the_day.loc[i,'Volume']
		if df_materialneeds.loc[task,'WBS_code']==total_the_day.loc[i,'WBS_code'] and 'NS' in df_materialneeds.loc[task,'Sequencing']:
			df_materialneeds.loc[task,'Quantity']=total_the_day.loc[i,'counts']
		else:
			pass


df_PO_col1 = pd.DataFrame(np.zeros((df_materialneeds.shape[0],len(eachday))),index=df_materialneeds.index, columns=eachday)
df_PO_col1.columns = df_PO_col1.columns.astype(str)

# column delivery is linear day-by-day
for task in range(df_materialneeds.shape[0]):
	for day in range(len(eachday)):
		eachdaytime=pd.to_datetime(eachday[day])
		df_materialneeds.loc[task, 'Finish_Date']=pd.to_datetime(df_materialneeds.loc[task, 'Finish_Date'])
		df_materialneeds.loc[task, 'Start_Date']=pd.to_datetime(df_materialneeds.loc[task, 'Start_Date'])
		if eachdaytime<=df_materialneeds.loc[task,'Finish_Date'] and eachdaytime>=df_materialneeds.loc[task,'Start_Date'] and df_materialneeds.loc[task,'Sequencing'][-2:]=='NS':
			df_PO_col1.loc[task,eachday[day]]=int(df_materialneeds.loc[task,'Quantity']/df_materialneeds.loc[task,'Duration'])+1
		else:
			pass

# raw material, rebar, concrete delivery is one-day before the start-date
df_PO_rebar1 = pd.DataFrame(np.zeros((df_materialneeds.shape[0],len(eachday))),index=df_materialneeds.index, columns=eachday)
df_PO_rebar1.columns = df_PO_rebar1.columns.astype(str)
for task in range(df_materialneeds.shape[0]):
	for day in range(len(eachday)):
		eachdaytime=pd.to_datetime(eachday[day])
		df_materialneeds.loc[task, 'Finish_Date'] = pd.to_datetime(df_materialneeds.loc[task, 'Finish_Date'])
		df_materialneeds.loc[task, 'Start_Date'] = pd.to_datetime(df_materialneeds.loc[task, 'Start_Date'])
		if eachdaytime<=(df_materialneeds.loc[task,'Start_Date']+timedelta(0.5)) and eachdaytime>=(df_materialneeds.loc[task,'Start_Date']-timedelta(0.5)) and df_materialneeds.loc[task,'Sequencing'][-2:]=='S2':
			df_PO_rebar1.loc[task,eachday[day]]=df_materialneeds.loc[task,'Quantity']
		else:
			pass

df_PO_concrete1 = pd.DataFrame(np.zeros((df_materialneeds.shape[0], len(eachday))), index=df_materialneeds.index, columns=eachday)
df_PO_concrete1.columns = df_PO_concrete1.columns.astype(str)
for task in range(df_materialneeds.shape[0]):
	for day in range(len(eachday)):
		eachdaytime=pd.to_datetime(eachday[day])
		df_materialneeds.loc[task, 'Finish_Date'] = pd.to_datetime(df_materialneeds.loc[task, 'Finish_Date'])
		df_materialneeds.loc[task, 'Start_Date'] = pd.to_datetime(df_materialneeds.loc[task, 'Start_Date'])
		if eachdaytime<=(df_materialneeds.loc[task,'Start_Date']+timedelta(0.5)) and eachdaytime>=(df_materialneeds.loc[task,'Start_Date']-timedelta(0.5)) and df_materialneeds.loc[task, 'Sequencing'][-2:]=='S3':
			df_PO_concrete1.loc[task,eachday[day]]=df_materialneeds.loc[task,'Quantity']

df_PO_col2=sum_eachday(df_PO_col1)
df_PO_rebar2=sum_eachday(df_PO_rebar1)
df_PO_concrete2=sum_eachday(df_PO_concrete1)

writer_materialneeds = pd.ExcelWriter('material_needs.xlsx', engine='xlsxwriter')
df_materialneeds.to_excel(writer_materialneeds, sheet_name='total_material_needs', index=False)

df_PO_col1.to_excel(writer_materialneeds, sheet_name='material_needs_prefabcol', index=False)
df_PO_rebar1.to_excel(writer_materialneeds, sheet_name='material_needs_insiturebar', index=False)
df_PO_concrete1.to_excel(writer_materialneeds, sheet_name='material_needs_insituconcrete', index=False)

df_PO_col2.to_excel(writer_materialneeds, sheet_name='sum_prefabcol', index=False)
df_PO_rebar2.to_excel(writer_materialneeds, sheet_name='sum_insiturebar', index=False)
df_PO_concrete2.to_excel(writer_materialneeds, sheet_name='sum_insituconcrete', index=False)
writer_materialneeds.save()


""""step 7: control the 7-flow ocnstraints-flow maturity index and create material delivery sheets"""
# import matplotlib.pyplot as plt
# sel_day=input('select lookahead start day in a year (-th):')
# LAS_start_date=eachday[int(sel_day)]
#
# x1=pd.to_datetime(LAS_start_date)+timedelta(1)
# x2=pd.to_datetime(LAS_start_date)+timedelta(2)
# x3=pd.to_datetime(LAS_start_date)+timedelta(3)
# x4=pd.to_datetime(LAS_start_date)+timedelta(4)
# x5=pd.to_datetime(LAS_start_date)+timedelta(5)
# x6=pd.to_datetime(LAS_start_date)+timedelta(6)
# x7=pd.to_datetime(LAS_start_date)+timedelta(7)
# x8=pd.to_datetime(LAS_start_date)+timedelta(8)
# x9=pd.to_datetime(LAS_start_date)+timedelta(9)
# x10=pd.to_datetime(LAS_start_date)+timedelta(10)
# x=[x1,x2,x3,x4,x5,x6,x7,x8,x9,x10]
#
# def y_accumu(dfs,sel):
# 	y1=dfs.loc[sel,'total']
# 	y2 = y1 + dfs.loc[sel + 1, 'total']
# 	y3 = y2 + dfs.loc[sel + 2, 'total']
# 	y4 = y3 + dfs.loc[sel + 3, 'total']
# 	y5 = y4 + dfs.loc[sel + 4, 'total']
# 	y6 = y5 + dfs.loc[sel + 5, 'total']
# 	y7 = y6 + dfs.loc[sel + 6, 'total']
# 	y8 = y7 + dfs.loc[sel + 7, 'total']
# 	y9 = y8 + dfs.loc[sel + 8, 'total']
# 	y10 = y9 + dfs.loc[sel + 9, 'total']
# 	y=[y1,y2,y3,y4,y5,y6,y7,y8,y9,y10]
# 	return y
#
# y_col=y_accumu(df_PO_col2,sel_day)
# y_rebar=y_accumu(df_PO_rebar2,sel_day)
# y_concrete=y_accumu(df_PO_concrete2,sel_day)
#
# plt.gca().set_color_cycle(['red','green','yellow'])
# plt.plot(x,y_col)
# plt.plot(x,y_rebar)
# plt.plot(x,y_concrete)
#
# plt.legend()
# plt.show()

"""find the LAS start date and associated WBS_code and floor numbers"""
sel_day=input('select lookahead start day in a year (-th):')
LAS_start_date=eachday[int(sel_day)]
LAS_start_date=pd.to_datetime(LAS_start_date)
LAS_finish_date=LAS_start_date+timedelta(10)

list_WBS=[]
for i in range(df_materialneeds.shape[0]):
	df_materialneeds.loc[i, 'Start_Date']=pd.to_datetime(df_materialneeds.loc[i,'Start_Date'])
	df_materialneeds.loc[i,'Finish_Date']=pd.to_datetime(df_materialneeds.loc[i,'Finish_Date'])
	if LAS_start_date>=df_materialneeds.loc[i,'Start_Date'] and LAS_start_date<=df_materialneeds.loc[i,'Finish_Date']:
		list_WBS.append(df_materialneeds.loc[i,'Sequencing'])
	else:
		pass

list_WBS1=[]
for i in list_WBS:
	if i not in list_WBS1:
		list_WBS1.append(i)

print(list_WBS1)
