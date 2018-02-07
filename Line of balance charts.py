import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
#
"""example plot multiple lines"""
# """read the excel files"""
# df_LOB = pd.read_excel('line of balance data.xlsx',sheetname = 'data1',
#                        skiprows=None,converters={'A1':float,'A2':float,
#                                                  'A3':float,'A4':float,
#                                                  'A5':float,'A6':float,
#                                                  'A7': float,'A8': float,
#                                                  'Zone':float})
# yZone = df_LOB['Zone']
#
# def Xdata(Activity):
#     X = df_LOB[Activity]
#     return X
#
# listAct = []
# for i in range(0,df_LOB.shape[1]-1):
#     a = 'A'+ str(i+1)
#     listAct.append(a)
#
# import matplotlib.pyplot as plt
# fig=plt.figure()
# fig.show()
# ax=fig.add_subplot(111)
#
# def myplot(X,Y,l):
#     ax.plot(X,Y,c='c',marker="+",label=l)
#
# for i in range(0,df_LOB.shape[1]-1):
#     myplot(df_LOB[listAct[i]],yZone,str(listAct[i]))
#     plt.text(df_LOB[listAct[i]][6],yZone[6],str(listAct[i]))
#
#
# plt.show()
"""example LOB lines"""
# import matplotlib.pyplot as plt
# import numpy as np
# fig=plt.figure()
# fig.show()
# ax = fig.add_subplot(111)
#
# list_x=np.array([0,4/0.75,9])
# list_y=np.array([0,4,4])
# list_xmat=np.array([list_x[0],list_x[0],list_x[2]])
# list_ymat=np.array([0,4,4])
#
# for i in range(1,5):
#     ax.plot(list_x,list_y,c='c',marker='+',label='Col')
#     plt.text(list_x[1],list_y[1],'Col')
#     list_x = list_x + 9
#     list_y = list_y + 4
#     ax.plot(list_xmat,list_ymat,c='r',label='material')
#     plt.text(list_xmat[1],list_ymat[1],'material')
#     list_xmat=list_xmat+9
#     list_ymat=list_ymat+4
#
# plt.show()


"""test_example LOB"""
fig = plt.figure()
fig.show()
ax = fig.add_subplot(111)

df_col = pd.read_excel('line of balance data.xlsx',sheetname = 'col',
                       skiprows=None,converters={'task_day':float,'y_val':float})
df_wallf = pd.read_excel('line of balance data.xlsx',sheetname = 'wallf',
                       skiprows=None,converters={'task_day':float,'y_val':float})
df_wall = pd.read_excel('line of balance data.xlsx',sheetname = 'wall',
                       skiprows=None,converters={'task_day':float,'y_val':float})
df_slabf = pd.read_excel('line of balance data.xlsx',sheetname = 'slabf',
                       skiprows=None,converters={'task_day':float,'y_val':float})
df_slab = pd.read_excel('line of balance data.xlsx',sheetname = 'slab',
                       skiprows=None,converters={'task_day':float,'y_val':float})


list_xcol = np.array(df_col['task_day'])
list_ycol = np.array(df_col['y_val'])
list_xwallf = np.array(df_wallf['task_day'])
list_ywallf = np.array(df_wallf['y_val'])
list_xwall = np.array(df_wall['task_day'])
list_ywall = np.array(df_wall['y_val'])
list_xslabf = np.array(df_slabf['task_day'])
list_yslabf = np.array(df_slabf['y_val'])
list_xslab = np.array(df_slab['task_day'])
list_yslab = np.array(df_slab['y_val'])
list_xmaterial = np.array([0,0,7])
list_ymaterial = np.array([0,6,6])

for i in range(1,5):
    ax.plot(list_xcol,list_ycol,c='c',marker='+',label='Col')
    plt.text(list_xcol[1],list_ycol[1],'Col')
    list_xcol = list_xcol + 9
    list_ycol = list_ycol + 4

    ax.plot(list_xwallf, list_ywallf, c='k', marker='+', label='wallf')
    plt.text(list_xwallf[2], list_ywallf[2], 'wallf')
    list_xwallf = list_xwallf + 9
    list_ywallf = list_ywallf + 4

    ax.plot(list_xwall, list_ywall, c='y', marker='+', label='wall')
    plt.text(list_xwall[3], list_ywall[3], 'wall')
    list_xwall = list_xwall + 9
    list_ywall = list_ywall + 4

    ax.plot(list_xslabf, list_yslabf, c='g', marker='+', label='slabf')
    plt.text(list_xslabf[5], list_yslabf[5], 'slabf')
    list_xslabf = list_xslabf + 9
    list_yslabf = list_yslabf + 4

    ax.plot(list_xslab, list_yslab, c='b', marker='+', label='slab')
    plt.text(list_xslab[7], list_yslab[7], 'Slab')
    list_xslab = list_xslab + 9
    list_yslab = list_yslab + 4

    ax.plot(list_xmaterial, list_ymaterial, c='r', marker='+', label='slab')
    plt.text(list_xmaterial[1], list_ymaterial[1], 'material')
    list_xmaterial = list_xmaterial + 7
    list_ymaterial = list_ymaterial + 6

plt.title('Line Of Balance chart',fontsize = 16)
plt.xlabel('look ahead days (unit: day)', fontsize = 12)
plt.ylabel('Floor-Zones (unit: zone)', fontsize = 12)
plt.show()
