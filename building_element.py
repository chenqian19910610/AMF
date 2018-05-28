# building elements and material information
from tkinter import *
from sqlite3 import dbapi2 as sqlite
import webbrowser

c=sqlite.connect('lmf_data.sqlite')
cur=c.cursor()


"""---------------------menu_shearwall----------------------------------------------------------------"""

columns_name=['elementID','DeliveryID','POID','WBScode','Family','Type','Area',
              'Length','Width','Volume','Designload','BaseConstraint','EstimatedReinforcementVolume',
              'ConstructZone','FloorNo','Need_Date']

def onlineview(event):
    webbrowser.open_new(r"https://autode.sk/2Lyf2dn")

def shearwall_data():
    global cur,c,columns_name,value,flag,shearwall,application, lb1,lb2,lb3,lb4,lb5,lb6,lb7,lb8,lb9,lb10
    flag='shearwall'
    value=['']*len(columns_name)
    shearwall=Tk()
    shearwall.title('Shearwall Elements')
    shearwall.wm_iconbitmap('logo.ico')

    Button(shearwall,width=25,text='show all belonging elements',command=show_all).grid(row=0,column=0)
    l_view = Label(shearwall,text='view BIM online',fg='blue', cursor='hand2')
    l_view.grid(row=0,column=1)
    l_view.bind("<Button-1>", onlineview)
    Button(shearwall, width=15, text='Main Menu', command=mainmenu).grid(row=0, column=2)

    Label(shearwall,text='Floor No.').grid(row=1,column=0)
    value[0] = Entry(shearwall)
    value[0].grid(row=1, column=1)
    Button(shearwall, width=15, text='Queryby Floor', command=query1).grid(row=1, column=2)

    Label(shearwall, text='Element ID in BIM').grid(row=2, column=0, sticky=W)
    value[1] = Entry(shearwall)
    value[1].grid(row=2, column=1)
    Button(shearwall, width=15, text='Queryby ID', command=query2).grid(row=2, column=2)

    for i in range(4,12):
        Label(shearwall, text=columns_name[i]).grid(row=3,column=i-4)
    Label(shearwall,text=columns_name[14]).grid(row=3,column=8)
    Label(shearwall, text=columns_name[0]).grid(row=3, column=9)

    def scrollbarv(*args):
        lb1.yview(*args)
        lb2.yview(*args)
        lb3.yview(*args)
        lb4.yview(*args)
        lb5.yview(*args)
        lb6.yview(*args)
        lb7.yview(*args)
        lb8.yview(*args)
        lb9.yview(*args)
        lb10.yview(*args)

    scbar=Scrollbar(orient='vertical',command=scrollbarv)
    lb1 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb2 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb3 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb4 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb5 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb6 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb7 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb8 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb9 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb10 = Listbox(shearwall, yscrollcommand=scbar.set)
    lb1.grid(row=4, column=0)
    lb1.configure(width=0,height=0)
    lb2.grid(row=4, column=1)
    lb2.configure(width=0, height=0)
    lb3.grid(row=4, column=2)
    lb3.configure(width=0, height=0)
    lb4.grid(row=4, column=3)
    lb4.configure(width=0, height=0)
    lb5.grid(row=4, column=4)
    lb5.configure(width=0, height=0)
    lb6.grid(row=4, column=5)
    lb6.configure(width=0, height=0)
    lb7.grid(row=4, column=6)
    lb7.configure(width=0, height=0)
    lb8.grid(row=4, column=7)
    lb8.configure(width=0, height=0)
    lb9.grid(row=4,column=8)
    lb9.configure(width=0, height=0)
    lb10.grid(row=4, column=9)
    lb10.configure(width=0, height=0)
    scbar.grid(row=4,column=10,stick=N+S)


    show_all()
    shearwall.resizable(width=False, height=False)
    shearwall.mainloop()

def show_all():
    global value,c,cur,columns_name,shearwall
    cur.execute('select * from P1_design_BIM_shearwall')

    for i in cur:
        lb1.insert(1, str(i[4]))
        lb2.insert(1, str(i[5]))
        lb3.insert(1, str(i[7])+'m2')
        lb4.insert(1, str(i[8]))
        lb5.insert(1, str(i[9]))
        lb6.insert(1, str(i[10])+'m3')
        lb7.insert(1, str(i[11]))
        lb8.insert(1, str(i[12]))
        lb9.insert(1, str(i[15]))
        lb10.insert(1, str(i[0]))

    c.commit()


def query1():
    global value, c, cur, columns_name, shearwall
    FloorNo = value[0].get()
    cur.execute('select * from P1_design_BIM_shearwall where FloorNo=?', [FloorNo])

    lb1.delete(0, END)
    lb2.delete(0, END)
    lb3.delete(0, END)
    lb4.delete(0, END)
    lb5.delete(0, END)
    lb6.delete(0, END)
    lb7.delete(0, END)
    lb8.delete(0, END)
    lb9.delete(0, END)
    lb10.delete(0, END)

    for i in cur:
        if i[0] == None:
            lb1.insert(1, 'Not Exist')
            lb2.insert(1, 'Not Exist')
            lb3.insert(1, 'Not Exist')
            lb4.insert(1, 'Not Exist')
            lb5.insert(1, 'Not Exist')
            lb6.insert(1, 'Not Exist')
            lb7.insert(1, 'Not Exist')
            lb8.insert(1, 'Not Exist')
            lb9.insert(1, 'Not Exist')
            lb10.insert(1, 'Not Exist')
        else:
            lb1.insert(1, str(i[4]))
            lb2.insert(1, str(i[5]))
            lb3.insert(1, str(i[7])+'m2')
            lb4.insert(1, str(i[8]))
            lb5.insert(1, str(i[9]))
            lb6.insert(1, str(i[10])+'m3')
            lb7.insert(1, str(i[11]))
            lb8.insert(1, str(i[12]))
            lb9.insert(1, str(i[15]))
            lb10.insert(1, str(i[0]))

    c.commit()

def query2():
    global value, c, cur, columns_name,shearwall
    elementID=value[1].get()
    cur.execute('select * from P1_design_BIM_shearwall where elementID=?',[elementID])

    lb1.delete(0, END)
    lb2.delete(0, END)
    lb3.delete(0, END)
    lb4.delete(0, END)
    lb5.delete(0, END)
    lb6.delete(0, END)
    lb7.delete(0, END)
    lb8.delete(0, END)
    lb9.delete(0, END)
    lb10.delete(0, END)

    for i in cur:
        if i[0]== None:
            lb1.insert(1, 'Not Exist')
            lb2.insert(1, 'Not Exist')
            lb3.insert(1, 'Not Exist')
            lb4.insert(1, 'Not Exist')
            lb5.insert(1, 'Not Exist')
            lb6.insert(1, 'Not Exist')
            lb7.insert(1, 'Not Exist')
            lb8.insert(1, 'Not Exist')
            lb9.insert(1, 'Not Exist')
            lb10.insert(1, 'Not Exist')
        else:
            lb1.insert(1, str(i[4]))
            lb2.insert(1, str(i[5]))
            lb3.insert(1, str(i[7]) + 'm2')
            lb4.insert(1, str(i[8]))
            lb5.insert(1, str(i[9]))
            lb6.insert(1, str(i[10]) + 'm3')
            lb7.insert(1, str(i[11]))
            lb8.insert(1, str(i[12]))
            lb9.insert(1, str(i[15]))
            lb10.insert(1, str(i[0]))

    Label(shearwall, text='WBS code').grid(row=5, column=0)
    lb_wbs = Listbox(shearwall)
    lb_wbs.grid(row=5, column=1)
    lb_wbs.configure(width=0, height=0)

    Label(shearwall, text='need date').grid(row=6, column=0)
    lb_needate = Listbox(shearwall)
    lb_needate.grid(row=6, column=1)
    lb_needate.configure(width=0, height=0)

    Label(shearwall, text='construct method').grid(row=7, column=0)
    lb_method = Listbox(shearwall)
    lb_method.grid(row=7, column=1)
    lb_method.configure(width=0, height=0)


    cur.execute('select * from TAKT_floor_04OG_10OG inner join P1_design_BIM_shearwall on P1_design_BIM_shearwall.WBScode=TAKT_floor_04OG_10OG.WBS_code where P1_design_BIM_shearwall.elementID=?',[elementID])
    for i in cur:
        print(i)
        lb_wbs.insert(1,i[0])
        lb_needate.insert(1,i[1])
    cur.execute('select * from P2_POIDs_WBS_coreshearwall where elementID=?',[elementID])
    for i in cur:
        lb_method.insert(1,i[16])
    c.commit()


"""---------------------menu_column----------------------------------------------------------------"""
columns_name_col=['elementID','DeliveryID','POID','WBScode','Family','Type',
              'Designload','Length','Volume','EstimatedReinforcementVolume','ConstructZone','FloorNo',
              'ColumnLocationMark','Need_Date']

def column_data():
    global cur,c,columns_name_col,value_col,flag,prefcolumn,application, lbcol1,lbcol2,lbcol3,lbcol4,lbcol5,lbcol6,lbcol7,\
        lbcol8
    flag='prefcolumn'
    value_col=['']*len(columns_name)
    prefcolumn=Tk()
    prefcolumn.title('Column Elements')
    prefcolumn.wm_iconbitmap('logo.ico')

    Button(prefcolumn,width=25,text='show all belonging elements',command=show_all_col).grid(row=0,column=0)
    l_view = Label(prefcolumn,text='view BIM online',fg='blue', cursor='hand2')
    l_view.grid(row=0,column=1)
    l_view.bind("<Button-1>", onlineview)
    Button(prefcolumn, width=15, text='Main Menu', command=mainmenu).grid(row=0, column=2)

    Label(prefcolumn,text='Floor No.').grid(row=1,column=0)
    value_col[0] = Entry(prefcolumn)
    value_col[0].grid(row=1, column=1)
    Button(prefcolumn, width=15, text='Queryby Floor', command=query1_col).grid(row=1, column=2)

    Label(prefcolumn, text='Element ID in BIM').grid(row=2, column=0, sticky=W)
    value_col[1] = Entry(prefcolumn)
    value_col[1].grid(row=2, column=1)
    Button(prefcolumn, width=15, text='Queryby ID', command=query2_col).grid(row=2, column=2)

    for i in range(4,9):
        Label(prefcolumn, text=columns_name_col[i]).grid(row=3,column=i-4)
    Label(prefcolumn,text=columns_name_col[11]).grid(row=3,column=5)
    Label(prefcolumn, text=columns_name_col[12]).grid(row=3, column=6)
    Label(prefcolumn, text=columns_name_col[0]).grid(row=3, column=7)


    def scrollbarv(*args):
        lbcol1.yview(*args)
        lbcol2.yview(*args)
        lbcol3.yview(*args)
        lbcol4.yview(*args)
        lbcol5.yview(*args)
        lbcol6.yview(*args)
        lbcol7.yview(*args)
        lbcol8.yview(*args)

    scbar=Scrollbar(orient='vertical',command=scrollbarv)
    lbcol1 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol2 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol3 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol4 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol5 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol6 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol7 = Listbox(prefcolumn, yscrollcommand=scbar.set)
    lbcol8 = Listbox(prefcolumn, yscrollcommand=scbar.set)

    lbcol1.grid(row=4, column=0)
    lbcol1.configure(width=0,height=0)
    lbcol2.grid(row=4, column=1)
    lbcol2.configure(width=0, height=0)
    lbcol3.grid(row=4, column=2)
    lbcol3.configure(width=0, height=0)
    lbcol4.grid(row=4, column=3)
    lbcol4.configure(width=0, height=0)
    lbcol5.grid(row=4, column=4)
    lbcol5.configure(width=0, height=0)
    lbcol6.grid(row=4, column=5)
    lbcol6.configure(width=0, height=0)
    lbcol7.grid(row=4, column=6)
    lbcol7.configure(width=0, height=0)
    lbcol8.grid(row=4, column=7)
    lbcol8.configure(width=0, height=0)
    scbar.grid(row=4,column=8,stick=N+S)


    show_all_col()
    prefcolumn.resizable(width=False, height=False)
    prefcolumn.mainloop()

def show_all_col():
    global value_col,c,cur,columns_name_col,prefcolumn
    cur.execute('select * from P1_design_BIM_column')

    for i in cur:
        lbcol1.insert(1, str(i[4]))
        lbcol2.insert(1, str(i[5]))
        lbcol3.insert(1, str(i[7])+'kN')
        lbcol4.insert(1, str(i[8])+'m')
        lbcol5.insert(1, str(i[9])+'m3')
        lbcol6.insert(1, str(i[12]))
        lbcol7.insert(1, str(i[13]))
        lbcol8.insert(1, str(i[0]))

    c.commit()

def query1_col():
    global value_col, c, cur, columns_name_col, prefcolumn
    FloorNo = value_col[0].get()
    cur.execute('select * from P1_design_BIM_column where FloorNo=?', [FloorNo])

    lbcol1.delete(0, END)
    lbcol2.delete(0, END)
    lbcol3.delete(0, END)
    lbcol4.delete(0, END)
    lbcol5.delete(0, END)
    lbcol6.delete(0, END)
    lbcol7.delete(0, END)
    lbcol8.delete(0, END)

    for i in cur:
        if i[0] == None:
            lbcol1.insert(1, 'Not Exist')
            lbcol2.insert(1, 'Not Exist')
            lbcol3.insert(1, 'Not Exist')
            lbcol4.insert(1, 'Not Exist')
            lbcol5.insert(1, 'Not Exist')
            lbcol6.insert(1, 'Not Exist')
            lbcol7.insert(1, 'Not Exist')
            lbcol8.insert(1, 'Not Exist')
        else:
            lbcol1.insert(1, str(i[4]))
            lbcol2.insert(1, str(i[5]))
            lbcol3.insert(1, str(i[7]) + 'kN')
            lbcol4.insert(1, str(i[8]) + 'm')
            lbcol5.insert(1, str(i[9]) + 'm3')
            lbcol6.insert(1, str(i[12]))
            lbcol7.insert(1, str(i[13]))
            lbcol8.insert(1, str(i[0]))

    c.commit()

def query2_col():
    global value_col, c, cur, columns_name_col,prefcolumn
    elementID = value_col[1].get()
    cur.execute('select * from P1_design_BIM_column where elementID=?',[elementID])

    lbcol1.delete(0, END)
    lbcol2.delete(0, END)
    lbcol3.delete(0, END)
    lbcol4.delete(0, END)
    lbcol5.delete(0, END)
    lbcol6.delete(0, END)
    lbcol7.delete(0, END)
    lbcol8.delete(0, END)

    for i in cur:
        if i[0]== None:
            lbcol1.insert(1, 'Not Exist')
            lbcol2.insert(1, 'Not Exist')
            lbcol3.insert(1, 'Not Exist')
            lbcol4.insert(1, 'Not Exist')
            lbcol5.insert(1, 'Not Exist')
            lbcol6.insert(1, 'Not Exist')
            lbcol7.insert(1, 'Not Exist')
            lbcol8.insert(1, 'Not Exist')

        else:
            lbcol1.insert(1, str(i[4]))
            lbcol2.insert(1, str(i[5]))
            lbcol3.insert(1, str(i[7]) + 'kN')
            lbcol4.insert(1, str(i[8]) + 'm')
            lbcol5.insert(1, str(i[9]) + 'm3')
            lbcol6.insert(1, str(i[12]))
            lbcol7.insert(1, str(i[13]))
            lbcol8.insert(1, str(i[0]))

    Label(prefcolumn, text='WBS code').grid(row=5, column=0)
    lb_wbs = Listbox(prefcolumn)
    lb_wbs.grid(row=5, column=1)
    lb_wbs.configure(width=0, height=0)

    Label(prefcolumn, text='need date').grid(row=6, column=0)
    lb_needate = Listbox(prefcolumn)
    lb_needate.grid(row=6, column=1)
    lb_needate.configure(width=0, height=0)

    Label(prefcolumn, text='construct method').grid(row=7, column=0)
    lb_method = Listbox(prefcolumn)
    lb_method.grid(row=7, column=1)
    lb_method.configure(width=0, height=0)


    cur.execute('select * from TAKT_floor_04OG_10OG inner join P1_design_BIM_column on P1_design_BIM_column.WBScode=TAKT_floor_04OG_10OG.WBS_code where P1_design_BIM_column.elementID=?',[elementID])
    for i in cur:
        print(i)
        lb_wbs.insert(1,i[0])
        lb_needate.insert(1,i[1])
    cur.execute('select * from P2_POIDs_WBS_column where elementID=?',[elementID])
    for i in cur:
        lb_method.insert(1,i[15])
    c.commit()

"""------------------------------------menu_slab-------------------------------------------------------------------"""
columns_name_slab=['elementID','DeliveryID','POID','WBScode','Family','Type','Volume',
              'Area','Perimeter','CoreThickness','FloorNo']

def slab_data():
    global cur,c,columns_name_slab,value_sl,flag,slab,application, lb1,lb2,lb3,lb4,lb5,lb6,lb7,lb8,lb9,lb10
    flag='slab'
    value_sl=['']*len(columns_name_slab)
    slab=Tk()
    slab.title('Slab Elements')
    slab.wm_iconbitmap('logo.ico')

    Button(slab,width=25,text='show all belonging elements',command=show_all_slab).grid(row=0,column=0)
    l_view = Label(slab,text='view BIM online',fg='blue', cursor='hand2')
    l_view.grid(row=0,column=1)
    l_view.bind("<Button-1>", onlineview)
    Button(slab, width=15, text='Main Menu', command=mainmenu).grid(row=0, column=2)

    Label(slab,text='Floor No.').grid(row=1,column=0)
    value_sl[0] = Entry(slab)
    value_sl[0].grid(row=1, column=1)
    Button(slab, width=15, text='Queryby Floor', command=query1_slab).grid(row=1, column=2)

    Label(slab, text='Element ID in BIM').grid(row=2, column=0, sticky=W)
    value_sl[1] = Entry(slab)
    value_sl[1].grid(row=2, column=1)
    Button(slab, width=15, text='Queryby ID', command=query2_slab).grid(row=2, column=2)

    for i in range(4,10):
        Label(slab, text=columns_name_slab[i]).grid(row=3,column=i-4)
    Label(slab,text=columns_name_slab[10]).grid(row=3,column=6)
    Label(slab, text=columns_name_slab[0]).grid(row=3, column=7)

    def scrollbarv(*args):
        lb1.yview(*args)
        lb2.yview(*args)
        lb3.yview(*args)
        lb4.yview(*args)
        lb5.yview(*args)
        lb6.yview(*args)
        lb7.yview(*args)
        lb8.yview(*args)


    scbar=Scrollbar(orient='vertical',command=scrollbarv)
    lb1 = Listbox(slab, yscrollcommand=scbar.set)
    lb2 = Listbox(slab, yscrollcommand=scbar.set)
    lb3 = Listbox(slab, yscrollcommand=scbar.set)
    lb4 = Listbox(slab, yscrollcommand=scbar.set)
    lb5 = Listbox(slab, yscrollcommand=scbar.set)
    lb6 = Listbox(slab, yscrollcommand=scbar.set)
    lb7 = Listbox(slab, yscrollcommand=scbar.set)
    lb8 = Listbox(slab, yscrollcommand=scbar.set)

    lb1.grid(row=4, column=0)
    lb1.configure(width=0,height=0)
    lb2.grid(row=4, column=1)
    lb2.configure(width=0, height=0)
    lb3.grid(row=4, column=2)
    lb3.configure(width=0, height=0)
    lb4.grid(row=4, column=3)
    lb4.configure(width=0, height=0)
    lb5.grid(row=4, column=4)
    lb5.configure(width=0, height=0)
    lb6.grid(row=4, column=5)
    lb6.configure(width=0, height=0)
    lb7.grid(row=4, column=6)
    lb7.configure(width=0, height=0)
    lb8.grid(row=4, column=7)
    lb8.configure(width=0, height=0)

    scbar.grid(row=4,column=8,stick=N+S)


    show_all_slab()
    slab.resizable(width=False, height=False)
    slab.mainloop()

def show_all_slab():
    global value_sl,c,cur,columns_name_slab,slab
    cur.execute('select * from P1_design_BIM_slab')

    for i in cur:
        lb1.insert(1, str(i[4]))
        lb2.insert(1, str(i[5]))
        lb3.insert(1, str(i[7])+'m3')
        lb4.insert(1, str(i[8])+'m2')
        lb5.insert(1, str(i[9]))
        lb6.insert(1, str(i[11]))
        lb7.insert(1, str(i[13]))
        lb8.insert(1, str(i[0]))


    c.commit()


def query1_slab():
    global value_sl, c, cur, columns_name_slab, slab
    FloorNo = value_sl[0].get()
    cur.execute('select * from P1_design_BIM_slab where FloorNo=?', [FloorNo])

    lb1.delete(0, END)
    lb2.delete(0, END)
    lb3.delete(0, END)
    lb4.delete(0, END)
    lb5.delete(0, END)
    lb6.delete(0, END)
    lb7.delete(0, END)
    lb8.delete(0, END)

    for i in cur:
        if i[0] == None:
            lb1.insert(1, 'Not Exist')
            lb2.insert(1, 'Not Exist')
            lb3.insert(1, 'Not Exist')
            lb4.insert(1, 'Not Exist')
            lb5.insert(1, 'Not Exist')
            lb6.insert(1, 'Not Exist')
            lb7.insert(1, 'Not Exist')
            lb8.insert(1, 'Not Exist')

        else:
            lb1.insert(1, str(i[4]))
            lb2.insert(1, str(i[5]))
            lb3.insert(1, str(i[7]) + 'm3')
            lb4.insert(1, str(i[8]) + 'm2')
            lb5.insert(1, str(i[9]))
            lb6.insert(1, str(i[11]))
            lb7.insert(1, str(i[13]))
            lb8.insert(1, str(i[0]))

    c.commit()

def query2_slab():
    global value, c, cur, columns_name_slab,slab
    elementID = value_sl[1].get()
    cur.execute('select * from P1_design_BIM_slab where elementID=?',[elementID])

    lb1.delete(0, END)
    lb2.delete(0, END)
    lb3.delete(0, END)
    lb4.delete(0, END)
    lb5.delete(0, END)
    lb6.delete(0, END)
    lb7.delete(0, END)
    lb8.delete(0, END)


    for i in cur:
        if i[0]== None:
            lb1.insert(1, 'Not Exist')
            lb2.insert(1, 'Not Exist')
            lb3.insert(1, 'Not Exist')
            lb4.insert(1, 'Not Exist')
            lb5.insert(1, 'Not Exist')
            lb6.insert(1, 'Not Exist')
            lb7.insert(1, 'Not Exist')
            lb8.insert(1, 'Not Exist')

        else:
            lb1.insert(1, str(i[4]))
            lb2.insert(1, str(i[5]))
            lb3.insert(1, str(i[7]) + 'm3')
            lb4.insert(1, str(i[8]) + 'm2')
            lb5.insert(1, str(i[9]))
            lb6.insert(1, str(i[11]))
            lb7.insert(1, str(i[13]))
            lb8.insert(1, str(i[0]))


    Label(slab, text='WBS code').grid(row=5, column=0)
    lb_wbs = Listbox(slab)
    lb_wbs.grid(row=5, column=1)
    lb_wbs.configure(width=0, height=0)

    Label(slab, text='need date').grid(row=6, column=0)
    lb_needate = Listbox(slab)
    lb_needate.grid(row=6, column=1)
    lb_needate.configure(width=0, height=0)

    Label(slab, text='construct method').grid(row=7, column=0)
    lb_method = Listbox(slab)
    lb_method.grid(row=7, column=1)
    lb_method.configure(width=0, height=0)


    cur.execute('select * from TAKT_floor_04OG_10OG inner join P1_design_BIM_slab on P1_design_BIM_slab.WBScode=TAKT_floor_04OG_10OG.WBS_code where P1_design_BIM_slab.elementID=?',[elementID])
    for i in cur:
        print(i)
        lb_wbs.insert(1,i[0])
        lb_needate.insert(1,i[1])

    cur.execute('select * from P2_POIDs_WBS_slab where elementID=?',[elementID])
    for i in cur:
        lb_method.insert(1,i[15])
    c.commit()




def mainmenu():
    if flag=='shearwall':
        shearwall.destroy()
    if flag=='prefcolumn':
        prefcolumn.destroy()
    if flag=='slab':
        slab.destroy()