# building elements and material information
from tkinter import *
from sqlite3 import dbapi2 as sqlite
import webbrowser

c=sqlite.connect('lmf_data.sqlite')
cur=c.cursor()

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

    cur.execute('select * from TAKT_floor_04OG_10OG inner join P1_design_BIM_shearwall on P1_design_BIM_shearwall.WBScode=TAKT_floor_04OG_10OG.WBS_code where P1_design_BIM_shearwall.elementID=?',[elementID])
    for i in cur:
        print(i)
        lb_wbs.insert(1,i[0])
        lb_needate.insert(1,i[1])


    c.commit()


def mainmenu():
    if flag=='shearwall':
        shearwall.destroy()