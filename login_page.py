# background colors:    http://www.color-hex.com/
# create the login page
from tkinter import *
from tkinter import ttk
from sqlite3 import dbapi2 as sqlite
from PIL import ImageTk, Image
import webbrowser

login=sqlite.connect('lmf_data.sql')
lcur=login.cursor()
Winstat=''

"""openwindows: building_element.py   material_tracking.py    material_inspection.py     receipts.py"""

def shearwall_data():
    application.destroy()
    login.close()
    import building_element
    a1=building_element.shearwall_data()
    open_win()

def column_data():
    application.destroy()
    login.close()
    import building_element
    a2=building_element.column_data()
    open_win()

def slab_data():
    application.destroy()
    login.close()
    import building_element
    a3=building_element.slab_data()
    open_win()

def lookahead():
    application.destroy()
    login.close()
    import material_tracking
    b1=material_tracking.lookahead()
    open_win()

def POs():
    application.destroy()
    login.close()
    import material_tracking
    b2=material_tracking.POs()
    open_win()

def qualitycheck():
    application.destroy()
    login.close()
    import material_inspection
    c1=material_inspection.qualitycheck()
    open_win()

def materialstatus():
    application.destroy()
    login.close()
    import material_inspection
    c2=material_inspection.materialstatus()
    open_win()

def PO_summary():
    application.destroy()
    login.close()
    import receipts
    d1 = receipts.PO_summary()
    open_win()

"""--------------------------------------------------------------------------------------------"""
def projectinfo():
    root1=Tk()
    root1.title('Project:Zurich Oerlikon, Andreasturm')
    root1.wm_iconbitmap('logo.ico')
    # root.configure(background='#e5eef4')
    info=open('Andreasturm_Information.txt','r').read()
    t1=Text(root1)
    t1.insert(INSERT,info)
    t1.pack()

def callwebsite(event):
    webbrowser.open_new(r"andreasturm.ch")

def main_window():
    global username,pwd,Winstat,root,application
    if Winstat=='application':
        application.destroy()
    root=Tk()
    root.title('Lean Material Flows')
    root.wm_iconbitmap('logo.ico')
    root.configure(background='#e5eef4')

    # l1=Label(root, text='BIM-RFID based Construction Material Management System',background='#e5eef4')
    # l1.grid(row=1,column=0)
    # l1.config(font=('Courier',10))

    Label(root,text='Username').grid(row=0,column=0)
    username=Entry(root,width=10)
    username.grid(row=0,column=1)
    Label(root, text='Password').grid(row=1, column=0)
    pwd = Entry(root, width=10, show='*')
    pwd.grid(row=1, column=1)
    Button(root,width=6, text='Login', command=checkpassword).grid(row=2,column=0)
    Button(root, width=6, text='Close', command=root.destroy).grid(row=2, column=1)

    root.resizable(width=False,height=False)

    root.mainloop()


def open_win():
    global application,Winstat
    Winstat='application'
    application=Tk()
    application.wm_iconbitmap('logo.ico')
    application.title('Andreasturm Project construction material management')

    image1 = Image.open('project.jpg')
    image1 = image1.resize((600, 600), Image.ANTIALIAS)
    img = ImageTk.PhotoImage(image1)
    panel = Label(application, image=img).grid(row=0, column=0)

    Button(application, text='Project Information', command=projectinfo).grid(row=1, column=0)
    l2 = Label(application, text='andreasturm.ch', fg='blue', cursor='hand2')
    l2.bind("<Button-1>", callwebsite)
    l2.grid(row=2, column=0)

    menu_bar=Menu(application)
    element_menu=Menu(menu_bar,tearoff=0)
    tracking_menu=Menu(menu_bar,tearoff=0)
    inspection_menu=Menu(menu_bar,tearoff=0)
    receipt_menu=Menu(menu_bar,tearoff=0)

    element_menu.add_command(label='Shearwall Element', command=shearwall_data)
    element_menu.add_command(label='Column Element', command=column_data)
    element_menu.add_command(label='Slab Element',command=slab_data)

    tracking_menu.add_command(label='Look Ahead Plan & Schedules', command=lookahead)
    tracking_menu.add_command(label='Purchasing Orders',command=POs)

    inspection_menu.add_command(label='Quality Check', command=qualitycheck)
    inspection_menu.add_command(label='Material Status',command=materialstatus)

    receipt_menu.add_command(label='PO summary',command=PO_summary)

    menu_bar.add_cascade(label='Building Elements',menu=element_menu)
    menu_bar.add_cascade(label='Material Tracking',menu=tracking_menu)
    menu_bar.add_cascade(label='Material Inspection',menu=inspection_menu)
    menu_bar.add_cascade(label='PO receipts', menu=receipt_menu)

    application.config(menu=menu_bar)

    application.resizable(width=False, height=False)
    application.mainloop()


def  checkpassword():
    global username,pwd,root
    u=username.get()
    p=pwd.get()
    if 'ibi'!=u and '1234'!=p:
        top=Tk()
        Label(top,width=30,text='wrong username or password').grid(row=0,column=0)
        # top.destroy()
        top.mainloop()
    else:
        root.destroy()
        open_win()


main_window()
