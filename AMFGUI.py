import sys
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication,QDialog,QMainWindow
from PyQt5.uic import loadUi

class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        loadUi('AMFGUI.ui', self)

        




app=QApplication(sys.argv)
window=MyWindow()
window.show()
sys.exit(app.exec())
