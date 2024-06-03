import os
import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from PyQt5.QtGui import QFont, QIcon
import PyQt5 as Qt
from datetime import datetime, timedelta
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from EPOL0E import EPOL0E
from ISD1A import ISD1A
from DECK1A import DECK1A
from SEED1A import SEED1A
from SL1AB import SL1AB
from AG1A import AG1A
from SCHN1A import SCHN1A
from ME1A_Ag import ME1A_Ag
from ME1A_Cu import ME1A_Cu
from TSXL import TSXL
from Summary import Summary

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Yield Reporter')
        self.setGeometry(0, 0, 800, 600)
        self.setWindowIcon(QIcon(resource_path('icon.png')))

        self.tabWidget = QTabWidget()
        self.setCentralWidget(self.tabWidget)

        self.EPOL0E = EPOL0E()
        self.ISD1A = ISD1A()
        self.DECK1A = DECK1A()
        self.SEED1A = SEED1A()
        self.SL1AB = SL1AB()
        self.AG1A = AG1A()
        self.SCHN1A = SCHN1A()
        self.ME1A_Ag = ME1A_Ag()
        self.ME1A_Cu = ME1A_Cu()
        self.TSXL = TSXL()
        self.Summary = Summary(self.EPOL0E, self.ISD1A, self.DECK1A, self.SEED1A, self.SL1AB, self.AG1A, self.SCHN1A, self.ME1A_Ag, self.ME1A_Cu, self.TSXL)
        
        self.tabWidget.addTab(self.EPOL0E, "EPOL0E")
        self.tabWidget.addTab(self.ISD1A, "ISD1A")
        self.tabWidget.addTab(self.DECK1A, "DECK1A")
        self.tabWidget.addTab(self.SEED1A, "SEED1A")
        self.tabWidget.addTab(self.SL1AB, "SL1A\\SL1B")
        self.tabWidget.addTab(self.AG1A, "AG1A")
        self.tabWidget.addTab(self.SCHN1A, "SCHN1A")
        self.tabWidget.addTab(self.ME1A_Ag, "ME1A Ag")
        self.tabWidget.addTab(self.ME1A_Cu, "ME1A Cu")
        self.tabWidget.addTab(self.TSXL, "TSXL")
        self.tabWidget.addTab(self.Summary, "Summary")
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont()
    font.setPointSize(10)
    app.setFont(font)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())