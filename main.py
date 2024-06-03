import os
import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from PyQt5.QtGui import QFont, QIcon
import PyQt5 as Qt
from datetime import datetime, timedelta
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
# Import tab
from Tab import Tab

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Yield Reporter')
        self.setGeometry(0, 0, 800, 600)
        # Set icon, can be omitted
        self.setWindowIcon(QIcon(resource_path('icon.png')))

        self.tabWidget = QTabWidget()
        self.setCentralWidget(self.tabWidget)

        # Assign tab
        self.Tab = Tab()
        
        self.tabWidget.addTab(self.Tab, "Tab")
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont()
    font.setPointSize(10)
    app.setFont(font)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())
