# coding: utf8
from ansicht import Ansicht,ObservableCollection
from customclass._ansicht import RBLItem,UI
from excel import System
import xlsxwriter
from IGF_log import getlog,getloglocal
from pyrevit import script, forms
from System.Windows.Forms import FolderBrowserDialog,DialogResult
from rpw import revit
from System.Collections.Generic import List
from System.Windows import Application, Window
from System.Windows.Controls import UserControl
from System.Windows.Threading import DispatcherPriority, Dispatcher
from Autodesk.Revit.UI import *
from System import Action

__title__ = "UIWindows"
__doc__ = """
exportiert ausgewählte Bauteilliste.

[2022.06.20]
Version: 2.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

uiapp = revit.uiapp
# mainWindow = Application.Current.MainWindow
#

# 获取Revit的UI线程的Dispatcher对象
uiDispatcher = Dispatcher.FromThread(System.Threading.Thread.CurrentThread)

# 在UI线程上获取主窗口
mainWindow = None
def getMainWindow(dummy):
    global mainWindow
    mainWindow = Application.Current.MainWindow
uiDispatcher.Invoke(DispatcherPriority.Normal, Action[Window](getMainWindow),None)



# 获取当前在Revit中打开的所有窗口
windows = List[Window](mainWindow.OwnedWindows)
print(liste)
