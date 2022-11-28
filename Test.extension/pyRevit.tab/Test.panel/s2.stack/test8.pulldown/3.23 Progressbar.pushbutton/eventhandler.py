# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from pyrevit import script
from System.Windows.Media import Brushes
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Windows import Visibility
from System.Windows.Input import Key
import System

Visible = Visibility.Visible
Hidden = Visibility.Hidden

Schwarz = Brushes.Black
Rot = Brushes.Red

class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

class UEBERNEHMEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        try:

            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            self.GUI.pb_t.Visibility = Visible
            self.GUI.pb_c.Visibility = Visible
            for n in range(1000):
                self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                # print(n)
                self.GUI.maxvalue = 1000
                self.GUI.value = n+1
                self.GUI.PB_text = str(n+1)+'/1000'

            self.GUI.pb_t.Visibility = Hidden
            self.GUI.pb_c.Visibility = Hidden
        except Exception as e:
            self.GUI.pb_t.Visibility = Hidden
            self.GUI.pb_c.Visibility = Hidden
            print(e)
        
    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text



    def GetName(self):
        return "Verbinden"