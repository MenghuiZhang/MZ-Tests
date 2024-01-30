# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,ExternalEvent,TaskDialogCommonButtons 

class ItemTemplateEventHandler(IExternalEventHandler):
    def __init__(self):
        self.name = ''
        self.GUI = None
        self.ExcuteApp = None
    
    def Execute(self,uiapp):
        try:
            self.ExcuteApp(uiapp)
        except Exception as e:
            self.GUI.IsEnabled = True
            TaskDialog.Show('Fehler',e.ToString())

    def GetName(self):
        return self.name