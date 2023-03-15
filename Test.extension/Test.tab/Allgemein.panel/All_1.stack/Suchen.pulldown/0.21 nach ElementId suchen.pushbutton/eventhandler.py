# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 

class ANZEIGEN(IExternalEventHandler):
    def __init__(self):
        self.elemid = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.elemid:
            TaskDialog.Show('Fehler','keine ElementId') 
            return
        try:
            elem = doc.GetElement(DB.ElementId(int(self.elemid)))
        except:
            TaskDialog.Show('Fehler','ungültige ElementId') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche ElementId') 
            return
            
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(elem.Id)
        uidoc.Selection.SetElementIds(sel)
        uidoc.ShowElements(elem)


    def GetName(self):
        return "Element anzeigen"

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.elemid = None
      
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.elemid:
            TaskDialog.Show('Fehler','keine ElementId') 
            return
        try:
            elem = doc.GetElement(DB.ElementId(int(self.elemid)))
        except:
            TaskDialog.Show('Fehler','ungültige ElementId') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche ElementId') 
            return
            
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(elem.Id)
        uidoc.Selection.SetElementIds(sel)

    def GetName(self):
        return "Element auswählen"