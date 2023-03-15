# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 

class ANZEIGEN(IExternalEventHandler):
    def __init__(self):
        self.guid = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.guid:
            TaskDialog.Show('Fehler','keine Revit-UniqueId') 
            return
        try:
            elem = doc.GetElement(self.guid)
        except:
            TaskDialog.Show('Fehler','ungültige Revit-UniqueId') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche Revit-UniqueId') 
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
        self.guid = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.guid:
            TaskDialog.Show('Fehler','keine Revit-UniqueId') 
            return
        try:
            elem = doc.GetElement(self.guid)
        except:
            TaskDialog.Show('Fehler','ungültige Revit-UniqueId') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche Revit-UniqueId') 
            return
            
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(elem.Id)
        uidoc.Selection.SetElementIds(sel)

    def GetName(self):
        return "Element auswählen"