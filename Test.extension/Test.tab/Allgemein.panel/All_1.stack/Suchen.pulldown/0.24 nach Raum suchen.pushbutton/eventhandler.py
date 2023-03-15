# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 

def IFC_Filter(ifc):
    param_equality=DB.FilterStringEquals()
    param_id = DB.ElementId(DB.BuiltInParameter.ROOM_NUMBER)
    param_prov=DB.ParameterValueProvider(param_id)
    param_value_rule=DB.FilterStringRule(param_prov,param_equality,ifc,True)
    param_filter = DB.ElementParameterFilter(param_value_rule)
    return param_filter


class ANZEIGEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.GUI.guid.Text:
            TaskDialog.Show('Fehler','keine Raumnummer') 
            return
        try:
            elemid = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WherePasses(IFC_Filter(self.GUI.guid.Text)).WhereElementIsNotElementType().ToElementIds()
        except:
            TaskDialog.Show('Fehler','ungültige Raumnummer') 
            return
        
        if len(elemid) == 0:
            TaskDialog.Show('Fehler','ungültige Raumnummer') 
            return
            
        uidoc.Selection.SetElementIds(elemid)
        uidoc.ShowElements(elemid)


    def GetName(self):
        return "Element anzeigen"

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
      
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.GUI.guid.Text:
            TaskDialog.Show('Fehler','keine Raumnummer') 
            return
        try:
            elemid = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WherePasses(IFC_Filter(self.GUI.guid.Text)).WhereElementIsNotElementType().ToElementIds()
        except:
            TaskDialog.Show('Fehler','ungültige Raumnummer') 
            return
        
        if len(elemid) == 0:
            TaskDialog.Show('Fehler','ungültige Raumnummer') 
            return
 
            
        uidoc.Selection.SetElementIds(elemid)

    def GetName(self):
        return "Element auswählen"