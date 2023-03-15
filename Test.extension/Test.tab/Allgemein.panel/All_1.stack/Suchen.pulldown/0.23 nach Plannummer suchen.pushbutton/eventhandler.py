# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 

def IFC_Filter(ifc):
    param_equality=DB.FilterStringEquals()
    param_id = DB.ElementId(DB.BuiltInParameter.SHEET_NUMBER)
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
            TaskDialog.Show('Fehler','keine Plannummer') 
            return
        try:
            elemid = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WherePasses(IFC_Filter(self.GUI.guid.Text)).WhereElementIsNotElementType().ToElementIds()
        except:
            TaskDialog.Show('Fehler','ungültige Plannummer') 
            return
        if len(elemid) == 0:
            TaskDialog.Show('Fehler','falsche Plannummer') 
            return
        elif len(elemid) > 1:
            TaskDialog.Show('Fehler','mehre Pläne gefunden') 
            return
            
        uidoc.ActiveView = doc.GetElement(elemid[0])


    def GetName(self):
        return "Element anzeigen"

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.GUI = None      
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.GUI.guid.Text:
            TaskDialog.Show('Fehler','keine Plannummer') 
            return
        try:
            
            elemid = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WherePasses(IFC_Filter(self.GUI.guid.Text)).WhereElementIsNotElementType().ToElementIds()
        except:
            TaskDialog.Show('Fehler','ungültige Plannummer') 
            return
        if len(elemid) == 0:
            TaskDialog.Show('Fehler','falsche Plannummer') 
            return
        elif len(elemid) > 1:
            TaskDialog.Show('Fehler','mehre Pläne gefunden') 
            return
            
        uidoc.Selection.SetElementIds(elemid)

    def GetName(self):
        return "Element auswählen"