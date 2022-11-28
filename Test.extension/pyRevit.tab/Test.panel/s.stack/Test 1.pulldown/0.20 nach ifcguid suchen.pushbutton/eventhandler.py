# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 



# def Raumnummer(ifc):
#     param_equality=DB.FilterStringContains()
#     param_id = DB.ElementId(DB.BuiltInParameter.ROOM_NUMBER)
#     param_prov=DB.ParameterValueProvider(param_id)
#     param_value_rule=DB.FilterStringRule(param_prov,param_equality,ifc,True)
#     param_filter = DB.ElementParameterFilter(param_value_rule)
#     return param_filter


class ANZEIGEN(IExternalEventHandler):
    def __init__(self):
        self.Guiclass = None
        # self.ifcfilter = Raumnummer
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        auswahl = doc.GetElement(uidoc.Selection.GetElementIds()[0])
        l = auswahl.Location.Point
        try:
            s = doc.GetSpaceAtPoint(l,doc.GetElement(auswahl.CreatedPhaseId))
            self.Guiclass.ifcguid.Text = s.Number

        except:pass

        # if not self.ifcguid:
        #     TaskDialog.Show('Fehler','keine Raumnr.') 
        #     return
        # try:
        #     elemid = DB.FilteredElementCollector(doc).WherePasses(self.ifcfilter(self.ifcguid)).WhereElementIsNotElementType().ToElementIds()[0]
        #     elem = doc.GetElement(elemid)
        # except:
        #     TaskDialog.Show('Fehler','ungültige Raumnr.') 
        #     return
        # if not elem:
        #     TaskDialog.Show('Fehler','falsche Raumnr.') 
        #     return
            
        # sel = uidoc.Selection.GetElementIds()
        # sel.Clear()
        # sel.Add(elem.Id)
        # uidoc.Selection.SetElementIds(sel)
        # uidoc.ShowElements(elem)


    def GetName(self):
        return "Element anzeigen"

class SELECT(IExternalEventHandler):
    def __init__(self):
        pass
        # self.ifcguid = None
        # self.Raumnummer = Raumnummer
      
        
    def Execute(self,app):
        return
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.ifcguid:
            TaskDialog.Show('Fehler','keine Raumnr.') 
            return
        try:
            elemid = DB.FilteredElementCollector(doc).WherePasses(self.ifcfilter(self.ifcguid)).WhereElementIsNotElementType().ToElementIds()[0]
            elem = doc.GetElement(elemid)
        except:
            TaskDialog.Show('Fehler','ungültige Raumnr.') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche Raumnr.') 
            return
            
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(elem.Id)
        uidoc.Selection.SetElementIds(sel)

    def GetName(self):
        return "Element auswählen"