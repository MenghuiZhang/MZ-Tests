# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit
import os


class Familien(object):
    def __init__(self,elem,name):
        self.elem = elem
        self.name = name

def Get_IS():
    _dict = {}
    FamilySymbols = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsElementType().ToElements()
    for el in FamilySymbols:
        famname = el.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
        if famname not in _dict.keys():
            _dict[famname] = el.Id
    return _dict

LISTE_IS = Get_IS()

class CHANGEFAMILY(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        elems = uidoc.Selection.GetElementIds()        
        t = DB.Transaction(doc,'change Bauteiltyp')
        t.Start()
        for item in elems:

            try:
                famid = self.GUI.liste_is[self.GUI.distribution.SelectedItem.ToString()]
                doc.GetElement(item).ChangeTypeId(famid)
            except:pass

        t.Commit()
        t.Dispose()


    def GetName(self):
        return "Bauteil austauschen"