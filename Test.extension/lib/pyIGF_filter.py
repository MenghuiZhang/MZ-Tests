# coding: utf8
# from pyrevit import script, forms
# import rpw
# import time
# import Autodesk.Revit.DB
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# import System
# from System.IO import Path, File
# from pyrevit import revit

import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

doc = __revit__.ActiveUIDocument.Document

def Fam_Exemplar_Filter(fam_name = None, typ_name = None, ansicht = False):
    param_equality=DB.FilterStringEquals()
    if not fam_name:
        UI.TaskDialog.Show('Fehler','Bitte Geben Sie einen Familienamen ein.')
        return False

    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).WherePasses(fam_filter)
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).OfClass(DB.FamilyInstance).WherePasses(fam_filter)
    if typ_name:
        typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        typ_prov=DB.ParameterValueProvider(typ_id)
        typ_value_rule=DB.FilterStringRule(typ_prov,param_equality,typ_name,True)
        typ_filter = DB.ElementParameterFilter(typ_value_rule)
        coll.WherePasses(typ_filter)

    return coll
