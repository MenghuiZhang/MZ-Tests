# coding: utf8
import os
from pyrevit import revit, DB
from pyrevit import script
import xlsxwriter
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
from System.Collections.Generic import List

__title__ = "Nummervergeben DSP"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

def Nummervergeben_Schema():
    cl = uidoc.Selection.GetElementIds()
    dict_ = {}
    for el in cl:
         elem = doc.GetElement(el)
         mep = elem.Space[doc.GetElement(elem.CreatedPhaseId)].Number
         dict_.setdefault(mep,[])
         dict_[mep].append(elem)

    t = DB.Transaction(doc,'Test')
    t.Start()
    for mep in dict_.keys():
        for n,elem in enumerate(dict_[mep]):
            elem.LookupParameter('ZS_MARK').Set('')
            elem.LookupParameter('IGF_X_Bauteil_ID_Text').Set('DSP_'+mep+ '.' + str(n))
    t.Commit()
# Nummervergeben_Schema()

def Nummervergeben_Modell():
    elems = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
    cl = []
    for el in elems:
        if el.Symbol.FamilyName.find('Deckensegel') != -1:
            conns = el.MEPModel.ConnectorManager.UnusedConnectors
            if conns.Size == 0:           cl.append(el)

    dict_ = {}
    for el in cl:
         elem = el
         mep = elem.Space[doc.GetElement(elem.CreatedPhaseId)].Number
         dict_.setdefault(mep,[])
         dict_[mep].append(elem)

    t = DB.Transaction(doc,'Test')
    t.Start()
    for mep in dict_.keys():
        for n,elem in enumerate(dict_[mep]):
            elem.LookupParameter('SBI_Bauteilnummerierung').Set('DSP_'+mep+ '.' + str(n))
    t.Commit()
Nummervergeben_Modell()