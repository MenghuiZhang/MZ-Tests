# coding: utf8
import os
from pyrevit import revit, UI, DB
from pyrevit import script


__title__ = "HK Typ duplizieren"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

DICT_MEP_NUMMER_NAME  = {}

elid = uidoc.Selection.GetElementIds()[0]
t = DB.Transaction(doc,'Copy')
t.Start()
el = doc.GetElement(elid)
name = ['Kermi Therm X2 Profil-V Hygiene Type10 400x900_15_rechts',
'Kermi Therm X2 Profil-V Hygiene Type10 400x900_15_rechts_EHV',
'Kermi Therm X2 Profil-V Hygiene Type10 400x900_15_links',
'Kermi Therm X2 Profil-V Hygiene Type10 400x900_15_links_EHV'
]
# _name = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
for n in name:
    el.Duplicate(n)
t.Commit()
