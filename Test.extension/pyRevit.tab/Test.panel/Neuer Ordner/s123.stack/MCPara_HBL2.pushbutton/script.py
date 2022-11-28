# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import System
import clr
from Autodesk.Revit.DB import *
clr.AddReference("RevitServices")
from RevitServices.Transactions import TransactionManager


__title__ = "MagiCAD_Load"
__doc__ = """FamilyPara(IGF_Material) erstellen. Familien in folgenden Kategorien: Luftkanalzubehör, Luftdurchlässe, Luftkanalformteile,
HLS-Bauteile, Rohrzubehör, Rohrformteile"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp


Families = []
for el in uidoc.Selection.GetElementIds():
    elem = doc.GetElement(el).Family
    Families.append(elem)

class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source = FamilySource.Family
        overwriteParameterValues = False
        return True

def LoadFamily(Families):
    for el in Families:
        famdoc = doc.EditFamily(el)
        trans2 = TransactionManager.Instance
        trans2.EnsureInTransaction(doc)
        for symbol in el.GetFamilySymbolIds():
            if not doc.GetElement(symbol).IsActive:
                doc.GetElement(symbol).Activate()
            break
        trans2.ForceCloseTransaction()
        try:
            famdoc.LoadFamily(doc, FamilyLoadOptions())
        except Exception as ex:
            logger.error(ex.ToString())
        famdoc.Close(True)


# AddFamilyParam0()  # rund
LoadFamily(Families)  # rechtickig