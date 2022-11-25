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


__title__ = "MagiCAD (Falilie laden)"
__doc__ = """FamilyPara(IGF_Material) erstellen. Familien in folgenden Kategorien: Luftkanalzubehör, Luftdurchlässe, Luftkanalformteile,
HLS-Bauteile, Rohrzubehör, Rohrformteile"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp


Family = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()][0].Family

list_parameter = [
'MC Length Instance',
'MC Height Instance',
'MC Width Instance',
'MC Diameter Instance'
]

list_definition = []
file = app.OpenSharedParameterFile()
group = file.Groups['MagiCAD'].Definitions
for defi in group:
    if defi.Name in list_parameter:
        list_definition.append(defi)

class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source = FamilySource.Family
        overwriteParameterValues = False
        return True

def AddFamilyParam(Familydaten):
    trans1 = TransactionManager.Instance
    trans1.ForceCloseTransaction()
    el = Familydaten

    fdoc = doc.EditFamily(el)
    m_familyMgr = fdoc.FamilyManager


    try:
        trans1.EnsureInTransaction(fdoc)

        for defi in list_definition:
            para = m_familyMgr.AddParameter(defi, BuiltInParameterGroup.PG_GEOMETRY, True)
            if para.Definition.Name == 'MC Height Instance':
                try:
                    m_familyMgr.SetFormula(para,'Höhe')
                except:
                    try:m_familyMgr.SetFormula(para,'Height')
                    except:pass
            elif para.Definition.Name == 'MC Length Instance':
                try:
                    m_familyMgr.SetFormula(para,'Länge')
                except:
                    try:m_familyMgr.SetFormula(para,'Length')
                    except:pass
            elif para.Definition.Name == 'MC Width Instance':
                try:
                    m_familyMgr.SetFormula(para,'Breite')
                except:
                    try:m_familyMgr.SetFormula(para,'Width')
                    except:pass

        fdoc.Regenerate()
        trans1.ForceCloseTransaction()
        # logger.info('Parameter von Familie {} wurde erstellt'.format(el.Name))
    except Exception as e:
        # logger.error(el.Name +': '+e.ToString())
        trans1.ForceCloseTransaction()
    trans2 = TransactionManager.Instance
    trans2.EnsureInTransaction(doc)
    for symbol in el.GetFamilySymbolIds():
        if not doc.GetElement(symbol).IsActive:
            doc.GetElement(symbol).Activate()
        break
    trans2.ForceCloseTransaction()
    try:
        fdoc.LoadFamily(doc, FamilyLoadOptions())
    except Exception as ex:
        logger.error(ex.ToString())
    fdoc.Close(True)

# def LoadFamily(family,Prodoc,famName):
#     famdoc = Prodoc.EditFamily(family)
#     trans2 = TransactionManager.Instance
#     trans2.EnsureInTransaction(Prodoc)
#     for symbol in family.GetFamilySymbolIds():
#         if not Prodoc.GetElement(symbol).IsActive:
#             Prodoc.GetElement(symbol).Activate()
#         break
#     trans2.ForceCloseTransaction()
#     try:
#         famdoc.LoadFamily(Prodoc, FamilyLoadOptions())
#         logger.info('Familie {} wurde geladen'.format(famName))
#     except Exception as ex:
#         logger.error(ex.ToString())
#     famdoc.Close(True)


AddFamilyParam(Family)