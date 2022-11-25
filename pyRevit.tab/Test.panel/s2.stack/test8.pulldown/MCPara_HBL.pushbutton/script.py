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


__title__ = "MagiCAD"
__doc__ = """FamilyPara(IGF_Material) erstellen. Familien in folgenden Kategorien: Luftkanalzubehör, Luftdurchlässe, Luftkanalformteile,
HLS-Bauteile, Rohrzubehör, Rohrformteile"""
__author__ = "Menghui Zhang"

# logger = script.get_logger()
# output = script.get_output()

# uidoc = rpw.revit.uidoc
# doc = rpw.revit.doc
app = rpw.revit.app
# uiapp = rpw.revit.uiapp


# Family = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()][0].Symbol.Family

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
        overwriteParameterValues = True
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source = FamilySource.Family
        overwriteParameterValues = True
        return True

def AddFamilyParam():
    trans1 = TransactionManager.Instance
    trans1.ForceCloseTransaction()
    # el = Familydaten

    # fdoc = doc.EditFamily(el)
    m_familyMgr = rpw.revit.doc.FamilyManager


    try:
        trans1.EnsureInTransaction(rpw.revit.doc)

        for defi in list_definition:
            try:

                para = m_familyMgr.AddParameter(defi, BuiltInParameterGroup.PG_GEOMETRY, True)
                if para.Definition.Name == 'MC Height Instance':
                    try:
                        m_familyMgr.SetFormula(para,'Höhe')
                    except:
                        try:m_familyMgr.SetFormula(para,'Height')
                        except:
                            try:m_familyMgr.SetFormula(para,'H')
                            except:pass
                elif para.Definition.Name == 'MC Length Instance':
                    try:
                        m_familyMgr.SetFormula(para,'Länge')
                    except:
                        try:m_familyMgr.SetFormula(para,'Length')
                        except:
                            try:m_familyMgr.SetFormula(para,'L')
                            except:pass
                elif para.Definition.Name == 'MC Width Instance':
                    try:
                        m_familyMgr.SetFormula(para,'Breite')
                    except:
                        try:m_familyMgr.SetFormula(para,'Width')
                        except:
                            try:m_familyMgr.SetFormula(para,'B')
                            except:pass
                elif para.Definition.Name == 'MC Diameter Instance':
                    try:
                        m_familyMgr.SetFormula(para,2 * 'Radius')
                    except:
                        try:m_familyMgr.SetFormula(para,'Durchmesser Luftkanal')
                        except:
                            try:m_familyMgr.SetFormula(para,'D')
                            except:
                                try:m_familyMgr.SetFormula(para,'Durchmesser')
                                except:pass
            except:pass
       # fdoc.Regenerate()
        trans1.ForceCloseTransaction()
        # logger.info('Parameter von Familie {} wurde erstellt'.format(el.Name))
    except Exception as e:
        # logger.error(el.Name +': '+e.ToString())
        trans1.ForceCloseTransaction()

    try:
        rpw.revit.doc.SaveAs('R:\\Familien\\430_L\\'+rpw.revit.doc.Title)
    except Exception as e:
        print(e)
    # LoadFamily(el,doc,el.Name)

def LoadFamily(family,Prodoc,famName):
    famdoc = Prodoc.EditFamily(family)
    trans2 = TransactionManager.Instance
    trans2.EnsureInTransaction(Prodoc)
    for symbol in family.GetFamilySymbolIds():
        if not Prodoc.GetElement(symbol).IsActive:
            Prodoc.GetElement(symbol).Activate()
        break
    trans2.ForceCloseTransaction()
    try:
        famdoc.LoadFamily(Prodoc, FamilyLoadOptions())
        logger.info('Familie {} wurde geladen'.format(famName))
    except Exception as ex:
        logger.error(ex.ToString())
    famdoc.Close(True)


AddFamilyParam()