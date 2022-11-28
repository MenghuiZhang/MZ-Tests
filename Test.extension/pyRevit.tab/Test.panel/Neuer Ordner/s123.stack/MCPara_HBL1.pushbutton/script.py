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


__title__ = "MagiCAD_Save"
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


def AddFamilyParam(Family):
    trans1 = TransactionManager.Instance
    trans1.ForceCloseTransaction()
    elems = Family
    for el in elems:
        fdoc = doc.EditFamily(el)
        m_familyMgr = fdoc.FamilyManager


        try:
            trans1.EnsureInTransaction(fdoc)

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
        if fdoc.Title.find('.rfa') != -1:
            fdoc.SaveAs('R:\\Familien\\430_L\\'+fdoc.Title)
        else:
            fdoc.SaveAs('R:\\Familien\\430_L\\'+fdoc.Title+'.rfa')
    except Exception as e:
        print(e)

AddFamilyParam(Families)  # rechtickig