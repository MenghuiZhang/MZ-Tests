# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from Autodesk.Revit.DB import *

start = time.time()


__title__ = "0.50 FamilyPara"
__doc__ = """FamilyPara erstellen """
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp


Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))

Luftkanalzu_daten = []
for el in Family:
    category = el.FamilyCategory.Name
    Family = el.Name
    if category in ['Luftkanalzubehör']:
        Luftkanalzu_daten.append(el)


class FamilyLoadOptions(IFamilyLoadOptions):
    def __init__(self): pass
    def OnFamilyFound(self,familyInUse, overwriteParameterValues = True): return True
    def OnSharedFamilyFound(self,familyInUse, source, overwriteParameterValues = True): return True

def LoadFamilyInDoc(Familydaten):
    for el in Familydaten:
        fdoc = doc.EditFamily(el)
        m_familyMgr = fdoc.FamilyManager
        try:
            t = Transaction(fdoc)
            t.Start('Add Family Parameter')
            isInstance = True
            pt = DB.ParameterType.Material
            pg = DB.BuiltInParameterGroup.PG_MATERIALS
            paramTw = m_familyMgr.AddParameter('IGF_Material', pg, pt, True)
            t.Commit()
        except Exception as e:
            logger.error(e)

        fdoc.LoadFamily(doc, FamilyLoadOptions())




Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))

Familydaten = []
for el in Family:
    category = el.FamilyCategory.Name
    Family = el.Name
    if category in ['Luftkanalzubehör']:
        Familydaten.append(el)
name = el.Name
fdoc = doc.EditFamily(el)
t = Transaction(fdoc)
t.Start('Add Family Parameter')
m_familyMgr = fdoc.FamilyManager
isInstance = True
pt = DB.ParameterType.Material
pg = DB.BuiltInParameterGroup.PG_MATERIALS
paramTw = m_familyMgr.AddParameter('IGF_Materialll', BuiltInParameterGroup.PG_MATERIALS, pt, isInstance)
t.Commit()
Task = UI.TaskDialog.Show('LoadFamily','Familie {} wird geladen'.format(Name))
option = uidoc.GetRevitUIFamilyLoadOptions()
familyEnd = fdoc.LoadFamily(doc, option)
# def ExcelData(filepath):
from Autodesk.Revit.DB import *
Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))

Familydaten = []
for el in Family:
    category = el.FamilyCategory.Name
    Family = el.Name
    if category in ['Luftkanalzubehör']:
        Familydaten.append(el)

class FamilyLoadOptions(IFamilyLoadOptions):
    def __init__(self): pass
    def OnFamilyFound(self,familyInUse, overwriteParameterValues = True): return True
    def OnSharedFamilyFound(self,familyInUse, source, overwriteParameterValues = True): return True

el = Familydaten[3]
print(el.Name)
fdoc = doc.EditFamily(el)
t = Transaction(fdoc)
t.Start('Add Family Parameter')
m_familyMgr = fdoc.FamilyManager
isInstance = True
pt = DB.ParameterType.Material
pg = DB.BuiltInParameterGroup.PG_MATERIALS
paramTw = m_familyMgr.AddParameter('IGF_Materiallllll', BuiltInParameterGroup.PG_MATERIALS, pt, isInstance)
t.Commit()

familyEnd = fdoc.LoadFamily(doc, FamilyLoadOptions())


#     DatenListe = []
#     book = ex.Workbooks.Open(filepath)
#     for sheet in book.Worksheets:
#         rows = sheet.UsedRange.Rows.Count
#         for row in range(2, rows + 1):
#             rowlist = []
#             for col in range(1, 4):
#                 Wert = sheet.Cells[row, col].Value2
#                 rowlist.append(Wert)
#             DatenListe.append(rowlist)
#     book.Save()
#     book.Close()
#     return DatenListe

# def Nr(Daten):
#     Ebene_Nr = {}
#     for el in Daten:
#         Ebene_Nr[el[0]] = el[1]
#     return Ebene_Nr

# def Abk(Daten):
#     Ebene_Abk = {}
#     for el in Daten:
#         Ebene_Abk[el[0]] = el[2]
#     return Ebene_Abk

# logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

# excelPath = rpw.ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\Ebenen_Nr_Abk.xlsx")
# Ebene_Daten = ExcelData(excelPath)
# EbeneNummer = Nr(Ebene_Daten)
# EbeneKurz = Abk(Ebene_Daten)

# t = Transaction(doc,"Ebenen_Nr_Abk")
# t.Start()

# with forms.ProgressBar(title='{value}/{max_value} MEP_Räume',
#                                cancellable=True, step=10) as pb:
#     n = 0
#     for item in spaces_collector:
#         if pb.cancelled:
#             script.exit()
#         n += 1
#         pb.update_progress(n,len(spaces))

#         level_Name = item.Level.Name
#         nummer = EbeneNummer[level_Name]
#         name = EbeneKurz[level_Name]
#         item.LookupParameter("IGF_RLT_Verteilung_EbenenSortieren").Set(nummer)
#         item.LookupParameter("IGF_RLT_Verteilung_EbenenName").Set(name)
# t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
