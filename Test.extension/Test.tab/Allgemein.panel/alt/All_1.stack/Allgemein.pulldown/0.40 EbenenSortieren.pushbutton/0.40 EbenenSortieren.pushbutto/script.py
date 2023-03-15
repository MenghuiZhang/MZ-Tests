# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "0.40 EbenenSortieren"
__doc__ = """EbenenSortieren für Lüftungsverteilung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp

ex = Excel.ApplicationClass()

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

def ExcelData(filepath):
    DatenListe = []
    book = ex.Workbooks.Open(filepath)
    for sheet in book.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            rowlist = []
            for col in range(1, 4):
                Wert = sheet.Cells[row, col].Value2
                rowlist.append(Wert)
            DatenListe.append(rowlist)
    book.Save()
    book.Close()    
    return DatenListe

def Nr(Daten):
    Ebene_Nr = {}
    for el in Daten:
        Ebene_Nr[el[0]] = el[1]
    return Ebene_Nr

def Abk(Daten):
    Ebene_Abk = {}
    for el in Daten:
        Ebene_Abk[el[0]] = el[2]
    return Ebene_Abk

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

excelPath = rpw.ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\Ebenen_Nr_Abk.xlsx")
Ebene_Daten = ExcelData(excelPath)
EbeneNummer = Nr(Ebene_Daten)
EbeneKurz = Abk(Ebene_Daten)

t = Transaction(doc,"Ebenen_Nr_Abk")
t.Start()

with forms.ProgressBar(title='{value}/{max_value} MEP_Räume',
                               cancellable=True, step=10) as pb:
    n = 0
    for item in spaces_collector:
        if pb.cancelled:
            script.exit()
        n += 1
        pb.update_progress(n,len(spaces))

        level_Name = item.Level.Name
        nummer = EbeneNummer[level_Name]
        name = EbeneKurz[level_Name]
        item.LookupParameter("IGF_RLT_Verteilung_EbenenSortieren").Set(nummer)
        item.LookupParameter("IGF_RLT_Verteilung_EbenenName").Set(name)
t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
