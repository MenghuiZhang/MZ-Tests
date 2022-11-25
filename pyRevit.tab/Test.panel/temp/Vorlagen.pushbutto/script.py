# coding: utf8
import clr
from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
import time
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import System

exapp = Excel.ApplicationClass()

start = time.time()

__title__ = "BIM-ID\nExport"
__doc__ = """..."""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
Excel_config = script.get_config()

system_luft = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsElementType()
system_rohr = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType()
system_elek = FilteredElementCollector(doc).OfClass(clr.GetClrType(Electrical.ElectricalSystem)).WhereElementIsElementType()

Luft = {}
Rohr = {}
Elek = {}

Gesamt = {}

excel_Adresse = ''

def read_config():
    try:
        excel_Adresse = str(Excel_config.excel)
    except:
        pass

def write_config():
    Excel_config.excel = excel_Adresse
    script.save_config()

def DatenErmitteln(Sys_coll):
    dict = {}
    for el in Sys_coll:
        sys_Name = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        Gewerke = el.LookupParameter('IGF_X_Kostengruppe').AsValueString()
        KG = el.LookupParameter('IGF_X_Gewerkkürzel').AsString()
        KN_01 = el.LookupParameter('IGF_X_Kennnummer_1').AsValueString()
        KN_02 = el.LookupParameter('IGF_X_Kennnummer_2').AsValueString()
        ID = el.LookupParameter('IGF_X_BIM-ID').AsString()
        dict[sys_Name] = [ID,Gewerke,KG,KN_01,KN_02]
    return dict

def ExcelExport(path):
    book = exapp.Workbooks.Open(path)
    Sheets = [s.Name for s in book.Worksheets]
    if not 'Luft' in Sheets:
        sheet = book.Worksheets.Add()
        sheet.Name = 'Luft'
        sheet.Cells[1,1] = 'Systemname'
        sheet.Cells[1,2] = 'Gewerkkürzrl'
        sheet.Cells[1,3] = 'Kostengruppen'
        sheet.Cells[1,4] = 'Kennnummer_1'
        sheet.Cells[1,5] = 'Kennnummer_2'
        sheet.Cells[1,6] = 'BIM-ID'

    if not 'Rohr' in Sheets:
        sheet = book.Worksheets.Add()
        sheet.Name = 'Rohr'
        sheet.Cells[1,1] = 'Systemname'
        sheet.Cells[1,2] = 'Gewerkkürzrl'
        sheet.Cells[1,3] = 'Kostengruppen'
        sheet.Cells[1,4] = 'Kennnummer_1'
        sheet.Cells[1,5] = 'Kennnummer_2'
        sheet.Cells[1,6] = 'BIM-ID'

    for sheet in book.Worksheets:
        if sheet.Name in Gesamt.Keys:
            Daten = Gesamt[sheet.Name]
            rows = sheet.UsedRange.Rows.Count
            n = 0
            for row in [2, rows+1]:
                sysname = sheet.UsedRange.Cells[row, 1].Value2
                if sysname in Daten.Keys:
                    sheet.UsedRange.Cells[row, 2] = Daten[sysname][1]
                    sheet.UsedRange.Cells[row, 3] = Daten[sysname][2]
                    sheet.UsedRange.Cells[row, 4] = Daten[sysname][3]
                    sheet.UsedRange.Cells[row, 5] = Daten[sysname][4]
                    sheet.UsedRange.Cells[row, 6] = Daten[sysname][0]
                    del Daten[sysname]
            if any(Daten.Keys):
                for name in Daten.Keys:
                    n += 1
                    sheet.Cells[rows + n,1] = name
                    sheet.Cells[rows + n,2] = Daten[name][1]
                    sheet.Cells[rows + n,3] = Daten[name][2]
                    sheet.Cells[rows + n,4] = Daten[name][3]
                    sheet.Cells[rows + n,5] = Daten[name][4]
                    sheet.Cells[rows + n,6] = Daten[name][0]
    book.Save()
    book.Close()

read_config()
if forms.alert("Neue Excel hinzufügen?", ok=False, yes=True, no=True):
    dialog = OpenFileDialog()
    dialog.Multiselect = False
    dialog.Title = "Excel-Datei mit BIM-ID"
    dialog.Filter = "Excel Dateien|*.xls;*.xlsx;*.csv"
    if dialog.ShowDialog() == DialogResult.OK:
        excel_Adresse = dialog.FileName
        write_config()

Luft = DatenErmitteln(system_luft)
Rohr = DatenErmitteln(system_rohr)
Elek = DatenErmitteln(system_elek)
Gesamt['Luft'] = Luft
Gesamt['Rohr'] = Rohr
Gesamt['Elektro'] = Elek

ExcelExport(excel_Adresse)
