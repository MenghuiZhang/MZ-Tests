# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.Attributes import *
from Autodesk.Revit.UI import *
import struct
import rpw
import time
import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from pyIGF_logInfo import getlog

start = time.time()

__title__ = "1.50 Architektur_Bauteile_Wechsel"
__doc__ = """Architektur_Bauteile_Wechsel"""
__author__ = "Menghui Zhang"
getlog(__title__)
exapp = Excel.ApplicationClass()
logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp

Walls_collector = DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
wall = Walls_collector.ToElementIds().ToArray()
WallType_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.WallType))
Floors_collector = DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType()
floor = Floors_collector.ToElementIds().ToArray()
FloorType_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.FloorType))

Liste = {}
newelement = []
element = wall + floor

for el in WallType_collector:
    name = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Liste[name] = el.Id
for el in FloorType_collector:
    name = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Liste[name] = el.Id

Walls_collector.Dispose()
WallType_collector.Dispose()
Floors_collector.Dispose()
FloorType_collector.Dispose()

def ExcelLesen(inExcelPath):
    outDict = {}
    book = exapp.Workbooks.Open(inExcelPath)

    for sheet in book.Worksheets:
        groupName = sheet.Name
        rows = sheet.UsedRange.Rows.Count
        with forms.ProgressBar(title='{value}/{max_value} Daten Lesen',cancellable=True, step=1) as pb:
            n_1 = 0
            for row in range(2, rows+1):
                n_1 += 1
                if pb.cancelled:
                    script.exit()
                    book.Save()
                    book.Close()
                pb.update_progress(n_1, rows-1)
                Ist = sheet.UsedRange.Cells[row, 1].Value2
                Soll = sheet.UsedRange.Cells[row, 4].Value2

                if Ist and Soll:
                    outDict[Ist] = Soll
        break
    book.Save()
    book.Close()

    return outDict

def BauteilWechsel(ids,Excel):
    step = 1
    # if len(ids) > 100:
    #     step = int(len(ids)/100)
    with forms.ProgressBar(title='{value}/{max_value} Bauteile',cancellable=True, step=step) as pb:
        n_1 = 0
        for id in ids:
            n_1 += 1
            if pb.cancelled:
                tran.RollBack()
                script.exit()
            pb.update_progress(n_1, len(ids))
            el = doc.GetElement(id)
            name = el.Name
            if Excel[name] in Liste.keys():
                el.ChangeTypeId(Liste[Excel[name]])

excelPath1  = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Heizlast\\Bauteilwechsel - Kopie.xlsx')
Daten = ExcelLesen(excelPath1)
for el in element:
    if doc.GetElement(el).Name in Daten.keys():
        newelement.append(el)


if forms.alert("Alle Element?", ok=False, yes=True, no=True):
    tran = Transaction(doc,'Bauteile wechseln')
    tran.Start()
    BauteilWechsel(newelement,Daten)
    tran.Commit()


if forms.alert("ausgewählte Element?", ok=False, yes=True, no=True):
    elem = []
    selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]
    for i in selection:
        try:
            if i.Category.Name == 'Wände':
                elem.append(i.Id)
            elif i.Category.Name == 'Geschossdecken':
                elem.append(i.Id)
        except:
            pass
    print(len(elem),len(Daten.keys()))
    elemen = []
    for el in elem:
        if doc.GetElement(el).Name in Daten.keys():
            elemen.append(el)

    tran = Transaction(doc,'Bauteile wechseln')
    tran.Start()
    BauteilWechsel(elemen,Daten)
    tran.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
