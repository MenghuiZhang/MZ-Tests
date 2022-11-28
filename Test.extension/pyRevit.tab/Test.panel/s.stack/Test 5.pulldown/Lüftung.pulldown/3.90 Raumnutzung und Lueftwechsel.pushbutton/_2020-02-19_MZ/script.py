# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import time
from pyrevit import script, forms
from rpw import *
import System


start = time.time()

__title__ = "3.02 Raumnutzung und Lüftungswechsel"
__doc__ = """Raumnutzung und Lüftungswechsel"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
my_config = script.get_config()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app

exapp = Excel.ApplicationClass()

spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
excelPath_default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\Raumnutzung und Lüftwechsel.xlsx"


def ExcelLesen(inExcelPath):
    outDaten = {}
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            zelleDaten = []
            name = sheet.UsedRange.Cells[row, 1].Value2
            for col in [2,3,4]:
                werte = sheet.UsedRange.Cells[row, col].Value2
                zelleDaten.append(werte)
            outDaten[name] = zelleDaten
    book.Save()
    book.Close()
    return outDaten

try:
    excelPath_default = my_config.excelPath_.decode('utf-8')
except:
    pass

excelPath = ui.forms.TextInput('Excel: ', default = excelPath_default)
excelDaten = ExcelLesen(excelPath)
for i in excelDaten:
    print(i,excelDaten[i])
my_config.excelPath_ = excelPath.encode('utf-8')
script.save_config()

if forms.alert("Raumnutzung und Lüftungswechsel überschreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="Raumnutzung und Lüftungswechsel", cancellable=True, step=1) as pb:
        n = 0
        trans = DB.Transaction(doc,"MEP-Raum-Beschriftung")
        trans.Start()
        for space in spaces_collector:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(list(spaces_collector)))

            isSchacht = space.LookupParameter("TGA_RLT_InstallationsSchacht").AsInteger()
            name = space.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
            nummer = space.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()
            if isSchacht == "1":
                try:
                    if not space.LookupParameter("TGA_RLT_VolumenstromProNummer").AsDouble():
                        space.LookupParameter("TGA_RLT_VolumenstromProNummer").Set(9)
                        space.LookupParameter("TGA_RLT_VolumenstromProFaktor").SetValueString("keine")
                except Exception as e:
                    logger.error("Raumnummer {}, Fehler {}".format(nummer,e))
            else:
                namelist1 = name.split(" ")
                WC_TH = name[:2]
                namelist2 = name.split("/")
                namelist = namelist1 + namelist2
                namelist.append(WC_TH)
                for i in namelist:
                    if i in excelDaten.keys():
                        daten = excelDaten[i]
                        try:
                            if not space.LookupParameter("TGA_RLT_VolumenstromProNummer").AsDouble():
                                space.LookupParameter("TGA_RLT_VolumenstromProNummer").Set(daten[1])
                                space.LookupParameter("TGA_RLT_VolumenstromProFaktor").SetValueString(str(daten[2]))
                        except Exception as e:
                            logger.error("Raumnummer {}, Fehler {}".format(nummer,e))
                        break
        trans.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
