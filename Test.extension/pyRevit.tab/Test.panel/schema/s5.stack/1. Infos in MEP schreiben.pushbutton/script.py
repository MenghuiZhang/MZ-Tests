# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import script
import os
from IGF_forms import ExcelSuche
from rpw import revit,DB
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal
# from excel import get_cell_Daten

__title__ = "1. MEP-Räume"
__doc__ = """

es gilt nur für die MEP-Räume in akuelle Ansiche des Schema-Modells.

Infos in MEP-Räume von TGA-Modell in Schema schreiben.

Vorgehensweise:

1. Ein MEP-Räume-Bauteilliste exportieren. sieht R:\pyRevit\10_Verknüpfung\02_Schema\Vorlage MEP-Räume
2. Exporteirte Excel anpassen (entsprechenden Parametername in 1. Zeile anpassen und dann die 2. Zeile löschen)
3. Entsprechende Ansicht (Bsp. Schema-Ansicht oder Bauteilliste) in Schema-Modell öffnen und das Skript durchlaufen.

[2022.08.02]
Version: 1.0
"""

__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

logger = script.get_logger()
exapplication = ex.ApplicationClass()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp
active_view = uidoc.ActiveView
revitversion = app.VersionNumber

# if revitversion == '2020':
#     import excel._NPOI_2020 as _NPOI
    

# else:
#     import excel._NPOI_2022 as _NPOI

projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number

config = script.get_config('Schema-MEP -' + projectinfo)
adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.Title = 'MEP-Räume Bauteilliste'
ExcelWPF.ShowDialog()
try:
    config.adresse = ExcelWPF.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel ausgewählt!')
    script.exit()

# Parameter_Dict = {}
# fs = _NPOI.FileStream(ExcelWPF.Adresse.Text,_NPOI.FileMode.Open,_NPOI.FileAccess.Read)
# book1 = _NPOI.np.XSSF.UserModel.XSSFWorkbook(fs)
# Parameter_Liste = []

# try:
#     sheet = book1.GetSheetAt(0)
#     rows = sheet.LastRowNum
#     cols = sheet.GetRow(0).LastCellNum-1

#     for col in range(1, cols):
#         para = get_cell_Daten(sheet,0,col)
#         if para == '' or  para == None:
#             break
#         Parameter_Liste.append(para)

#     for row in range(1, rows):
#         liste = []
#         nummer = get_cell_Daten(sheet,row,0)
#         if nummer == '' or  nummer == None:
#             continue
#         for col in range(1, len(Parameter_Liste)):
#             value = get_cell_Daten(sheet,row,col)
#             liste.append(value)
#         Parameter_Dict[nummer] = liste
# except Exception as e:
#     print(e)
Parameter_Dict = {}
book1 = exapplication.Workbooks.Open(ExcelWPF.Adresse.Text)
Parameter_Liste = []
try:
    for sheet in book1.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        cols = sheet.UsedRange.Columns.Count
        Bauteilnummer = sheet.Cells[1, 1].Value2
        for col in range(2, cols + 1):
            para = sheet.Cells[1, col].Value2
            if para == '' or  para == None:
                break
            Parameter_Liste.append(para)

        for row in range(2, rows + 1):
            liste = []
            nummer = sheet.Cells[row, 1].Value2
            if nummer == '' or  nummer == None:
                continue
            for col in range(2, len(Parameter_Liste) + 2):
                value = sheet.Cells[row, col].Value2
                liste.append(value)
            Parameter_Dict[nummer] = liste
    book1.Save()
    book1.Close()
except Exception as e:
    logger.error(e)
    Marshal.FinalReleaseComObject(sheet)
    Marshal.FinalReleaseComObject(book1)
    exapplication.Quit()
    Marshal.FinalReleaseComObject(exapplication)
    script.exit()


mep = DB.FilteredElementCollector(doc,active_view.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).ToElements()
t =DB.Transaction(doc,'Infos. in MEP schreiben')
t.Start()
for el in mep:
    if el.Number in Parameter_Dict.keys():
        for n,para in enumerate(Parameter_Liste):
            param = el.LookupParameter(para)
            if param.StorageType.ToString() == 'String':
                param.Set(str(Parameter_Dict[el.Number][n]))
            elif param.StorageType.ToString() == 'Integer':
                param.Set(int(round(float(Parameter_Dict[el.Number][n]))))
            else:
                param.SetValueString(str(Parameter_Dict[el.Number][n]))
    else:
        logger.error('Raum {} existiert nur in Schema-Modell. Bitte einmal prüfen, ob die Raumnummer in Schema-Modell richtig eingetragen wird.'.format(el.Number))
t.Commit()