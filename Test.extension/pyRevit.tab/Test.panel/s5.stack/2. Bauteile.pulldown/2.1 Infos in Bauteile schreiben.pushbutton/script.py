# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import script
import os
from IGF_forms import ExcelSuche
from rpw import revit,DB,UI
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal

__title__ = "2.1 Infos in Bauteile schreiben"
__doc__ = """

es gilt nur für Bauteile (Ventile, HLS-Bauteile etc.) des Schema-Modells und nur in Bauteilliste-Ansicht.

Infos in HLS-Bauteile von TGA-Modell in Schema schreiben.

Vorgehensweise:

1. Eine Bauteilliste exportieren. Vorlage siehen Sie R:\pyRevit\10_Verknüpfung\02_Schema\Vorlage Bauteilinfos
2. Exporteirte Excel anpassen (entsprechenden Parametername in 1. Zeile anpassen und dann die 2. Zeile löschen)
3. Entsprechende Bauteilliste in Schema-Modell öffnen und das Skript durchlaufen.

[2022.08.03]
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
if active_view.ViewType.ToString() != 'Schedule':
    UI.TaskDialog.Show('Fehler','Bitte ein Bauteilliste öffnen!')
    script.exit()

projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number

config = script.get_config('Schema-HLSInfos -' + projectinfo)

adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.Title = 'Bauteilliste für Informationen der Bauteile'
ExcelWPF.ShowDialog()
try:
    config.adresse = ExcelWPF.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel ausgewählt!')
    script.exit()

Parameter_Dict = {}
book1 = exapplication.Workbooks.Open(ExcelWPF.Adresse.Text)
Parameter_Liste = []
Bauteilnummer = ''
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


Bauteile = DB.FilteredElementCollector(doc,active_view.Id).ToElements()

dict_neu = {}
for el in Bauteile:
    bauteil = el.LookupParameter(Bauteilnummer)
    if bauteil == None:
        UI.TaskDialog.Show('Fehler','Parameter für Bauteilnummer in excel ist Falsch. Bitte korrigieren!')
        script.exit()
    else:
        Id = bauteil.AsString()
        if Id not in dict_neu.keys():
            dict_neu[Id] = []
        dict_neu[Id].append(el)

for el in Parameter_Dict.keys():
    if el not in dict_neu.keys():
        logger.error('Bauteil {} existiert nur in TGA-Modell. Bitte einmal prüfen, ob der Bauteil in Schema-Modell gezeichnet wird.'.format(el))

t =DB.Transaction(doc,'Bauteilnummer')
t.Start()
for el in Bauteile:
    bauteil = el.LookupParameter(Bauteilnummer)
    if bauteil == None:
        UI.TaskDialog.Show('Fehler','Parameter für Bauteilnummer in excel ist Falsch. Bitte korrigieren!')
        t.RollBack()
        script.exit()
    Id = bauteil.AsString()
    if Id in Parameter_Dict.keys():
        for n,para in enumerate(Parameter_Liste):
            param = el.LookupParameter(para)
            if param:
                if param.StorageType.ToString() == 'String':
                    if param.IsReadOnly == False:param.Set(str(Parameter_Dict[Id][n]))
                    else:print('Parameter {} von Element {} ist schreibgeschützt.'.format(para,Id))
                else:
                    if param.IsReadOnly == False:param.SetValueString(str(Parameter_Dict[Id][n]))
                    else:print('Parameter {} von Element {} ist schreibgeschützt.'.format(para,Id))
            else:
                print('Parameter {} existiert nicht. Bitte anlegen.'.format(para))
    else:
        logger.error('Bauteil {} (ElementId {}) existiert nur in Schema-Modell. Bitte einmal prüfen, ob die Bauteilnummer in Schema-Modell richtig eingetragen wird.'.format(Id,el.Id.ToString()))
t.Commit()
