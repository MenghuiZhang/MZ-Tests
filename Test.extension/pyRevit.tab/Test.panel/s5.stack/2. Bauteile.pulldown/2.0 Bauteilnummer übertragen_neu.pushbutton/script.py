# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import script
import os
from IGF_forms import ExcelSuche
from rpw import revit,DB,UI
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext


__title__ = "2.0 BauteilId übertragen nach Einbauort_neu"
__doc__ = """

es gilt nur für HLS-Bauteile des Schema-Modells und nur in Bauteilliste-Ansicht.

Infos in HLS-Bauteile von TGA-Modell in Schema schreiben.

Vorgehensweise:

1. Ein HLS-Bauteile-Bauteilliste exportieren. Vorlage siehen Sie R:\pyRevit\10_Verknüpfung\02_Schema\HLS-Bauteile Nummer
2. Exporteirte Excel anpassen (entsprechenden Parametername in 1. Zeile anpassen und dann die 2. Zeile löschen)
3. Entsprechende Bauteilliste in Schema-Modell öffnen und das Skript durchlaufen lassen.

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

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp
active_view = uidoc.ActiveView
if active_view.ViewType.ToString() != 'Schedule':
    UI.TaskDialog.Show('Fehler','Bitte ein Bauteilliste öffnen!')
    script.exit()

projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number

config = script.get_config('Schema-HLSID -' + projectinfo)

adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.Title = 'HLS-Bauteilnummer Bauteilliste'
ExcelWPF.ShowDialog()
try:
    config.adresse = ExcelWPF.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel ausgewählt!')
    script.exit()

Parameter_Dict = {}
ExcelPackage.LicenseContext = LicenseContext.NonCommercial
fs = FileStream(ExcelWPF.Adresse.Text,FileMode.Open,FileAccess.Read)
book = ExcelPackage(fs)
Parameter = ''
try:
    for sheet in book.Workbook.Worksheets:
        rows = sheet.Dimension.End.Row
        Parameter = sheet.Cells[1, 2].Value

        for row in range(2, rows + 1):
            nummer = sheet.Cells[row, 1].Value
            Id = sheet.Cells[row, 2].Value
            if nummer == '' or  nummer == None:
                continue
            if nummer not in Parameter_Dict.keys():
                Parameter_Dict[nummer] = []
            Parameter_Dict[nummer].append(Id)

    book.Save()
    book.Dispose()

except Exception as e:
    logger.error(e)
    book.Save()
    book.Dispose()
    script.exit()


Bauteile = DB.FilteredElementCollector(doc,active_view.Id).ToElements()

dict_neu = {}
for el in Bauteile:
    space = el.Space[doc.GetElement(el.CreatedPhaseId)]
    if not space:
        continue
    else:
        if space.Number not in dict_neu.keys():
            dict_neu[space.Number] = []
        dict_neu[space.Number].append(el)

for el in Parameter_Dict.keys():
    if el not in dict_neu.keys():
        logger.error('Raum {} existiert nur in TGA-Modell. Bitte einmal prüfen, ob der Raum in Schema-Modell gezeichnet wird.'.format(el))

t = DB.Transaction(doc,'Bauteilnummer')
t.Start()
for el in dict_neu.keys():
    if el in Parameter_Dict.keys():
        if len(Parameter_Dict[el]) == len(dict_neu[el]):
            if len(Parameter_Dict[el]) == 1:
                params = dict_neu[el][0].LookupParameter(Parameter)
                if not params:
                    UI.TaskDialog.Show('Fehler','Parameter existiert nicht. Bitte Parameter für Bauteilnummer in Excel einmal prüfen!')
                    t.RollBack()
                    script.exit()
                params.Set(str(Parameter_Dict[el][0]))
            else:
                logger.warning("Mehr als 1 Bauteil in MEP-Raum {} eingesetzt. Bitte Bauteilnummer in die Bauteile in disen Raum manuell eintragen".format(el))
        else:
            logger.error('Anzahl der Bauteile für Raum {} passt nicht. Anzahl in TGA-Modell: {}; in Schema: {}'.format(el,len(Parameter_Dict[el]),len(dict_neu[el])))
    else:
        logger.error('Raum {} existiert nur in Schema-Modell. Bitte einmal prüfen, ob die Raumnummer in Schema-Modell richtig eingetragen wird.'.format(el))
t.Commit()