# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import xlsxwriter
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "0.1 schreiben"
__doc__ = """den Werte der Paramete in MEP Räume lesen und ihnen in Excel schreiben"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()


uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

def get_value(param):
    """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""

    value = revit.query.get_param_value(param)

    try:
        unit = param.DisplayUnitType

        value = DB.UnitUtils.ConvertFromInternalUnits(
            value,
            unit)

    except Exception as e:
        pass

    return value

Parameter_Raum = rpw.ui.forms.TextInput('Parameter: ', default = 'Parameter_Name')
Para_liste = []
while Parameter_Raum != 'Parameter_Name':
    Para_liste.append(Parameter_Raum)
    Parameter_Raum = rpw.ui.forms.TextInput('Parameter: ', default = 'Parameter_Name')

if not any(Para_liste):
    logger.error("Kein Parameter ausgewählt")
    total = time.time() - start
    logger.info("total time: {}".format(total))
    script.exit()

Para_Projekt = []
Para_Projekt.append(Para_liste)

with forms.ProgressBar(title='{value}/{max_value} Daten lesen',
                               cancellable=True, step=10) as pb:
    n = 0

    for Space in spaces_collector:
        if pb.cancelled:
            script.exit()
        n += 1
        pb.update_progress(n, len(spaces))

        Para_Raum = []
        for i in Para_liste:
            wert = get_value(Space.LookupParameter(i))
            Para_Raum.append(wert)
        Para_Projekt.append(Para_Raum)


Col = []
path = rpw.ui.forms.TextInput('Excel Path: ', default = 'C:\\Users\\werksstudent\\Desktop\\Test.xlsx')
if path != 'C:\\Users\\werksstudent\\Desktop\\Test.xlsx':
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    with forms.ProgressBar(title='{value}/{max_value} Daten schreiben',
                                   cancellable=True, step=10) as pb1:
        a = 0

        n_1 = 0
        for row in Para_Projekt:
            if pb1.cancelled:
                script.exit()
            n_1 += 1
            pb1.update_progress(n_1, len(Para_Projekt))

            for col in range(len(row)):
                worksheet.write(a, col,  row[col])
            a += 1

    workbook.close()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
