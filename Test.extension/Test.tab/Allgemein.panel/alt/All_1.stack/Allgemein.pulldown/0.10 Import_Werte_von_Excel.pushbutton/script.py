# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import xlrd
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "0.10 import"
__doc__ = """den Wert in Excel lesen und ihn ins entsprechende Parameter von MEP Raum schreiben"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()


from pyIGF_logInfo import getlog
getlog(__title__)



Col = []
path = rpw.ui.forms.TextInput('Excel Path: ', default = 'C:\\Users\\werksstudent\\Desktop\\Test.xlsx')
if path != 'C:\\Users\\werksstudent\\Desktop\\Test.xlsx':
    book = xlrd.open_workbook(path)
    logger.info('Anzahl des Sheets: {}'.format(book.nsheets))
    logger.info('Namen des Sheets: {}'.format(book.sheet_names()))

    sheet_name = rpw.ui.forms.TextInput('Sheet: ', default = 'Tabelle0')
    if sheet_name != 'Tabelle0':
        sheet = book.sheet_by_name(sheet_name)
        rows = sheet.nrows
        cols = sheet.ncols
        print(rows)
        print(cols)

        with forms.ProgressBar(title='{value}/{max_value} Daten lesen',
                               cancellable=True, step=10) as pb:
            n = 0

            for row in range(rows):
                if pb.cancelled:
                    script.exit()
                n += 1
                pb.update_progress(n, rows)

                Row = []
                for col in range(cols):
                    cell = sheet.cell_value(row,col)
                    Row.append(cell)
                print(Row)
                Col.append(Row)


    else:
        script.exit()

    output.print_table(
        table_data=Col,
        title="Daten aus Excel",
        columns=Col[0]
    )
else:
    script.exit()

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




para_name = rpw.ui.forms.TextInput('Parameter: ', default = 'Name')

while para_name != 'Name':
    index = rpw.ui.forms.TextInput('Spaltennummer: ', default = '1')
    Index = int(index) - 1
    t = Transaction(doc, "Werteschreiben")
    t.Start()
    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb1:
        n_1 = 0
        for space in spaces_collector:
            if pb1.cancelled:
                script.exit()
            n_1 += 1
            pb1.update_progress(n_1, len(spaces))

            nummer = get_value(space.LookupParameter('Nummer'))
            para = space.LookupParameter(para_name)
            for ROW in Col:
                if str(nummer) == str(ROW[0]):
                    para.Set(ROW[Index])
                else:
                    pass
    t.Commit()
    para_name = rpw.ui.forms.TextInput('Parameter: ', default = 'Name')



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))