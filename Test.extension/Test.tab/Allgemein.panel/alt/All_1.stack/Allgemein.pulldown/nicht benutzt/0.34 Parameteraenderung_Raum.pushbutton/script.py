# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "9.34 Parameteränderung_Raum"
__doc__ = """Parameteränderung in MEP Raum"""
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
Alt = rpw.ui.forms.TextInput('Alte Parameter: ', default = "Alte Parameter")

table = []
Neu = '0'
while Alt != 'Alte Parameter':
    t = Transaction(doc, "Werteschreiben")
    t.Start()
    Neu = rpw.ui.forms.TextInput('Neue Parameter: ', default = "Neue Parameter")

    for Space in spaces_collector:
        name = get_value(Space.LookupParameter('Name'))
        nummer = get_value(Space.LookupParameter('Nummer'))
        Werte = get_value(Space.LookupParameter(Alt))
        Wert_alt = get_value(Space.LookupParameter(Neu))
        Space.LookupParameter(Neu).Set(Werte)
        Wert_neu = get_value(Space.LookupParameter(Neu))
        table.append([nummer,name,Werte,Wert_alt,Wert_neu])

    t.Commit()

    #  Sortieren nach Raumnummer
    table.sort()

    output.print_table(
        table_data=table,
        columns=['Nummer', 'Name', Alt , Neu+'-Alt',  Neu+'-Neu']
    )

    Alt = rpw.ui.forms.TextInput('Alte Parameter: ', default = "Alte Parameter")



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
