# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "0.31 Wert_Schreiben_Raum"
__doc__ = """Wert Schreiben"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)


# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

Para = rpw.ui.forms.TextInput('Parameter: ', default = "Parameter")

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

table = []
while Para != 'Parameter':

    t = Transaction(doc, 'Übertragen')
    t.Start()


    Wert = rpw.ui.forms.TextInput('Wert: ', default = "Wert")

    for Space in spaces_collector:
        name = get_value(Space.LookupParameter('Name'))
        nummer = get_value(Space.LookupParameter('Nummer'))
        para = Space.LookupParameter(Para)
        para.SetValueString(str(Wert))

    t.Commit()



    Para = rpw.ui.forms.TextInput('Parameter: ', default = "Parameter")


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
