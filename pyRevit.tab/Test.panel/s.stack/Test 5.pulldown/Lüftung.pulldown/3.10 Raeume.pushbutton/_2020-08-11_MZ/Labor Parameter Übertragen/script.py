# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "1.Übertragen"
__doc__ = """Luftmengenberechnung"""
__author__ = "MZ"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

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

for S in spaces_collector:
    t = Transaction(doc,'Werte Übertragen')
    t.Start()
    LaborSumme = get_value(S.LookupParameter('TGA_RLT_AbluftSummeLabor'))
    LaborNeu = S.LookupParameter('IGF_RLT_AbluftminSummeLabor')
    LaborNeu.SetValueString(str(LaborSumme))
    t.Commit()



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
