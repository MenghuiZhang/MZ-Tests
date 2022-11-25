# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr

start = time.time()

Zustand = time.strftime("%d.%m.%Y", time.localtime())

__title__ = "1.02 Raumzustand_02"
__doc__ = """Raumzustand_02"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()


def get_value(param):
    value = revit.query.get_param_value(param)
    try:
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(
            value,
            unit)
    except Exception as e:
        pass
    return value

def MEP_Raum(Familie,Familie_Id):
    Raum_Daten = []
    with forms.ProgressBar(title='{value}/{max_value} Raumzustand übernehmen',
                           cancellable=True, step=10) as pb:
        t = Transaction(doc, "Raumzustand übernehmen")
        t.Start()
        n = 0
        for raum in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))
            Nummer = get_value(raum.LookupParameter('Nummer'))
            Name = get_value(raum.LookupParameter('Name'))
            Flaeche = round(get_value(raum.LookupParameter('Fläche')),3)
            Umfang = round(get_value(raum.LookupParameter('Umfang')),1)
            Hoehe = round(get_value(raum.LookupParameter('Lichte Höhe')),1)
            Volumen = round(get_value(raum.LookupParameter('Volumen')),3)
            if raum.LookupParameter('IGF_A_Raumnummer_Vorher02'):
                raum.LookupParameter('IGF_A_Raumnummer_Vorher02').Set(Nummer)
            if raum.LookupParameter('IGF_A_Raumname_Vorher02'):
                raum.LookupParameter('IGF_A_Raumname_Vorher02').Set(Name)
            if raum.LookupParameter('IGF_A_Fläche_Vorher02'):
                raum.LookupParameter('IGF_A_Fläche_Vorher02').SetValueString(str(Flaeche))
            if raum.LookupParameter('IGF_A_Volumen_Vorher02'):
                raum.LookupParameter('IGF_A_Volumen_Vorher02').SetValueString(str(Volumen))
            if raum.LookupParameter('IGF_A_Umfang_Vorher02'):
                raum.LookupParameter('IGF_A_Umfang_Vorher02').SetValueString(str(Umfang))
            if raum.LookupParameter('IGF_A_LichteHöhe_Vorher02'):
                raum.LookupParameter('IGF_A_LichteHöhe_Vorher02').SetValueString(str(Hoehe))
            if raum.LookupParameter('IGF_A_Datum_Vorher02'):
                raum.LookupParameter('IGF_A_Datum_Vorher02').Set(Zustand)
        t.Commit()

MEP_Raum(spaces_collector,spaces)

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
