# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script
import rpw
import time

start = time.time()


__title__ = "MFC RLT"
__doc__ = """ """
__author__ = "Tim Bartels"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc  # type: Document
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces = DB.FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces) \
    .ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

berechnung_nach = {
    'Fläche': 1,
    'Luftwechsel': 2,
    'manuell': 4,
    'nurZUMa': 5,
    'nurABMa': 6,
    'keine': 9
}


class MEPRaum:
    def __init__(self, element_id):

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['flaeche', 'Fläche'],
            ['volumen', 'Volumen'],
            ['personen', 'Personen'],
            ['kez', 'TGA_RLT_VolumenstromProNummer'],
            ['vol_faktor', 'TGA_RLT_VolumenstromProFaktor'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(
                self.element, revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        self.zuluft_min = self.zuluft_min_berechnen()

    def get_element_attr(self, element, param_name):
        param = element.LookupParameter(param_name)
        value = revit.query.get_param_value(param)

        # Einheiten konvertieren (Revit internal -> Projekteinheiten)
        try:
            unit = param.DisplayUnitType

            value = DB.UnitUtils.ConvertFromInternalUnits(
                value,
                unit)

        except Exception as e:
            pass

        return value

    def zuluft_min_berechnen(self):
        zuluft = 0

        logger.info("Fläche: {}".format(self.flaeche))

        if int(self.kez) == berechnung_nach['Fläche']:
            logger.info("Berechnung nach Fläche")

            zuluft = self.flaeche * self.vol_faktor

        elif int(self.kez) == berechnung_nach['Luftwechsel']:
            logger.info("Berechnung nach Luftwechsel")

            zuluft = self.volumen * self.vol_faktor

        logger.info("ZUL = {}".format(zuluft))

        return zuluft

    def table_row(self):
        return [
            self.nummer,
            self.name,
            self.zuluft_min,
            self.get_element_attr(self.element, "TGA_RLT_ZuluftminRaum")
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


table_data = []
for space_id in spaces:
    mepraum = MEPRaum(space_id)

    table_data.append(mepraum.table_row())

output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=['Nummer', 'Name', 'ZUL pyRevit', 'ZUL Excel']
)

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
