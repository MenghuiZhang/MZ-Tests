# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time

start = time.time()


__title__ = "MFC ANL"
__doc__ = """Anlagenberechnung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

# Systemen aus Projekt
System_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctSystem_Reference)
systemen = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
logger.info("{} Duct Systemen ausgewählt".format(len(systemen)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()

if not systemen:
    logger.error("Keine Duct Systemen in aktuellem Projekt gefunden")
    script.exit()


class MEPRaum:
    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['ABL_24h', 'TGA_RLT_AbluftSumme24h'],
            ['Zuluft', 'TGA_RLT_ZuluftminRaum'],
            ['Abluft', 'TGA_RLT_AbluftminRaum'],
            ['Zu_Anl_Nr', 'TGA_RLT_AnlagenNrZuluft'],
            ['Ab_Anl_Nr', 'TGA_RLT_AnlagenNrAbluft'],
            ['Ab_Anl_Nr_24h', 'TGA_RLT_AnlagenNr24hAbluft'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))



    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return self.__get_value_in_project_units(param)

    def __get_value_in_project_units(self, param):
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


    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.ABL_24h,
            self.Zuluft,
            self.Abluft,
            self.Zu_Anl_Nr,
            self.Ab_Anl_Nr,
            self.Ab_Anl_Nr_24h,
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


table_data = []
mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} Luftmengenberechnung',
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())


output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=[ 'Name', 'Nummer','AbluftSumme24h', 'Zuluft', 'Abluft', 'Zu_Anl_Nr',
              'Ab_Anl_Nr', 'Ab_Anl_Nr_24h' ,'Anl_Ger_Nr']
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))

# Werte zuückschreiben + Abfrage
#if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
#    with rpw.db.Transaction("Luftwechsel berechnen"):
#        [mepraum.werte_schreiben() for mepraum in mepraum_liste]

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
