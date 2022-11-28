# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import clr

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
    .OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystem))
systemen = System_collector.ToElementIds()


logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
logger.info("{} Duct Systemen ausgewählt".format(len(systemen)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()

if not systemen:
    logger.error("Keine Duct Systemen in aktuellem Projekt gefunden")
    script.exit()


class DuctSystem:
    def __init__(self, element_id):
        """
        Definiert DuctSystem Klasse mit allen object properties für die
        Anlagen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Systemname'],
            ['Ger_Nr', 'TGA_RLT_AnlagenGeräteNr'],
            ['Ger_Anzahl', 'TGA_RLT_AnlagenGeräteAnzahl'],
            ['Anl_Nr', 'TGA_RLT_AnlagenNr'],
            ['Anl_Pro_Anz', 'TGA_RLT_AnlagenProzentualAnzahl'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
#        logger.info("{}".format(self.nummer))



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
            self.name,
            self.Ger_Nr,
            self.Ger_Anzahl,
            self.Anl_Nr,
            self.Anl_Pro_Anz,
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

#    def __str__(self):
#        return '{}\t{}'.format(self.nummer, self.name)


table_data = []
Systemen_liste = []
with forms.ProgressBar(title='{value}/{max_value} Anlagenberechnung',
                       cancellable=True, step=10) as pb:

    for n, System_id in enumerate(systemen):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(systemen))
        ductsystem = DuctSystem(System_id)

        Systemen_liste.append(ductsystem)
        table_data.append(DuctSystem.table_row())


output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=[ 'Name','Ger_Nr', 'Ger_Anzahl', 'Anl_Nr', 'Anl_Pro_Anz' ]#'Nummer','AbluftSumme24h', 'Zuluft', 'Abluft', 'Zu_Anl_Nr',
            # 'Ab_Anl_Nr', 'Ab_Anl_Nr_24h' ,'Anl_Ger_Nr'       ]
#             'ABL Excel', 'Labor pyRevit', 'Labor Excel',]
)

logger.info("{} Räume in MEP_Raumliste".format(len(Systemen_liste)))

# Werte zuückschreiben + Abfrage
#if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
#    with rpw.db.Transaction("Luftwechsel berechnen"):
#        [mepraum.werte_schreiben() for mepraum in mepraum_liste]

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
