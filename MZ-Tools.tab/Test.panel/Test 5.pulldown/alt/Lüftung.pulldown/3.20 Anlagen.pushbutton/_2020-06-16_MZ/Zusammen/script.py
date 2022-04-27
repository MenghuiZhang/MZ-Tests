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
logger.info("{} Luftkanalsystemen ausgewählt".format(len(systemen)))

if not spaces and not systemen:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    logger.error("Keine Duct Systemen in aktuellem Projekt gefunden")
    script.exit()


class MEPRaum:
    Zu_Anl_Nr_Liste = []
    Ab_Anl_Nr_Liste = []
    Ab_24h_Anl_Nr_Liste = []

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
            ['Zu_Anl_Nr', 'TGA_RLT_AnlagenNrZuluft'],
            ['Ab_Anl_Nr', 'TGA_RLT_AnlagenNrAbluft'],
            ['Ab_Anl_Nr_24h', 'TGA_RLT_AnlagenNr24hAbluft'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        # Anlagenliste erstellen
        if self.Zu_Anl_Nr is not None and self.Zu_Anl_Nr is not 0:
            self.Zu_Anl_Nr_Liste.append(self.Zu_Anl_Nr)

        if self.Ab_Anl_Nr is not None and self.Ab_Anl_Nr is not 0:
            self.Ab_Anl_Nr_Liste.append(self.Ab_Anl_Nr)

        if self.Ab_Anl_Nr_24h is not None and self.Ab_Anl_Nr_24h is not 0:
            self.Ab_24h_Anl_Nr_Liste.append(self.Ab_Anl_Nr_24h)


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
            self.Zu_Anl_Nr,
            self.Ab_Anl_Nr,
            self.Ab_Anl_Nr_24h,
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)

class DuctSystem:
    def __init__(self, element_id):
        """
        Definiert DuctSystem Klasse mit allen object properties für die
        Anlagenberechnung.
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
        logger.info("{}".format(self.name))


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
        return "DuctSystem({})".format(self.element_id)

    def __str__(self):
        return '{}'.format(self.name)



table_data = []
mepraum_liste = []




with forms.ProgressBar(title='{value}/{max_value} Anlagen_Liste Erstellen',
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())

table_data1 = []
Systemen_liste = []
with forms.ProgressBar(title='{value}/{max_value} Anlagenberechnung',
                       cancellable=True, step=10) as pb1:

    for n, System_id in enumerate(systemen):
        if pb1.cancelled:
            script.exit()

        pb1.update_progress(n + 1, len(systemen))
        ductsystem = DuctSystem(System_id)

        Systemen_liste.append(ductsystem)
        table_data1.append(ductsystem.table_row())


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

Zuluft_Liste = []
i = 0
zu = set(MEPRaum.Zu_Anl_Nr_Liste)
Zu = list(zu)
while i < len(Zu):
    Zuluft_Liste.append(0)
    i += 1

Abluft_Liste = []
j = 0
ab = set(MEPRaum.Ab_Anl_Nr_Liste)
Ab = list(ab)
while j < len(Ab):
    Abluft_Liste.append(0)
    j += 1

Abluft_24h_Liste = []
k = 0
ab_24h = set(MEPRaum.Ab_24h_Anl_Nr_Liste)
Ab_24h = list(ab_24h)
while k < len(Ab_24h):
    Abluft_24h_Liste.append(0)
    k += 1

for space in spaces_collector:
    Zul_Nr = get_value(space.LookupParameter('TGA_RLT_AnlagenNrZuluft'))
    Zul = get_value(space.LookupParameter('TGA_RLT_ZuluftminRaum'))
    Abl_Nr = get_value(space.LookupParameter('TGA_RLT_AnlagenNrAbluft'))
    Abl = get_value(space.LookupParameter('TGA_RLT_AbluftminRaum'))
    Abl_24h_Nr = get_value(space.LookupParameter('TGA_RLT_AnlagenNr24hAbluft'))
    Abl_24h = get_value(space.LookupParameter('TGA_RLT_AbluftSumme24h'))
    for ii in range(len(Zu)):
        if Zul_Nr == Zu[ii]:
            Zuluft_Liste[ii] = Zul + Zuluft_Liste[ii]
    for jj in range(len(Ab)):
        if Abl_Nr == Ab[jj]:
            Abluft_Liste[jj] = Abl + Abluft_Liste[jj]
    for kk in range(len(Ab_24h)):
        if Abl_24h_Nr == Ab_24h[kk]:
            Abluft_24h_Liste[kk] = Abl_24h + Abluft_24h_Liste[kk]





output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=[ 'Name', 'Nummer','Zu_Anl_Nr', 'Ab_Anl_Nr', 'Ab_Anl_Nr_24h' ]
)
output.print_table(
    table_data=table_data1,
    title="Luftmengen",
    columns=[ 'Name','Ger_Nr', 'Ger_Anzahl', 'Anl_Nr', 'Anl_Pro_Anz' ]
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} ZuluftAnlagen".format(len(Zu)))
logger.info("{} ZuluftAnlagen".format(Zu))
logger.info("{} Anlagen_Zuluft".format(Zuluft_Liste))
logger.info("{} AbluftAnlagen".format(len(Ab)))
logger.info("{} AbluftAnlagen".format(Ab))
logger.info("{} Anlagen_Abluft".format(Abluft_Liste))
logger.info("{} Abluft24hAnlagen".format(len(Ab_24h)))
logger.info("{} Abluft24hAnlagen".format(Ab_24h))
logger.info("{} Anlagen_Abluft_24h".format(Abluft_24h_Liste))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
