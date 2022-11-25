# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr
import operator


start = time.time()


__title__ = "2.Anlagen"
__doc__ = """Anlagenberechnung"""
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

# Systemen aus Projekt
System_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystem))\
    .WhereElementIsNotElementType()

systemen = System_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
logger.info("{} Luftkanalsystemen ausgewählt".format(len(systemen)))

if not spaces and not systemen:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    logger.error("Keine Luftkanalsystemen in aktuellem Projekt gefunden")
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
            ['ZU', 'IGF_RLT_ZuluftminRaum'],
            ['AB', 'IGF_RLT_AbluftminRaum'],
            ['AB24H', 'IGF_RLT_AbluftminRaumL24h'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        self.ebene = self.element.Level.Name

        # Anlagenliste erstellen
        if self.Zu_Anl_Nr:
            self.Zu_Anl_Nr_Liste.append(self.Zu_Anl_Nr)

        if self.Ab_Anl_Nr:
            self.Ab_Anl_Nr_Liste.append(self.Ab_Anl_Nr)

        if self.Ab_Anl_Nr_24h:
            self.Ab_24h_Anl_Nr_Liste.append(self.Ab_Anl_Nr_24h)

        # Prüfung. Wenn es Luftmengen gibt aber kein Anlagennummer, dann falsch.

        if self.ZU > 0 and self.Zu_Anl_Nr is None:
            logger.error("Kein Zuluft_Anlage_Nummer in Raum{} gefunden".format(self.nummer))
            script.exit()

        if self.AB > 0 and self.Ab_Anl_Nr is None:
            logger.error("Kein Abluft_Anlage_Nummer in Raum{} gefunden".format(self.nummer))
            script.exit()

        if self.AB24H > 0 and self.Ab_Anl_Nr_24h is None:
            logger.error("Kein Abluft24h_Anlage_Nummer in Raum{} gefunden".format(self.nummer))
            script.exit()


    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter ({}) konnte nicht gefunden werden".format(param_name))
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
            self.ebene,
            self.Zu_Anl_Nr,
            self.Ab_Anl_Nr,
            self.Ab_Anl_Nr_24h,
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)

class DuctSystem:
    anllll = []
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
        self.anllll.append(self.Anl_Nr)


    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter ({}) konnte nicht gefunden werden".format(param_name))
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
            self.Anl_Nr,
            self.Ger_Nr,
            self.name,
            self.Ger_Anzahl,
            self.Anl_Pro_Anz,
        ]

    def __repr__(self):
        return "DuctSystem({})".format(self.element_id)

    def __str__(self):
        return '{}'.format(self.name)


table_data_MepRaum = []
mepraum_liste = []

with forms.ProgressBar(title='{value}/{max_value} Anlagenliste',
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data_MepRaum.append(mepraum.table_row())

# Sorteiren nach Raumnummer
table_data_MepRaum.sort()

table_data_System = []
Systemen_liste = []

with forms.ProgressBar(title='{value}/{max_value} Systemliste',
                       cancellable=True, step=10) as pb1:

    for n, System_id in enumerate(systemen):
        if pb1.cancelled:
            script.exit()

        pb1.update_progress(n + 1, len(systemen))
        ductsystem = DuctSystem(System_id)

        Systemen_liste.append(ductsystem)
        table_data_System.append(ductsystem.table_row())

# Sorteiren nach Anlagennummer und dann Gerätenummer
table_data_System.sort(key = operator.itemgetter(0, 1))

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

# Luftsumme_Liste erstellen
Zuluft_Liste = []
zu = set(MEPRaum.Zu_Anl_Nr_Liste)
Zu = list(zu)
for i in range(len(Zu)):
    Zuluft_Liste.append(0)
logger.info("{}".format(Zu))

Abluft_Liste = []
ab = set(MEPRaum.Ab_Anl_Nr_Liste)
Ab = list(ab)
for j in range(len(Ab)):
    Abluft_Liste.append(0)
logger.info("{}".format(Ab))

Abluft_24h_Liste = []
ab_24h = set(MEPRaum.Ab_24h_Anl_Nr_Liste)
Ab_24h = list(ab_24h)
for k in range(len(Ab_24h)):
    Abluft_24h_Liste.append(0)
logger.info("{}".format(Ab_24h))

# Luftmenge berechnen
with forms.ProgressBar(title='{value}/{max_value} Luftmenge der Anlage berechnen',
                       cancellable=True, step=10) as pb2:

    n_1 = 0
    for space in spaces_collector:

        if pb2.cancelled:
            script.exit()
        n_1 += 1
        pb2.update_progress(n_1, len(spaces))

        Zul_Nr = get_value(space.LookupParameter('TGA_RLT_AnlagenNrZuluft'))
        Zul = get_value(space.LookupParameter('IGF_RLT_ZuluftminRaum'))
        Abl_Nr = get_value(space.LookupParameter('TGA_RLT_AnlagenNrAbluft'))
        Abl = get_value(space.LookupParameter('IGF_RLT_AbluftminRaum'))
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





# Luftmenge Liste für jedes Gerät erstellen
for Ger_Luft in table_data_System:
    Ger_Luft.insert(4,0)
    Ger_Luft.insert(4,0)
    Ger_Luft.insert(4,0)
    Ger_Luft.append(0)
    Ger_Luft.append(0)
    Ger_Luft.append(0)


# Berechnen
with forms.ProgressBar(title='{value}/{max_value} Luftmenge der Geräte berechnen',
                       cancellable=True, step=10) as pb3:

    n_2 = 0
    for Anl_Ger in table_data_System:

        if pb3.cancelled:
            script.exit()

        n_2 += 1
        pb3.update_progress(n_2, len(table_data_System))

        SystemName = Anl_Ger[2]
        AnlagenNr = Anl_Ger[0]
        GeräteAnzahl = Anl_Ger[3]
        AnlagenProAnzahl = Anl_Ger[7]

        if GeräteAnzahl == 0:
            GeräteAnzahl = 1
        if AnlagenProAnzahl == 0:
            AnlagenProAnzahl = 1

        for iii in range(len(Zu)):
            if AnlagenNr == int(Zu[iii]):
                Anl_Ger[4] = round(Zuluft_Liste[iii] / GeräteAnzahl, 1)
                Anl_Ger[8] = round(Zuluft_Liste[iii] / AnlagenProAnzahl, 1)

        for jjj in range(len(Ab)):
            if AnlagenNr == int(Ab[jjj]):
                Anl_Ger[5] = round(Abluft_Liste[jjj] / GeräteAnzahl, 1)
                Anl_Ger[9] = round(Abluft_Liste[jjj] / AnlagenProAnzahl, 1)

        for kkk in range(len(Ab_24h)):
            if AnlagenNr == int(Ab_24h[kkk]):
                Anl_Ger[6] = round(Abluft_24h_Liste[kkk] / GeräteAnzahl, 1)
                Anl_Ger[10] = round(Abluft_24h_Liste[kkk] / AnlagenProAnzahl, 1)




output.print_table(
    table_data=table_data_MepRaum,
    title="MepRaum",
    columns=[ 'Name', 'Nummer','Ebene', 'Zu_Anl_Nr', 'Ab_Anl_Nr', 'Ab_Anl_Nr_24h' ]
)

output.print_table(
    table_data=table_data_System,
    title="System",
    columns=[ 'Anl. Nr','Ger. Nr', 'Systemname', 'Ger. Anzahl', 'Zuluft', 'Abluft',
            '24h-Abluft', 'Anl_Pro_Anz', 'Zuluft', 'Abluft', '24h-Abluft' ]
)

table_data_system = []
for Date in table_data_System:
    if any([Date[0],Date[4],Date[5],Date[6],Date[8],Date[9],Date[10]]):
        table_data_system.append(Date)

output.print_table(
    table_data=table_data_system,
    title="System",
    columns=[ 'Anl. Nr','Ger. Nr', 'Systemname', 'Ger. Anzahl', 'Zuluft', 'Abluft',
            '24h-Abluft', 'Anl_Pro_Anz', 'Zuluft', 'Abluft', '24h-Abluft' ]
)


# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):

    t = Transaction(doc, "Werteschreiben")
    t.Start()

    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb4:

        n_3 = 0
        for Systemen in System_collector:

            if pb4.cancelled:
                script.exit()
            n_3 += 1

            pb4.update_progress(n_3, len(systemen))

            AnlNr = get_value(Systemen.LookupParameter("TGA_RLT_AnlagenNr"))
            GerNr = get_value(Systemen.LookupParameter("TGA_RLT_AnlagenGeräteNr"))
            Zumenge = Systemen.LookupParameter("TGA_RLT_AnlagenZuMenge")
            Abmenge = Systemen.LookupParameter("TGA_RLT_AnlagenAbMenge")
            Ab24hmenge = Systemen.LookupParameter("TGA_RLT_Anlagen24hAbMenge")
            Zumenge_Pro = Systemen.LookupParameter("TGA_RLT_AnlagenProzentualZuMenge")
            Abmenge_Pro = Systemen.LookupParameter("TGA_RLT_AnlagenProzentualAbMenge")
            Ab24hmenge_Pro = Systemen.LookupParameter("TGA_RLT_AnlagenProzentual24hAbMenge")

            for date in table_data_System:
                if date[0] == AnlNr and date[1] == GerNr:
                    Zumenge.SetValueString(str(date[4]))
                    Abmenge.SetValueString(str(date[5]))
                    Ab24hmenge.SetValueString(str(date[6]))
                    Zumenge_Pro.SetValueString(str(date[8]))
                    Abmenge_Pro.SetValueString(str(date[9]))
                    Ab24hmenge_Pro.SetValueString(str(date[10]))
                else:
                    pass





    t.Commit()

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Systemen in Systemliste".format(len(Systemen_liste)))

anll = set(DuctSystem.anllll)
Anll = list(anll)
Anll.sort()
logger.info("{}".format(Anll))
total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
