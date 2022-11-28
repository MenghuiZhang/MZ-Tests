# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr


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
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        # Anlagenliste erstellen
        if self.Zu_Anl_Nr:
            self.Zu_Anl_Nr_Liste.append(self.Zu_Anl_Nr)

        if self.Ab_Anl_Nr:
            self.Ab_Anl_Nr_Liste.append(self.Ab_Anl_Nr)

        if self.Ab_Anl_Nr_24h:
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



# Luftmenge berechnen
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

# Luftmenge Liste für jedes Gerät erstellen
zumenge = []
abmenge = []
ab24hmenge = []
zumenge_pro = []
abmenge_pro = []
ab24hmenge_pro = []
for s in range(len(systemen)):
    zumenge.append(0)
    abmenge.append(0)
    ab24hmenge.append(0)
    zumenge_pro.append(0)
    abmenge_pro.append(0)
    ab24hmenge_pro.append(0)


def Table_row():
    """ Gibt eine Datenreihe für Luftkanalsystem aus. Für die tabellarische Übersicht."""
    return [
        AnlagenNr,
        zumenge[Nr],
        abmenge[Nr],
        ab24hmenge[Nr],
        zumenge_pro[Nr],
        abmenge_pro[Nr],
        ab24hmenge_pro[Nr]
    ]

# Geräteanzahl prüfen und fenlende Informationen ergänzen

Fehlende_Info = []

for Sys in System_collector:
    Sys_Name = get_value(Sys.LookupParameter('Systemname'))
    Ger_Anzahl = get_value(Sys.LookupParameter('TGA_RLT_AnlagenGeräteAnzahl'))
    Anl_Pro_Anzahl = get_value(Sys.LookupParameter('TGA_RLT_AnlagenProzentualAnzahl'))
    Anl_Nummer = get_value(Sys.LookupParameter('TGA_RLT_AnlagenNr'))

    if Anl_Nummer:
        if Ger_Anzahl == 0 or Anl_Pro_Anzahl == 0:
            Fehlende_Info.append(Sys_Name)

if len(Fehlende_Info) > 0:
    logger.info('keine Informationen zur Geräteanzahl/AnlagenProzentualAnzahl/Gerätenummer in den folgenden Systemen')
    logger.info('{}'.format(Fehlende_Info))
    if forms.alert('Felenden Informationen in Modell schreiben?', ok=False, yes=True, no=True):
        t0 = Transaction(doc, "Werteschreiben")
        t0.Start()
        for sys in System_collector:
            for a in range(len(Fehlende_Info)):
                if get_value(sys.LookupParameter('Systemname')) == Fehlende_Info[a]:
                    logger.info('{}'.format(Fehlende_Info[a]))

                    Ger_Anz = sys.LookupParameter("TGA_RLT_AnlagenGeräteAnzahl")
                    Ger_Anz.SetValueString(str(1))

                    Anl_Pro_Anz = sys.LookupParameter("TGA_RLT_AnlagenProzentualAnzahl")
                    Anl_Pro_Anz.SetValueString(str(1))

                    Ger_nr = sys.LookupParameter("TGA_RLT_AnlagenGeräteNr")
                    Ger_nr.SetValueString(str(1))

        t0.Commit()




# Berechnen

Nr = 0
AnlagenNr = 0
table_data_Luftkanalsystemen = []

for System in System_collector:
    SystemName = get_value(System.LookupParameter('Systemname'))
    AnlagenNr = get_value(System.LookupParameter('TGA_RLT_AnlagenNr'))
    GeräteAnzahl = get_value(System.LookupParameter('TGA_RLT_AnlagenGeräteAnzahl'))
    AnlagenProAnzahl = get_value(System.LookupParameter('TGA_RLT_AnlagenProzentualAnzahl'))

    for iii in range(len(Zu)):
        if AnlagenNr == int(Zu[iii]):
            zumenge[Nr] = round(Zuluft_Liste[iii] / GeräteAnzahl, 1)
            zumenge_pro[Nr] = round(Zuluft_Liste[iii] / AnlagenProAnzahl, 1)

    for jjj in range(len(Ab)):
        if AnlagenNr == int(Ab[jjj]):
            abmenge[Nr] = round(Abluft_Liste[jjj] / GeräteAnzahl, 1)
            abmenge_pro[Nr] = round(Abluft_Liste[jjj] / AnlagenProAnzahl, 1)

    for kkk in range(len(Ab_24h)):
        if AnlagenNr == int(Ab_24h[kkk]):
            ab24hmenge[Nr] = round(Abluft_24h_Liste[kkk] / GeräteAnzahl, 1)
            ab24hmenge_pro[Nr] = round(Abluft_24h_Liste[kkk] / AnlagenProAnzahl, 1)

    table_data_Luftkanalsystemen.append(Table_row())

    logger.info("AnlagenNr ist {}".format(AnlagenNr))
    logger.info("GeräteAnzahl ist {}".format(GeräteAnzahl))
    logger.info("AnlagenProAnzahl ist {}".format(AnlagenProAnzahl))
    logger.info("Zuluft{}".format(zumenge[Nr]))
    logger.info("Abluft{}".format(abmenge[Nr]))
    logger.info("Abluft_24h{}".format(ab24hmenge[Nr]))
    logger.info("Zuluft_pro{}".format(zumenge_pro[Nr]))
    logger.info("Abluft_pro{}".format(abmenge_pro[Nr]))
    logger.info("Abluft_24h_pro{}".format(ab24hmenge_pro[Nr]))
    logger.info(60 * "=")

    Nr = Nr + 1


output.print_table(
    table_data=table_data_MepRaum,
    title="MepRaum",
    columns=[ 'Name', 'Nummer','Zu_Anl_Nr', 'Ab_Anl_Nr', 'Ab_Anl_Nr_24h' ]
)

output.print_table(
    table_data=table_data_System,
    title="System",
    columns=[ 'Name','Ger_Nr', 'Ger_Anzahl', 'Anl_Nr', 'Anl_Pro_Anz' ]
)

output.print_table(
    table_data=table_data_Luftkanalsystemen,
    title="Luftmengen",
    columns=[ 'AnlagenNr','Zu', 'Ab', 'Ab_24h', 'Zu_Pro', 'Ab_Pro', 'Ab_24h_Pro' ]
)



# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    t = Transaction(doc, "Werteschreiben")
    t.Start()
    num = 0
    for Systemen in System_collector:
        Zumenge = Systemen.LookupParameter("TGA_RLT_AnlagenZuMenge")
        Zumenge.SetValueString(str(zumenge[num]))

        Abmenge = Systemen.LookupParameter("TGA_RLT_AnlagenAbMenge")
        Abmenge.SetValueString(str(abmenge[num]))

        Ab24hmenge = Systemen.LookupParameter("TGA_RLT_Anlagen24hAbMenge")
        Ab24hmenge.SetValueString(str(ab24hmenge[num]))

        Zumenge_Pro = Systemen.LookupParameter("TGA_RLT_AnlagenProzentualZuMenge")
        Zumenge_Pro.SetValueString(str(zumenge_pro[num]))

        Abmenge_Pro = Systemen.LookupParameter("TGA_RLT_AnlagenProzentualAbMenge")
        Abmenge_Pro.SetValueString(str(abmenge_pro[num]))

        Ab24hmenge_Pro = Systemen.LookupParameter("TGA_RLT_AnlagenProzentual24hAbMenge")
        Ab24hmenge_Pro.SetValueString(str(ab24hmenge_pro[num]))

        num = num +1
    t.Commit()

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Systemen in Systemliste".format(len(Systemen_liste)))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
