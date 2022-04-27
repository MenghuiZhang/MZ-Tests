# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time

start = time.time()


__title__ = "5.Geschoss"
__doc__ = """Kühllastberechnung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc


# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()


class MEPRaum:
    Ebene_Liste = []
    Schacht_Liste = []
    Zu_Anl_Nr_Liste = []
    Ab_Anl_Nr_Liste = []
    Ab_24h_Anl_Nr_Liste = []
    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties für die
        Kühlleistung Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['ebene', 'Ebene'],
            ['Zu_Anl_Nr', 'TGA_RLT_AnlagenNrZuluft'],
            ['Ab_Anl_Nr', 'TGA_RLT_AnlagenNrAbluft'],
            ['Ab_Anl_Nr_24h', 'TGA_RLT_AnlagenNr24hAbluft'],
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],
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


        # Schachtliste erstellen

        if self.Install_Schacht:
            self.Schacht_Liste.append(self.Install_Schacht_Name)

        # Ebeneliste erstellen
        if self.ebene:
            self.Ebene_Liste.append(self.ebene)


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
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)





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




table_data = []
mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} Kühlleistung berechnen',
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())


# Luft_Liste erstellen

Zu = set(MEPRaum.Zu_Anl_Nr_Liste)
ZuAnlNrListe = list(Zu)

Ab = set(MEPRaum.Ab_Anl_Nr_Liste)
AbAnlNrListe = list(Ab)

Ab_24h = set(MEPRaum.Ab_24h_Anl_Nr_Liste)
Ab24hAnlNrListe = list(Ab_24h)

Eb = set(MEPRaum.Ebene_Liste)
EbeneListe = list(Eb)

Sc = set(MEPRaum.Schacht_Liste)
SchachtListe = list(Sc)

for Spa in spaces_collector:
    eb = Spa.LookupParameter('Ebene').AsString()
    logger.info('Ebene: {}'.format(eb))


# Luftverteilung_List

Schacht_gesamt = []
Schacht_Zu = []
Schacht_Ab = []
Schacht_Ab24h = []
Ebene_zu = []
Ebene_ab = []
Ebene_ab24h = []
Anl_zu = []
Anl_ab = []
Anl_ab24h = []

for a in range(len(ZuAnlNrListe)):
    Anl_zu.append(0)
for b in range(len(AbAnlNrListe)):
    Anl_ab.append(0)
for c in range(len(Ab24hAnlNrListe)):
    Anl_ab24h.append(0)

for d in range(len(EbeneListe)):
    Ebene_zu.append(Anl_zu)
    Ebene_ab.append(Anl_ab)
    Ebene_ab24h.append(Anl_ab24h)

for e in range(len(SchachtListe)):
    Schacht_Zu.append(Ebene_zu)
    Schacht_Ab.append(Ebene_ab)
    Schacht_Ab24h.append(Ebene_ab24h)

for f in range(len(SchachtListe)):
    gesamt = []
    gesamt.append(Schacht_Zu[f])
    gesamt.append(Schacht_Ab[f])
    gesamt.append(Schacht_Ab24h[f])
    Schacht_gesamt.append(gesamt)



# berechnen
#for Space in spaces_collector:
#    Zu_Schacht = get_value(space.LookupParameter('TGA_RLT_SchachtZuluft'))
#    Ab_Schacht =
#    Ab24h_Schacht =
#    Zu_Anl =
#    Ab_Anl =
#    Ab24h_Anl =
#    Ebenen =





output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=['Nummer', 'Name', ]
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Schächte in Schächteliste".format(len(SchachtListe)))
logger.info("Schächte: {}".format(SchachtListe))
logger.info("{} Ebene in Ebeneliste".format(len(EbeneListe)))
logger.info("Ebene: {}".format(EbeneListe))
logger.info("{} Zuluftanlagen in Zuluftanlagenliste".format(len(ZuAnlNrListe)))
logger.info("Zuluftanlagen: {}".format(ZuAnlNrListe))
logger.info("{} Abluftanlagen in Abluftanlagenliste".format(len(AbAnlNrListe)))
logger.info("Abluftanlagen: {}".format(AbAnlNrListe))
logger.info("{} Abluft_24h_Anlagen in Abluft_24h_Anlagenliste".format(len(Ab24hAnlNrListe)))
logger.info("Abluft_24h_Anlagen: {}".format(Ab24hAnlNrListe))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
