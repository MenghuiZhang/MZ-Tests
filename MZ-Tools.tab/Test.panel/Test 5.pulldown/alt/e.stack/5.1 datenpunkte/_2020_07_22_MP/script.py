# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time

start = time.time()


__title__ = "7.Datenpunkte"
__doc__ = """Kenngut"""
__author__ = "Maximilian Prachtel"

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
            ['KG_GW', 'IGF_GA_Kenngut_Gerwerk'],
            ['KG_BN', 'IGF_GA_Kenngut_Baunummer'],
            ['KG_EB', 'IGF_GA_Kenngut_Ebene'],
            ['KG_A_1', 'IGF_GA_Kenngut_Anlage_1'],
            ['KG_A_2', 'IGF_GA_Kenngut_Anlage_2'],
            ['KG_DP', 'IGF_GA_Kenngut_Datenpunkt'],
            ['KG_DP_Z', 'IGF_GA_Kenngut_DP-Zähler'],
            ['KG_RN', 'IGF_GA_Kenngut_Raumnummer'],
            ['KG_BT', 'IGF_GA_Kenngut_Bauteil'],

        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        # Beschriftungskürzel für Schema

        self.KZ_S = self.KZ_S_erstellen()

        # Beschriftungskürzel für Gundriss

        self.KZ_G = self.KZ_G_erstellen()

        # Zusammengesetzes Kenngut

        self.KG = self.KG_erstellen()



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
                unit)s

        except Exception as e:
            pass

        return value

    def KZ_S_erstellen(self):
        Kz_S = ''

        Kz_S = str(self.KG_GW) + '-' + str(self.KG_EB) + '-' + str(self.KG_RN) +\
               '-' + str(self.KG_A_1) + str(self.KG_A_2)

        return Kz_S

    def KZ_G_erstellen(self):

        Kz_G = ''

        Kz_G = str(self.KG_GW) + '-' + str(self.KG_A_1) + str(self.KG_A_2)

        return Kz_G

    def KG_erstellen(self):
        Kg = ''

        Kg = str(self.KG_GW) + '-' + str(self.KG_BN) + '-' + str(self.KG_EB) +\
             '-' + str(self.KG_A_1) + str(self.KG_A_2) + '-' + str(self.KG_DP) +\
             '-' + str(self.KG_DP_Z) + '-' + str(self.KG_RN) + '-' + str(self.KG_BT)

        return Kg


    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.name,
            self.nummer,
            self.KG,
            self.KZ_G,
            self.KG_S,
            self.KG_GW,
            self.KG_BN,
            self.KG_EB,
            self.KG_A_1,
            self.KG_A_2,
            self.KG_DP,
            self.KG_DP_Z,
            self.KG_RN,
            self.KG_BT,
        ]


    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                logger.info(
                    "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                self.element.LookupParameter(
                    param_name).SetValueString(str(wert))

        wert_schreiben("IGF_GA_Kuerzel_G", self.KG_G)
        wert_schreiben("IGF_GA_Kuerzel_S", self.KG_S)
        wert_schreiben("IGF_GA_Kenngut", self.KG)



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
    columns=['Name', 'Nummer', 'KG', 'KZ_G', 'KG_S', 'Gewerk', 'Baunummer',
             'Ebene','Anlage Teil 1','Anlage Teil 2', 'Datenpunkt',
             'DP-Zähler', 'Raumnummer', 'Bauteil', ]
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
s
# Werte zuückschreiben + Abfrage
if forms.alert('Beschriftungskürzel in Modell schreiben?', ok=False, yes=True, no=True):
    with rpw.db.Transaction("Beschriftungskürzel schreiben"):
        [mepraum.werte_schreiben() for mepraum in mepraum_liste]

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
