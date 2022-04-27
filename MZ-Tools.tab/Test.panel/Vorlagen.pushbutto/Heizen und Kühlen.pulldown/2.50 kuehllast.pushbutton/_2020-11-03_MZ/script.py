# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import operator

start = time.time()


__title__ = "Kühllast"
__doc__ = """Kühllastberechnung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc, active_view.Id) \
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
            ['Vol_zu', 'IGF_RLT_ZuluftminRaum'],
            ['T_zu', 'LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE'],
            ['T_raum', 'LIN_BA_DESIGN_COOLING_TEMPERATURE'],
            ['Kuehllast_Labor_Raum', 'IGF_K_KühllastLaborRaum'],
            ['Kuehllast_Labor_PWK', 'IGF_S_KühllastLaborPWK'],
            ['Kuehllast_Gebaeude', 'LIN_BA_CALCULATED_COOLING_LOAD'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        # ZuluftKühlleistung und Kühlleistung vorab Labor berechnen
        self.P_zu = self.zuluft_kuelleistung_berechnen()
        self.Kuehllast_gesamt = self.KuehllastGesamt_Berechnen()
        self.P_ULK = self.KuelleistungULK_Berechnen()
        self.ebene = self.element.Level.Name

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



    def zuluft_kuelleistung_berechnen(self):
        Kuelhlleistung = 0

        if self.Vol_zu and self.T_zu > -273.15 and self.T_raum > -273.15:
            Kuelhlleistung = self.Vol_zu * 1000 * 1.2 * 1.006 * (self.T_zu - self.T_raum) / 3600

        logger.info("Kühlleistung = {}".format(Kuelhlleistung))

        return round(Kuelhlleistung, 2)


    def KuehllastGesamt_Berechnen(self):
        KuehllastGesamt = 0

        KuehllastGesamt = self.Kuehllast_Gebaeude + self.Kuehllast_Labor_Raum


        return KuehllastGesamt

    def KuelleistungULK_Berechnen(self):
        KuelleistungULK = 0

        KuelleistungULK = self.P_zu + self.Kuehllast_gesamt


        return KuelleistungULK


    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                logger.info(
                    "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                self.element.LookupParameter(
                    param_name).SetValueString(str(wert))

        wert_schreiben("IGF_RLT_ZuluftKühlleistung", self.P_zu)
        wert_schreiben("IGF_K_KühllastGesamt", self.Kuehllast_gesamt)
        wert_schreiben("IGF_K_KühlleistungULK", self.P_ULK)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.ebene,
            self.Vol_zu,
            self.T_zu,
            self.T_raum,
            self.Kuehllast_Labor_Raum,
            self.Kuehllast_Labor_PWK,
            self.Kuehllast_Gebaeude,
            self.Kuehllast_gesamt,
            self.P_zu,
            self.P_ULK,

        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


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

# Sortieren nach Raumnummer
table_data.sort()

output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=['Nummer', 'Name', 'Ebene', 'Vol_zu', 'T_zu', 'T_raum',
             'Kühllast_Labor_Raum', 'Kühllast_Labor_PWK', 'Kühllast_Gebaeude',
             'Kühllast_Gesamt','Kühlleistung_Zuluft','Kühlleistung_ULK' ,
             ]
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb2:

        n_1 = 0

        with rpw.db.Transaction("Kühllast"):
            for mepraum in mepraum_liste:
                if pb2.cancelled:
                    script.exit()
                n_1 += 1
                pb2.update_progress(n_1, len(spaces))

                mepraum.werte_schreiben()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
