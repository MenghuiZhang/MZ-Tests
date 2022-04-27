# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import operator

start = time.time()


__title__ = "6.Kenngut"
__doc__ = """Kenngut beschriften"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

# Luftkanalzubehör aus aktueller Ansicht
ducts_collector = DB.FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(DB.BuiltInCategory.OST_DuctAccessory)
ducts = ducts_collector.ToElementIds()

logger.info("{} Luftkanalzubehör ausgewählt".format(len(ducts)))

if not ducts:
    logger.error("Keine Luftkanalzubehör in aktueller Ansicht gefunden")
    script.exit()


class LUFTKanalzubehoer:
    def __init__(self, element_id):
        """
        Definiert Luftkanalzubehör Klasse mit allen object properties für die Kenngut Beschriftung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Systemname'],
            ['Systemtyp', 'Systemtyp'],
            ['KZ', 'Kennzeichen'],
            ['KG_GW', 'IGF_GA_Kenngut_Gerwerk'],
            ['KG_BN', 'IGF_GA_Kenngut_Baunummer'],
            ['KG_EB', 'IGF_GA_Kenngut_Ebene'],
            ['KG_A1', 'IGF_GA_Kenngut_Anlage_1'],
            ['KG_A2', 'IGF_GA_Kenngut_Anlage_2'],
            ['KG_DP', 'IGF_GA_Kenngut_Datenpunkt'],
            ['KG_DPZ', 'IGF_GA_Kenngut_DP-Zähler'],
            ['KG_RN', 'IGF_GA_Kenngut_Raumnummer'],
            ['KG_BT', 'IGF_GA_Kenngut_Bauteil'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(40 * "=")
        logger.info("{}\t{}".format(self.Systemtyp, self.name))

        # Beschriftungskürzel für Gundriss  erstellen

        self.KZ_G = self.KZG_erstellen()

        # Beschriftungskürzel für Schema erstellen

        self.KZ_S = self.KZS_erstellen()

        # Zusammengesetzes Kenngut erstellen

        self.KG = self.KG_erstellen()

    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter ( {} ) konnte nicht gefunden werden".format(param_name))
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


    def KZG_erstellen(self):
        KZG = ''

        KZG = str(self.KG_GW) + '-' + str(self.KG_A1) + str(self.KG_A2)

        logger.info("IGF_GA_Kuerzel_G = {}".format(KZG))

        return KZG

    def KZS_erstellen(self):
        KZS = ''

        KZS = str(self.KG_GW) + '-' + str(self.KG_EB) + '.' + str(self.KG_RN) + str(self.KG_A1) + str(self.KG_A2)

        logger.info("IGF_GA_Kuerzel_S = {}".format(KZS))

        return KZS

    def KG_erstellen(self):
        KG = ''

        KG = str(self.KG_GW) + str(self.KG_BN) + str(self.KG_EB) +  str(self.KG_A1) + str(self.KG_A2) + str(self.KG_DP) + str(self.KG_DPZ)

        logger.info("IGF_GA_Kenngut = {}".format(KG))

        return KG

    def werte_schreiben(self):
        """Schreibt die Beschriftung zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                logger.info(
                    "{} - {} schreiben ({})".format(self.Systemtyp, param_name, wert))
                self.element.LookupParameter(
                    param_name).Set(wert)

        wert_schreiben("IGF_GA_Kuerzel_G", self.KZ_G)
        wert_schreiben("IGF_GA_Kuerzel_S", self.KZ_S)
        wert_schreiben("IGF_GA_Kenngut", self.KG)

    def table_row(self):
        """ Gibt eine Datenreihe für den Luftkanalzubehörs aus. Für die tabellarische Übersicht."""
        return [
            self.Systemtyp,
            self.name,
            self.KZ,
            self.KZ_G,
            self.KZ_S,
            self.KG,
            self.KG_GW,
            self.KG_BN,
            self.KG_EB,
            self.KG_A1,
            self.KG_A2,
            self.KG_DP,
            self.KG_DPZ,
            self.KG_RN,
            self.KG_BT,
        ]

    def __repr__(self):
        return "Luftkanalzubehör({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.Systemtyp, self.name)


table_data = []
Luftkanalzubehoer_Liste = []
with forms.ProgressBar(title='{value}/{max_value} Kenngut',
                       cancellable=True, step=10) as pb:

    for n, duct_id in enumerate(ducts):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(ducts))
        Luftkanalzubehoer = LUFTKanalzubehoer(duct_id)

        Luftkanalzubehoer_Liste.append(Luftkanalzubehoer)
        table_data.append(Luftkanalzubehoer.table_row())

# Sortieren nach Kennzeichen
table_data.sort(key = operator.itemgetter(2))

output.print_table(
    table_data=table_data,
    title="Luftkanalzubehör",
    columns=['Systemtyp', 'Name', 'Kennzeichen', 'KZ_G', 'KZ_S', 'KG', 'KG_GW', 'KG_BN', 'KG_EB',
             'KG_A1','KG_A2', 'KG_DP', 'KG_DPZ','KG_RN', 'KG_BT',]

)

logger.info("{} Luftkanalzubehör in Luftkanalzubehör_Liste".format(len(Luftkanalzubehoer_Liste)))

# Beschriftung in Modell schreiben + Abfrage
if forms.alert('Kenngut in Modell schreiben?', ok=False, yes=True, no=True):
    with rpw.db.Transaction("Kenngut Beschriften"):
        [Luftkanalzubehoer.werte_schreiben() for Luftkanalzubehoer in Luftkanalzubehoer_Liste]

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
