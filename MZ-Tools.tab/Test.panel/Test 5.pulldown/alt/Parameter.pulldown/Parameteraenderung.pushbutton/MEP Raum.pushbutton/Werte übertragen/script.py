# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "0.Übertragen"
__doc__ = """Luftmengenberechnung"""
__author__ = "MZ"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))


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
            ['ZU_min', 'TGA_RLT_ZuluftminRaum'],
            ['ZU_max', 'TGA_RLT_ZuluftmaxRaum'],
            ['ZU_min_ZR', 'TGA_RLT_ZuluftminRaumZR'],
            ['AB_24h', 'TGA_RLT_AbluftSumme24h'],
            ['AB_min', 'TGA_RLT_AbluftminRaum'],
            ['AB_LB', 'TGA_RLT_AbluftSummeLabor'],
            ['AB_max', 'TGA_RLT_AbluftmaxRaum'],
            ['AB_LB_24h', 'TGA_RLT_AbluftSummeLabor24h'],
            ['AB_min_ZR', 'TGA_RLT_AbluftminRaumZR'],
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



    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            self.element.LookupParameter(
                param_name).SetValueString(str(wert))

        wert_schreiben("IGF_RLT_ZuluftminRaum", self.ZU_min)
        wert_schreiben("IGF_RLT_ZuluftmaxRaum", self.ZU_max)
        wert_schreiben("Angegebener Zuluftstrom", self.ZU_min)
    #    wert_schreiben("IGF_RLT_ZuluftvorabRaumZR", self.ZU_min_ZR)
        wert_schreiben("TGA_RLT_AbluftSumme24h", self.AB_24h)
        wert_schreiben("IGF_RLT_AbluftminRaum", self.AB_min)
        wert_schreiben("IGF_RLT_AbluftminSummeLabor",self.AB_LB)
        wert_schreiben("IGF_RLT_AbluftmaxRaum", self.AB_max)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor24h", self.AB_LB_24h)
        wert_schreiben("Angegebener Abluftluftstrom", self.AB_min)
    #    wert_schreiben("IGF_RLT_AbluftvorabRaumZR", self.AB_min_ZR)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.ZU_min,
            self.ZU_max,
            self.ZU_min_ZR,
            self.AB_24h,
            self.AB_min,
            self.AB_LB,
            self.AB_max,
            self.AB_LB_24h,
            self.AB_min_ZR,

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

#  Sortieren nach Raumnummer
table_data.sort()

output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=['Nummer', 'Name', 'ZU_min', 'ZU_max', 'ZU_min_ZR', 'AB_24h','AB_min','AB_LB',
             'AB_max', 'AB_LB_24h', 'AB_min_ZR']
)


logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))

if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb2:



        with rpw.db.Transaction("Luftwechsel berechnen"):
            n_1 = 0
            for mepraum in mepraum_liste:
                if pb2.cancelled:
                    script.exit()

                n_1 += 1

                pb2.update_progress(n_1, len(spaces))

                mepraum.werte_schreiben()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
