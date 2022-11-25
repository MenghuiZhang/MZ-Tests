# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time

start = time.time()


__title__ = "Verteilung"
__doc__ = """EbenenName und EbenenSortieren"""
__author__ = "Tim Bartels"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
#active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()


logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
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
            ['ebene', 'Ebene'],
            ['name', 'Name'],
            ['nummer', 'Nummer'],
        ]



        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        self.eb = self.Ebene_Abkuerzung()
        self.eb_nr = self.Ebene_nr()


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

    def Ebene_Abkuerzung(self):
        eb = 'EG'

        if str(self.ebene) == '17024631':
            eb = '2.UG'
        elif str(self.ebene) == '17024607':
            eb = '1.UG'
        elif str(self.ebene) == '17024593':
            eb = 'EG'
        elif str(self.ebene) == '17024596':
            eb = '1.OG'
        elif str(self.ebene) == '17024599':
            eb = '2.OG'
        elif str(self.ebene) == '17024602':
            eb = '3.OG'
        elif str(self.ebene) == '17024604':
            eb = '4.OG'

        return eb


    def Ebene_nr(self):
        eb_nr = 0

        if str(self.ebene) == '17024631':
            eb_nr = -2
        elif str(self.ebene) == '17024607':
            eb_nr = -1
        elif str(self.ebene) == '17024593':
            eb_nr = 0
        elif str(self.ebene) == '17024596':
            eb_nr = 1
        elif str(self.ebene) == '17024599':
            eb_nr = 2
        elif str(self.ebene) == '17024602':
            eb_nr = 3
        elif str(self.ebene) == '17024604':
            eb_nr = 4

        return eb_nr



    def werte_schreiben(self):
        """Schreibt die Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                logger.info(
                    "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                self.element.LookupParameter(
                    param_name).Set(wert)

        wert_schreiben("IGF_RLT_Verteilung_EbenenName", self.eb)
        wert_schreiben("IGF_RLT_Verteilung_EbenenSortieren", self.eb_nr)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.ebene,
            self.name,
            self.nummer,
            self.eb,
            self.eb_nr,
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
    title="MEP Räume",
    columns=['Ebene','Name', 'Nummer','EB','EB_Nr', ]
)


logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with rpw.db.Transaction("Luftwechsel berechnen"):
        [mepraum.werte_schreiben() for mepraum in mepraum_liste]

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
