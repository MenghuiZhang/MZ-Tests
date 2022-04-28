# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB, db
from pyrevit import script, forms
import clr

__title__ = "5.00 Betreibbetriebskennzeichung"
__doc__ = """BKZ beschriften"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc




# Luftkanalzubehör aus aktueller Ansicht
ducts_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType()

HLS_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()

L_Sys_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Plumbing.PipingSystem))

R_Sys_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystem))

ElementIdListe = []
for el in ducts_collector:
    ElementIdListe.append(el.Id)

for el in HLS_collector:
    ElementIdListe.append(el.Id)

Sys = {}
for el in L_Sys_collector:
    try:
        Nr = el.LookupParameter('IGF_GA_BKZ_System_Nummer').AsString()
        Typ = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsString()
        Name = el.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
        if Typ in Sys.keys():
            Sys[Typ][Name] = Nr
        else:
            Sys[Typ] = {Name:Nr}

    except:
        pass


for el in R_Sys_collector:
    try:
        Nr = el.LookupParameter('IGF_GA_BKZ_System_Nummer').AsString()
        Typ = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsString()
        Name = el.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
        if Typ in Sys.keys():
            Sys[Typ][Name] = Nr
        else:
            Sys[Typ] = {Name:Nr}
    except:
        pass

phase = doc.Phases[0]

class Element:
    def __init__(self, element_id):
        """
        Definiert Element Klasse mit allen object properties für die Kenngut Beschriftung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Systemname'],
            ['KG_GW', 'IGF_GA_BKZ_Gewerk'],
            ['KG_BN', 'IGF_GA_BKZ_Baunummer'],
            ['KG_EB', 'IGF_GA_BKZ_Ebene'],
            ['KG_Art', 'IGF_GA_BKZ_Anlage_Art'],
            ['KG_AN', 'IGF_GA_BKZ_Anlage_Nummer'],
            ['KG_DP', 'IGF_GA_BKZ_Datenpunkt'],
            ['KG_DPZ', 'IGF_GA_BKZ_DP-Zähler'],
            ['KG_RN', 'IGF_GA_BKZ_Raumnummer'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(40 * "=")
        logger.info("{}\t{}".format(self.Systemtyp, self.name))

        self.Systemtyp = self.element.LookupParameter('Systemtyp').AsValueString()

        try:
            self.KG_Art = self.element.Symbol.LookupParameter('IGF_GA_BKZ_Anlage_Familie_Art').AsString()
        except:
            pass

        try:
            self.KG_AN = Sys[self.Systemtyp][self.name]
        except:
            pass

        if self.element.Space[phase]:
            if self.element.Space[phase].LookupParameter('IGF_GA_BKZ_Zone'):
                nr = self.element.Space[phase].LookupParameter('IGF_GA_BKZ_Zone').AsString()
                if nr:
                    self.KG_Art = nr[0]
                    self.KG_AN = nr[1:]


        # Beschriftungskürzel für Gundriss erstellen

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

        return get_value(param)



    def KZG_erstellen(self):
        KZG = ''

        KZG = str(self.KG_GW) + '-' + str(self.KG_Art) + str(self.KG_AN)

        logger.info("IGF_GA_Kuerzel_G = {}".format(KZG))

        return KZG

    def KZS_erstellen(self):
        KZS = ''

        KZS = str(self.KG_GW) + '-' + str(self.KG_EB) + '.' + str(self.KG_RN) + str(self.KG_Art) + str(self.KG_AN)

        logger.info("IGF_GA_Kuerzel_S = {}".format(KZS))

        return KZS

    def KG_erstellen(self):
        KG = ''

        KG = str(self.KG_GW) + str(self.KG_BN) + str(self.KG_EB) +  str(self.KG_Art) + str(self.KG_AN) + str(self.KG_RN) + str(self.KG_DP) + str(self.KG_DPZ)

        logger.info("IGF_GA_BKZ = {}".format(KG))

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
        wert_schreiben("IGF_GA_BKZ", self.KG)
        wert_schreiben("IGF_GA_BKZ_Anlage_Art", self.KG_Art)
        wert_schreiben("IGF_GA_BKZ_Anlage_Nummer", self.KG_AN)

    def table_row(self):
        """ Gibt eine Datenreihe für den Luftkanalzubehörs aus. Für die tabellarische Übersicht."""
        return [
            self.Systemtyp,
            self.name,
            self.KG,
            self.KZ_G,
            self.KZ_S,
            self.KG_GW,
            self.KG_BN,
            self.KG_EB,
            self.KG_Art,
            self.KG_AN,
            self.KG_DP,
            self.KG_DPZ,
            self.KG_RN,
        ]

table_data = []
Element_Liste = []
with forms.ProgressBar(title='{value}/{max_value} Betreibbetriebskennzeichung',
                       cancellable=True, step=10) as pb:

    for n, id in enumerate(ElementIdListe):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(ElementIdListe))
        elem = Element(id)

        Element_Liste.append(elem)
        table_data.append(elem.table_row())

table_data.sort()

output.print_table(
    table_data=table_data,
    title="Luftkanalzubehör",
    columns=['Systemtyp', 'Name',  'BKZ', 'Kürzel_G', 'Kürzel_S', 'Gewerk', 'Baunummer', 'Ebene',
             'Anlage_Art','Anlage_Nr.', 'Datenpunkt', 'DP-Zähler','Raumnummer']
)


# Beschriftung in Modell schreiben + Abfrage
if forms.alert('BKZ in Modell schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} BKZ schreiben',
                           cancellable=True, step=10) as pb2:

        with db.Transaction("BKZ schreiben"):
            for n_1, eleme in enumerate(Element_Liste):
                if pb2.cancelled:
                    script.exit()
                pb2.update_progress(n_1, len(Element_Liste))
                eleme.werte_schreiben()

