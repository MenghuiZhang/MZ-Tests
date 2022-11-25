# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
#import xlsxwriter

start = time.time()


__title__ = "3.1 Raumluft_Raum"
__doc__ = """Luftmengenberechnung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

# MEP Räume mit Überströmung
param_id = revit.query.get_project_parameter_id('TGA_RLT_RaumÜberströmungAus')
if param_id:
    pvprov = DB.ParameterValueProvider(param_id)
    pfilter = DB.FilterStringGreater()
    vrule = DB.FilterStringRule(pvprov, pfilter, "", False)

    param_filter = DB.ElementParameterFilter(vrule)
    spaces_ueberstroemung = spaces_collector \
        .WherePasses(param_filter) \
        .ToElements()


logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
logger.info("{} MEP Räume mit Überströmung".format(len(spaces_ueberstroemung)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

berechnung_nach = {
    'Fläche': 1,
    'Luftwechsel': 2,
    'Person': 3,
    'manuell': 4,
    'nurZUMa': 5,
    'nurABMa': 6,
    'keine': 9
}


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
            ['flaeche', 'Fläche'],
            ['volumen', 'Volumen'],
            ['personen', 'Personenzahl'],
            ['kez', 'TGA_RLT_VolumenstromProNummer'],
            ['vol_faktor', 'TGA_RLT_VolumenstromProFaktor'],
            ['raum_druckstufe', 'TGA_RLT_RaumDruckstufeEingabe'],
            ['ueberstroemung', 'IGF_RLT_ÜberströmungRaum'],
            ['ABL_24h', 'IGF_RLT_AbluftSumme24h'],
            ['ABL_Labor', 'IGF_RLT_AbluftminSummeLabor'],
            # Nachbetrieb
            ['nachtbetrieb', 'IGF_RLT_Nachtbetrieb'],
            ['NB_Von', 'IGF_RLT_NachtbetriebVon'],
            ['NB_Bis', 'IGF_RLT_NachtbetriebBis'],
            ['NB_LW', 'IGF_RLT_NachtbetriebLW'],
            # Tiefnachtbetrieb
            ['tiefernachtbetrieb', 'IGF_RLT_TieferNachtbetrieb'],
            ['T_NB_Von', 'IGF_RLT_TieferNachtbetriebVon'],
            ['T_NB_Bis', 'IGF_RLT_TieferNachtbetriebBis'],
            ['T_NB_LW', 'IGF_RLT_TieferNachtbetriebLW']
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        self.abluft_labor_24h = self.ABL_Labor + self.ABL_24h
        self.flaeche = round(self.flaeche, 0)
        self.ebene = self.element.Level.Name

        self.zuluft_min = self.zuluft_min_berechnen()

        self.abluft_ges = self.Abluft_Ges_Berechnen(self.zuluft_min)
        self.abluft_ohne_24h,self.abluft_ohne_Ueber,self.abluft_min = self.abluft_berechnen()

        self.nb_dauer = self.nb_dauer_berechnen()
        self.zu_nacht = self.zuluft_nacht_berechnen()
        self.ab_nacht,self.ab_nacht_ohne24h = self.abluft_nacht_berechnen()

        self.tiefer_nb_dauer = self.tiefer_nb_dauer_berechnen()
        self.tiefer_zu_nacht = self.tiefer_zuluft_nacht_berechnen()
        self.tiefer_ab_nacht,self.tiefer_ab_nacht_ohne24h = self.tiefer_abluft_nacht_berechnen()


        self.IGF_Druckstufe = self.Druckstufe_Berechnen()

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

    def __lab_abl_24h_druckstufe_pruefen(self, zuluft):
        if self.abluft_labor_24h > zuluft:
            logger.info(
                "{} m/h3 Labor und 24h Abluft berücksichtigt".format(self.abluft_labor_24h))

            zuluft = self.abluft_labor_24h

        if self.raum_druckstufe > 0:
            logger.info(
                "{} - Druckstufe berücksichtigt".format(self.raum_druckstufe))

            zuluft = zuluft + self.raum_druckstufe

        return zuluft

    def Abluft_Ges_Berechnen(self,Zuluft):
        Abluft_Ges = 0
        if int(self.kez) in [
            berechnung_nach['Fläche'],
            berechnung_nach['Luftwechsel'],
            berechnung_nach['Person'],]:

            Abluft_Ges = Zuluft + self.ueberstroemung - self.raum_druckstufe

        elif int(self.kez) == berechnung_nach['nurABMa']:
            Abluft_Ges = self.vol_faktor

        elif int(self.kez) == berechnung_nach['manuell']:
            Abluft_Ges = self.vol_faktor - self.raum_druckstufe

        return round(Abluft_Ges,2)


    def abluft_berechnen(self):
        abluft_ohne_24h = 0
        abluft_ohne_Ueber = 0

        abluft_ohne_24h = self.abluft_ges - self.ABL_24h
        abluft_ohne_Ueber = self.abluft_ges - self.ueberstroemung
        abluft_min = self.abluft_ges - self.ABL_24h - self.ueberstroemung

        return round(abluft_ohne_24h,2),round(abluft_ohne_Ueber,2),round(abluft_min,2)


    def zuluft_min_berechnen(self):
        zuluft = 0

        logger.info("Fläche: {}".format(self.flaeche))

        if int(self.kez) == berechnung_nach['Fläche']:
            logger.info("Berechnung nach Fläche")
            if self.flaeche == 0:
                logger.error("Die Fläche des Raumes {} {} ist 0m2".format(
                    self.nummer, self.name))
                return

            zuluft = self.flaeche * self.vol_faktor

        elif int(self.kez) == berechnung_nach['Luftwechsel']:
            logger.info("Berechnung nach Luftwechsel")

            zuluft = self.volumen * self.vol_faktor

        elif int(self.kez) == berechnung_nach['Person']:
            logger.info("Berechnung nach Personen")

            zuluft = self.personen * self.vol_faktor

        elif int(self.kez) in [
                berechnung_nach['manuell'],
                berechnung_nach['nurZUMa']]:

            zuluft = self.vol_faktor


        zuluft = self.__lab_abl_24h_druckstufe_pruefen(zuluft)

        logger.info("ZUL = {}".format(zuluft))

        return round(zuluft, 2)

    def Druckstufe_Berechnen(self):
        DS = 0
        DS = self.zuluft_min - self.abluft_ges + self.ueberstroemung
        return DS

    def nb_dauer_berechnen(self):
        nb_dauer = 0

        if self.nachtbetrieb:

            if self.tiefernachtbetrieb:
                nb_dauer = self.NB_Bis - self.NB_Von - self.T_NB_Bis + self.T_NB_Von

            else:
                nb_dauer = self.NB_Bis - self.NB_Von + 24.00

        logger.info("Nachtbetrieb_Dauer = {}".format(nb_dauer))
        return nb_dauer

    def abluft_nacht_berechnen(self):
        ab_nacht_ohne24h = 0
        ab_nacht_ges = 0

        if self.nachtbetrieb:
            ab_nacht_ges = self.Abluft_Ges_Berechnen(self.zu_nacht)
            ab_nacht_ohne24h = ab_nacht_ges - self.ABL_24h

        return round(ab_nacht_ges,2),round(ab_nacht_ohne24h,2)

    def zuluft_nacht_berechnen(self):
        zu_nacht = 0

        if self.nachtbetrieb:
            zu_nacht = self.NB_LW * self.volumen

            zu_nacht = self.__lab_abl_24h_druckstufe_pruefen(zu_nacht)

        logger.info("ZUL_Nacht = {}".format(zu_nacht))

        return zu_nacht

    def tiefer_nb_dauer_berechnen(self):
        tiefer_nb_dauer = 0

        if self.tiefernachtbetrieb:

            # Tim: 24:00 kann nicht gerechnet werden. Es müssen Gleitkommazahlen verwendet werden z.B 24.00
            tiefer_nb_dauer = self.T_NB_Bis - self.T_NB_Von + 24.00

        logger.info("TieferNachtbetrieb_Dauer = {}".format(tiefer_nb_dauer))
        return tiefer_nb_dauer

    def tiefer_abluft_nacht_berechnen(self):
        tiefer_ab_nacht = 0
        tiefer_ab_nacht_ohne24h = 0

        if self.tiefernachtbetrieb:
            tiefer_ab_nacht = self.Abluft_Ges_Berechnen(self.tiefer_zu_nacht)
            tiefer_ab_nacht_ohne24h = tiefer_ab_nacht - self.ABL_24h

        return round(tiefer_ab_nacht,2),round(tiefer_ab_nacht_ohne24h,2)

    def tiefer_zuluft_nacht_berechnen(self):
        tiefer_zu_nacht = 0

        if self.nachtbetrieb:
            tiefer_zu_nacht = self.T_NB_LW * self.volumen

            tiefer_zu_nacht = self.__lab_abl_24h_druckstufe_pruefen(tiefer_zu_nacht)

        logger.info("Tief_ZUL_Nacht = {}".format(tiefer_zu_nacht))

        return tiefer_zu_nacht


    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                logger.info(
                    "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                self.element.LookupParameter(
                    param_name).SetValueString(str(wert))

        wert_schreiben("IGF_RLT_AbluftminRaumGes", self.abluft_ges)
        wert_schreiben("IGF_RLT_AbluftminRaumOhne24h", self.abluft_ohne_24h)
        wert_schreiben("IGF_RLT_AbluftminRaumOhneÜber", self.abluft_ohne_Ueber)
        wert_schreiben("IGF_RLT_AbluftminRaum", self.abluft_min)

        wert_schreiben("Angegebener Zuluftstrom", self.zuluft_min)
        wert_schreiben("Angegebener Abluftluftstrom", self.abluft_ohne_Ueber)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.zu_nacht)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.ab_nacht)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer",self.tiefer_nb_dauer)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tiefer_zu_nacht)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tiefer_ab_nacht)
        # Neue Parameter
        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zuluft_min)
        wert_schreiben("IGF_RLT_Druckstufe", self.IGF_Druckstufe)
        wert_schreiben("IGF_RLT_AbluftminSummeLabor24h", self.abluft_labor_24h)
        wert_schreiben("IGF_RLT_AbluftminRaum24h", self.ABL_24h)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.ebene,
            self.zuluft_min,
            self.raum_druckstufe,
            self.ueberstroemung,
            self.abluft_ges,
            self.abluft_min,
            self.ABL_24h,
            self.abluft_ohne_24h,
            self.abluft_ohne_Ueber,
            self.ABL_Labor,
            self.abluft_labor_24h,
            self.nb_dauer,
            self.zu_nacht,
            self.ab_nacht,
            self.ab_nacht_ohne24h,
            self.tiefer_nb_dauer,
            self.tiefer_zu_nacht,
            self.tiefer_ab_nacht,
            self.tiefer_ab_nacht_ohne24h,
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
    columns=['Nummer', 'Name', 'Ebene', 'Zul_min', 'Druckstufe','Überstrom',
             'Abl_Gesamt', 'Abl_min', 'Abl_24h','Abl_ohne 24h',
             'Abl_ohne Über','Abl_Labor','abluft_labor_24h', 'nb Dauer',
             'nb zu', 'nb ab','nb ab_ohne24h','tief nb Dauer','tief nb zu',
             'tief nb ab','tief nb ab ohne 24h']
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb2:

        n_1 = 0

        with rpw.db.Transaction("Luftwechsel berechnen"):
            for mepraum in mepraum_liste:
                if pb2.cancelled:
                    script.exit()
                n_1 += 1
                pb2.update_progress(n_1, len(spaces))

                mepraum.werte_schreiben()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
