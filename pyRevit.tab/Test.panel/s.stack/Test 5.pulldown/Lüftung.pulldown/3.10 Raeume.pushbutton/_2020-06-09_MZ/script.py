# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time

start = time.time()


__title__ = "MFC RLT"
__doc__ = """Luftmengenberechnung"""
__author__ = "Tim Bartels"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
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
            ['kez', 'TGA_RLT_VolumenstromProNummer'],
            ['vol_faktor', 'TGA_RLT_VolumenstromProFaktor'],
            ['ABL_24h', 'TGA_RLT_AbluftSumme24h'],
            ['ABL_Labor', 'TGA_RLT_AbluftSummeLabor'],
            ['raum_druckstufe', 'TGA_RLT_RaumDruckstufeEingabe'],
            ['personen', 'Personenzahl'],
            ['ueberstroemung', 'TGA_RLT_RaumÜberströmungMenge'],
            ['nachtbetrieb', 'IGF_RLT_Nachtbetrieb'],
            ['NB_Von', 'IGF_RLT_NachtbetriebVon'],
            ['NB_Bis', 'IGF_RLT_NachtbetriebBis'],
            ['NB_LW', 'IGF_RLT_NachtbetriebLW'],
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

        self.zuluft_min = self.zuluft_min_berechnen()
        self.abluft_min = self.abluft_min_berechnen()

        self.nb_dauer = self.nb_dauer_berechnen()
        self.zu_nacht = self.zuluft_nacht_berechnen()
        self.ab_nacht = self.abluft_nacht_berechnen()

        self.tiefer_nb_dauer = self.tiefer_nb_dauer_berechnen()
        self.tiefer_zu_nacht = self.tiefer_zuluft_nacht_berechnen()
        self.tiefer_ab_nacht = self.tiefer_abluft_nacht_berechnen()

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

    def abluft_min_berechnen(self):
        abluft = 0

        if int(self.kez) in [
                berechnung_nach['Fläche'],
                berechnung_nach['Luftwechsel'],
                berechnung_nach['Person'],
                berechnung_nach['manuell']]:

            abluft = self.zuluft_min - self.raum_druckstufe

        elif int(self.kez) == berechnung_nach['nurABMa']:
            abluft = self.vol_faktor

        stroemt_ueber = [self.__get_value_in_project_units(s.LookupParameter("TGA_RLT_RaumÜberströmungMenge"))
                         for s in spaces_ueberstroemung
                         if s.LookupParameter('TGA_RLT_RaumÜberströmungAus').AsString() == self.nummer]

        logger.info("strömt_über raus: {}".format(sum(stroemt_ueber)))
        logger.info("strömt_über rein: {}".format(self.ueberstroemung))
        logger.info("ABL_ZR = {}".format(abluft))

        abluft = abluft + self.ueberstroemung - sum(stroemt_ueber)

        logger.info("ABL = {}".format(abluft))

        return abluft

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

        logger.info("ZUL_ZR = {}".format(zuluft))

        # Tim: in extra Funktion um Wiederholung zu vermeiden
        zuluft = self.__lab_abl_24h_druckstufe_pruefen(zuluft)

        logger.info("ZUL = {}".format(zuluft))

        return round(zuluft, 2)

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
        ab_nacht = 0

        if self.nachtbetrieb:
            ab_nacht = self.zu_nacht - self.raum_druckstufe

            if int(self.kez) == berechnung_nach['nurABMa']:
                ab_nacht = self.vol_faktor

        stroemt_ueber = [self.__get_value_in_project_units(s.LookupParameter("TGA_RLT_RaumÜberströmungMenge"))
                         for s in spaces_ueberstroemung
                         if s.LookupParameter('TGA_RLT_RaumÜberströmungAus').AsString() == self.nummer]

        logger.info("strömt_über rein: {}".format(self.ueberstroemung))
        logger.info("ABL_NACHT_ZR = {}".format(ab_nacht))

        ab_nacht = ab_nacht + self.ueberstroemung - sum(stroemt_ueber)
        logger.info("ABL_NACHT = {}".format(ab_nacht))

        return ab_nacht

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

        if self.tiefernachtbetrieb:
            tiefer_ab_nacht = self.tiefer_zu_nacht - self.raum_druckstufe

            if int(self.kez) == berechnung_nach['nurABMa']:
                tiefer_ab_nacht = self.vol_faktor

        stroemt_ueber = [self.__get_value_in_project_units(s.LookupParameter("TGA_RLT_RaumÜberströmungMenge"))
                         for s in spaces_ueberstroemung
                         if s.LookupParameter('TGA_RLT_RaumÜberströmungAus').AsString() == self.nummer]

        logger.info("strömt_über rein: {}".format(self.ueberstroemung))
        logger.info("Tiefer_ABL_NACHT_ZR = {}".format(tiefer_ab_nacht))

        tiefer_ab_nacht = tiefer_ab_nacht + self.ueberstroemung - sum(stroemt_ueber)
        logger.info("Tiefer_ABL_NACHT = {}".format(tiefer_ab_nacht))

        return tiefer_ab_nacht

    def tiefer_zuluft_nacht_berechnen(self):
        tiefer_zu_nacht = 0

        if self.nachtbetrieb:
            tiefer_zu_nacht = self.T_NB_LW * self.volumen

            # Tim: in extra Funktion um Wiederholung zu vermeiden
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

        wert_schreiben("TGA_RLT_ZuluftminRaum", self.zuluft_min)
        wert_schreiben("TGA_RLT_AbluftminRaum", self.abluft_min)
        wert_schreiben("TGA_RLT_AbluftSummeLabor24h", self.abluft_labor_24h)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.zu_nacht)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.ab_nacht)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer",self.tiefer_nb_dauer)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tiefer_zu_nacht)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tiefer_ab_nacht)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.zuluft_min,
            self.__get_element_attr("TGA_RLT_ZuluftminRaum"),
            self.abluft_min,
            self.__get_element_attr("TGA_RLT_AbluftminRaum"),
            self.abluft_labor_24h,
            self.__get_element_attr("TGA_RLT_AbluftSummeLabor24h"),
            self.nb_dauer,
            self.zu_nacht,
            self.ab_nacht,
            self.tiefer_nb_dauer,
            self.tiefer_zu_nacht,
            self.tiefer_ab_nacht
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
    title="Luftmengen",
    columns=['Nummer', 'Name', 'ZUL pyRevit', 'ZUL Excel', 'ABL pyRevit',
             'ABL Excel', 'Labor pyRevit', 'Labor Excel', 'nb Dauer',
             'nb zu', 'nb ab','tief nb Dauer','tief nb zu','tief nb ab']
)

logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with rpw.db.Transaction("Luftwechsel berechnen"):
        [mepraum.werte_schreiben() for mepraum in mepraum_liste]

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
