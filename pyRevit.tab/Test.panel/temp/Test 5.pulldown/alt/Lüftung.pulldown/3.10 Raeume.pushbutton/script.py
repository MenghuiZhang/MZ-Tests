# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit,DB
from pyrevit import script, forms
import time
from IGF_log import getlog
from IGF_lib import get_value

start = time.time()

__title__ = "3.1 Raumluft_Raum"
__doc__ = """Luftmengenberechnung

[30.09.2021]
Version: 1.1"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
active_view = doc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc,active_view.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

spaces_ueberstroemung2 = []
for ele in spaces_collector:
    summe2 = get_value(ele.LookupParameter("TGA_RLT_RaumÜberströmungAus"))
    if not summe2 in [None,0]:
        spaces_ueberstroemung2.append(ele)

berechnung_nach = {
    "Fläche": '1',
    "Luftwechsel": '2',
    "Person": '3',
    "manuell": '4',
    "nurZUMa": '5',
    "nurABMa": '6',
    "nurZU_Fläche": '5.1',
    "nurZU_Luftwechsel": '5.2',
    "nurZU_Person": '5.3',
    "nurAB_Fläche": '6.1',
    "nurAB_Luftwechsel": '6.2',
    "nurAB_Person": '6.3',
    "keine": '9'
}

def luft_round(luft):
    zahl = luft%5
    if zahl < 5 and zahl > 0:
        return (int(luft/5)+1) * 5
    else:
        return luft

class MEPRaum:
    def __init__(self, elem_id):
        self.elem_id = elem_id
        self.elem = doc.GetElement(self.elem_id)
        attr = [
            ["name", "Name"],
            ["nummer", "Nummer"],
            ["flaeche", "Fläche"],
            ["volumen", "Volumen"],
            ["personen", "Personenzahl"],

            ["vol_faktor", "TGA_RLT_VolumenstromProFaktor"],
            ["raum_druckstufe", "IGF_RLT_RaumDruckstufeEingabe"],
            ["ueberstroemungIn", "IGF_RLT_ÜberströmungSummeIn"],
            ["ueberstroemungAus", "IGF_RLT_ÜberströmungSummeAus"],
            ["ueberstroemung2", "TGA_RLT_RaumÜberströmungMenge"],

            ["ABL_24h", "IGF_RLT_AbluftSumme24h"],
            ["ABLmin_Labor", "IGF_RLT_AbluftminSummeLabor"],
            ["ABLmax_Labor", "IGF_RLT_AbluftmaxSummeLabor"],

            # Nachbetrieb
            ["nachtbetrieb", "IGF_RLT_Nachtbetrieb"],
            ["NB_Von", "IGF_RLT_NachtbetriebVon"],
            ["NB_Bis", "IGF_RLT_NachtbetriebBis"],
            ["NB_LW", "IGF_RLT_NachtbetriebLW"],

            # Tiefnachtbetrieb
            ["tiefernachtbetrieb", "IGF_RLT_TieferNachtbetrieb"],
            ["T_NB_Von", "IGF_RLT_TieferNachtbetriebVon"],
            ["T_NB_Bis", "IGF_RLT_TieferNachtbetriebBis"],
            ["T_NB_LW", "IGF_RLT_TieferNachtbetriebLW"]
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(revit_name))

        self.stroemt_ueber = [get_value(s.LookupParameter("TGA_RLT_RaumÜberströmungMenge"))
                         for s in spaces_ueberstroemung2
                         if s.LookupParameter('TGA_RLT_RaumÜberströmungAus').AsString() == self.nummer]
        self.ueberAusMa = sum(self.stroemt_ueber)

        self.ab_min_labor_und_24h = self.ABLmin_Labor + self.ABL_24h
        self.ab_max_labor_und_24h = self.ABLmax_Labor + self.ABL_24h

        self.flaeche = round(self.flaeche, 0)

        self.kez = self.elem.LookupParameter("TGA_RLT_VolumenstromProNummer").AsValueString()
        self.hinweis = ''

        self.Luftmengen_berechnen()

        self.angeabluft = self.angezuluft

        self.abluft_ges = self.Abluft_Ges_Berechnen(self.zuluft_max)

        self.nb_dauer = self.nb_dauer_berechnen()
        self.zu_nacht = self.zuluft_nacht_berechnen()
        self.ab_nacht,self.ab_nacht_ohne24h = self.abluft_nacht_berechnen()

        self.tiefer_nb_dauer = self.tiefer_nb_dauer_berechnen()
        self.tiefer_zu_nacht = self.tiefer_zuluft_nacht_berechnen()
        self.tiefer_ab_nacht,self.tiefer_ab_nacht_ohne24h = self.tiefer_abluft_nacht_berechnen()

        self.IGF_Druckstufe = self.Druckstufe_Berechnen()
        self.IGF_Legende = self.DruckstufeLegende_Berechnen()


    def get_element_attr(self, param_name):
        param = self.elem.LookupParameter(param_name)
        if not param:
            logger.error("Parameter {} konnte nicht gefunden werden".format(param_name))
            return
        return get_value(param)
    
    def Labmin_24h_Druckstufe_Pruefen(self,zuluftmin):
        if self.ab_min_labor_und_24h > zuluftmin:
            logger.info("{} m/h3 Laborminabluft und 24h Abluft berücksichtigt".format(self.ab_min_labor_und_24h))
            zuluftmin = self.ab_min_labor_und_24h

        if self.raum_druckstufe > 0:
            logger.info("{} - Druckstufe berücksichtigt".format(self.raum_druckstufe))
            zuluftmin = zuluftmin + self.raum_druckstufe

        return zuluftmin

    def Labmax_24h_Druckstufe_Pruefen(self,zuluftmax):
        if self.ab_max_labor_und_24h > zuluftmax:
            logger.info("{} m/h3 Labormaxabluft und 24h Abluft berücksichtigt".format(self.ab_max_labor_und_24h))
            zuluftmax = self.ab_max_labor_und_24h

        if self.raum_druckstufe > 0:
            logger.info("{} - Druckstufe berücksichtigt".format(self.raum_druckstufe))
            zuluftmax = zuluftmax + self.raum_druckstufe

        return zuluftmax

    def Luftmengen_berechnen(self):
        self.zuluft_min = 0
        self.zuluft_max = 0
        self.angezuluft = 0
        self.abluft_min = 0
        self.abluft_max = 0

        if self.kez == berechnung_nach["nurZU_Fläche"]:
            logger.info("Berechnung nach Fläche. Nur Zuluft")
            if self.flaeche == 0:
                logger.error("Die Fläche des Raumes {} {} ist 0m2".format(self.nummer, self.name))
                return self.zuluft_min,self.zuluft_max,self.angezuluft

            self.zuluft_min = self.flaeche * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max

            abweichung = self.ueberstroemungAus - self.zuluft_max - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa
            if self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 < self.ueberstroemungAus + self.ueberAusMa:
                self.hinweis = "Abluft durch Überströmung | Abweichung: +{} m3/h".format(int(abweichung))
                self.zuluft_max = self.ueberstroemungAus - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa

            else:
                self.hinweis = "Achtung: Nur Zuluft, aber die Überströmung ist kleiner als die Zuluft. Abweichung: {} m3/h".format(int(abweichung))

            
            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h


        elif self.kez == berechnung_nach["nurZU_Luftwechsel"]:
            logger.info("Berechnung nach Luftwechsel. Nur Zuluft")

            self.zuluft_min = self.volumen * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max
            
            abweichung = self.ueberstroemungAus - self.zuluft_max - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa

            if self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 < self.ueberstroemungAus + self.ueberAusMa:
                self.hinweis = "Abluft durch Überströmung | Abweichung: +{} m3/h".format(int(abweichung))
                self.zuluft_max = self.ueberstroemungAus - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa

            else:
                self.hinweis = "Achtung: Nur Zuluft, aber die Überströmung ist kleiner als die Zuluft. Abweichung: {} m3/h".format(int(abweichung))

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            
            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez == berechnung_nach["nurZU_Person"]:
            self.zuluft_min = self.personen * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max
            abweichung = self.ueberstroemungAus - self.zuluft_max - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa
            if self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 < self.ueberstroemungAus + self.ueberAusMa:
                self.hinweis = "Abluft durch Überströmung | Abweichung: +{} m3/h".format(int(abweichung))
                self.zuluft_max = self.ueberstroemungAus - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa

            else:
                self.hinweis = "Achtung: Nur Zuluft, aber die Überströmung ist kleiner als die Zuluft. Abweichung: {} m3/h".format(int(abweichung))

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h
            
        elif self.kez == berechnung_nach["nurAB_Fläche"]:
            logger.info("Berechnung nach Fläche. Nur Abluft")
            if self.flaeche == 0:
                logger.error("Die Fläche des Raumes {} {} ist 0m2".format(self.nummer, self.name))
                return self.zuluft_min,self.zuluft_max,self.angezuluft
            self.zuluft_min = self.flaeche * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.angezuluft = self.zuluft_max
            abweichung = self.zuluft_max - self.ueberstroemungIn - self.ueberstroemung2

            if self.zuluft_max < self.ueberstroemungIn + self.ueberstroemung2:
                self.hinweis = "Zuluft durch Überströmung | Abweichung: {} m3/h".format(int(abweichung))
                self.zuluft_min = self.zuluft_max = 0

            else:
                self.hinweis = "Achtung: Nur Abluft, aber die Überströmung ist kleiner als die Zuluft. Abweichung: {} m3/h".format(int(abweichung))

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            if self.abluft_min < 0:
                self.zuluft_min = self.zuluft_min - self.abluft_min
                self.abluft_min = 0
            
            if self.abluft_max < 0:
                self.zuluft_max = self.zuluft_max - self.abluft_max
                self.abluft_max = 0

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez == berechnung_nach["nurAB_Luftwechsel"]:
            self.zuluft_min = self.volumen * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max
            abweichung = self.zuluft_max - self.ueberstroemungIn - self.ueberstroemung2

            if self.zuluft_max < self.ueberstroemungIn + self.ueberstroemung2:
                self.hinweis = "Zuluft durch Überströmung | Abweichung: {} m3/h".format(int(abweichung))
                self.zuluft_min = self.zuluft_max = 0

            else:
                self.hinweis = "Achtung: Nur Abluft, aber die Überströmung ist kleiner als die Zuluft. Abweichung: {} m3/h".format(int(abweichung))

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            if self.abluft_min < 0:
                self.zuluft_min = self.zuluft_min - self.abluft_min
                self.abluft_min = 0
            
            if self.abluft_max < 0:
                self.zuluft_max = self.zuluft_max - self.abluft_max
                self.abluft_max = 0

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez == berechnung_nach["nurAB_Person"]:
            self.zuluft_min = self.personen * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max
            abweichung = self.zuluft_max - self.ueberstroemungIn - self.ueberstroemung2
            if self.zuluft_max < self.ueberstroemungIn + self.ueberstroemung2:
                self.hinweis = "Zuluft durch Überströmung | Abweichung: {} m3/h".format(int(abweichung))
                self.zuluft_min = self.zuluft_max = 0

            else:
                self.hinweis = "Achtung: Nur Abluft, aber die Überströmung ist kleiner als die Zuluft. Abweichung: {} m3/h".format(int(abweichung))

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            if self.abluft_min < 0:
                self.zuluft_min = self.zuluft_min - self.abluft_min
                self.abluft_min = 0
            
            if self.abluft_max < 0:
                self.zuluft_max = self.zuluft_max - self.abluft_max
                self.abluft_max = 0
                
            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez == berechnung_nach["Fläche"]:
            logger.info("Berechnung nach Fläche")
            if self.flaeche == 0:
                logger.error("Die Fläche des Raumes {} {} ist 0m2".format(self.nummer, self.name))
                return self.zuluft_min,self.zuluft_max,self.angezuluft

            self.zuluft_min = self.flaeche * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            if self.abluft_min < 0:
                self.zuluft_min = self.zuluft_min - self.abluft_min
                self.abluft_min = 0
            
            if self.abluft_max < 0:
                self.zuluft_max = self.zuluft_max - self.abluft_max
                self.abluft_max = 0

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez == berechnung_nach["Luftwechsel"]:
            logger.info("Berechnung nach Luftwechsel")
            self.zuluft_min = self.volumen * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            if self.abluft_min < 0:
                self.zuluft_min = self.zuluft_min - self.abluft_min
                self.abluft_min = 0
            
            if self.abluft_max < 0:
                self.zuluft_max = self.zuluft_max - self.abluft_max
                self.abluft_max = 0

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez == berechnung_nach["Person"]:
            logger.info("Berechnung nach Personen")
            self.zuluft_min = self.personen * self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max

            self.abluft_min = self.zuluft_min + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa
            self.abluft_max = self.zuluft_max + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.raum_druckstufe - self.ueberAusMa

            if self.abluft_min < 0:
                self.zuluft_min = self.zuluft_min - self.abluft_min
                self.abluft_min = 0
            
            if self.abluft_max < 0:
                self.zuluft_max = self.zuluft_max - self.abluft_max
                self.abluft_max = 0

            self.abluft_min = luft_round(self.abluft_min)
            self.abluft_max = luft_round(self.abluft_max)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez  == berechnung_nach["manuell"]:
            self.zuluft_min = self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max

            self.abluft_min = self.vol_faktor - self.raum_druckstufe + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.ueberAusMa
            
            self.abluft_min = self.abluft_max = luft_round(self.abluft_min)
            self.abluft_ohne_24h = self.abluft_max - self.ABL_24h

        elif self.kez  == berechnung_nach["nurZUMa"]:
            self.zuluft_min = self.vol_faktor
            self.zuluft_min = self.zuluft_max = luft_round(self.zuluft_min)
            self.zuluft_min = self.Labmin_24h_Druckstufe_Pruefen(self.zuluft_min)
            self.zuluft_max = self.Labmax_24h_Druckstufe_Pruefen(self.zuluft_max)
            self.angezuluft = self.zuluft_max

        elif self.kez == berechnung_nach["nurABMa"]:
            self.angezuluft = self.vol_faktor
            self.angezuluft = luft_round(self.angezuluft)
            self.abluft_min = self.abluft_max = self.vol_faktor


    def Abluft_Ges_Berechnen(self,Zuluft):
        Abluft_Ges = 0

        if self.kez in [
            berechnung_nach["Fläche"],
            berechnung_nach["Luftwechsel"],
            berechnung_nach["Person"],
            berechnung_nach["nurAB_Fläche"],
            berechnung_nach["nurAB_Luftwechsel"],
            berechnung_nach["nurAB_Person"],]:

            Abluft_Ges = Zuluft - self.raum_druckstufe

        elif self.kez in [
            berechnung_nach["nurZU_Fläche"],
            berechnung_nach["nurZU_Luftwechsel"],
            berechnung_nach["nurZU_Person"],]:
            if self.zuluft_min == self.ueberstroemungAus - self.ueberstroemungIn - self.ueberstroemung2 + self.ueberAusMa:
                Abluft_Ges = 0
            else:
                Abluft_Ges = Zuluft - self.raum_druckstufe

        elif self.kez == berechnung_nach["nurABMa"]:
            Abluft_Ges = self.vol_faktor

        elif self.kez == berechnung_nach["manuell"]:
            Abluft_Ges = self.vol_faktor - self.raum_druckstufe

        if Abluft_Ges > 0:
            Abluft_Ges = luft_round(Abluft_Ges)

        return Abluft_Ges

    def Druckstufe_Berechnen(self):
        DS = 0
        DS = self.zuluft_max - self.abluft_max
        return DS

    def DruckstufeLegende_Berechnen(self):
        Lege = '0'
        n = 0

        if self.raum_druckstufe > 0:
            n = int(self.raum_druckstufe/10)
            if self.raum_druckstufe <= 50:
                Lege = n * '+'
            else:
                Lege = str(n) + '+'
        else:
            n = int(abs(self.raum_druckstufe)/10)
            if self.raum_druckstufe >= -50:
                Lege = n * '-'
            else:
                Lege = str(n) + '-'
        return Lege

    def nb_dauer_berechnen(self):
        nb_dauer = 0

        if self.nachtbetrieb:

            if self.tiefernachtbetrieb:
                nb_dauer = self.NB_Bis - self.NB_Von - self.T_NB_Bis + self.T_NB_Von

            else:
                nb_dauer = self.NB_Bis - self.NB_Von + 24.00

        return nb_dauer

    def abluft_nacht_berechnen(self):
        ab_nacht_ohne24h = 0
        ab_nacht_ges = 0

        if self.nachtbetrieb:
            ab_nacht_ges = self.Abluft_Ges_Berechnen(self.zu_nacht)
            ab_nacht_ohne24h = ab_nacht_ges - self.ABL_24h

        return luft_round(ab_nacht_ges),luft_round(ab_nacht_ohne24h)

    def zuluft_nacht_berechnen(self):
        zu_nacht = 0

        if self.nachtbetrieb:
            zu_nacht = self.NB_LW * self.volumen

            zu_nacht = self.__lab_abl_24h_druckstufe_pruefen(zu_nacht)

        return luft_round(zu_nacht)

    def tiefer_nb_dauer_berechnen(self):
        tiefer_nb_dauer = 0

        if self.tiefernachtbetrieb:
            tiefer_nb_dauer = self.T_NB_Bis - self.T_NB_Von + 24.00

        return tiefer_nb_dauer

    def tiefer_abluft_nacht_berechnen(self):
        tiefer_ab_nacht = 0
        tiefer_ab_nacht_ohne24h = 0

        if self.tiefernachtbetrieb:
            tiefer_ab_nacht = self.Abluft_Ges_Berechnen(self.tiefer_zu_nacht)
            tiefer_ab_nacht_ohne24h = tiefer_ab_nacht - self.ABL_24h

        return luft_round(tiefer_ab_nacht),luft_round(tiefer_ab_nacht_ohne24h)

    def tiefer_zuluft_nacht_berechnen(self):
        tiefer_zu_nacht = 0

        if self.nachtbetrieb:
            tiefer_zu_nacht = self.T_NB_LW * self.volumen

            tiefer_zu_nacht = self.__lab_abl_24h_druckstufe_pruefen(tiefer_zu_nacht)

        return luft_round(tiefer_zu_nacht)


    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.elem.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)

        
        wert_schreiben("IGF_RLT_AbluftminSummeLabor24h", self.ab_min_labor_und_24h)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor24h", self.ab_max_labor_und_24h)

        wert_schreiben("IGF_RLT_AbluftminRaum", self.abluft_min)
        wert_schreiben("IGF_RLT_AbluftmaxRaum", self.abluft_max)

        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zuluft_min)
        wert_schreiben("IGF_RLT_ZuluftmaxRaum", self.zuluft_max)

        wert_schreiben("Angegebener Zuluftstrom", self.angezuluft)
        wert_schreiben("Angegebener Abluftluftstrom", self.angeabluft)

        wert_schreiben("IGF_RLT_AbluftminRaumGes", self.abluft_ges)

        # Berechnen für Anlagen
        wert_schreiben("IGF_RLT_AnlagenRaumAbluft", self.abluft_ohne_24h)
        wert_schreiben("IGF_RLT_AnlagenRaumZuluft", self.zuluft_max)
        wert_schreiben("IGF_RLT_AnlagenRaum24hAbluft", self.ABL_24h)

        wert_schreiben("IGF_RLT_RaumBilanz", self.IGF_Druckstufe)
        wert_schreiben("IGF_RLT_RaumDruckstufeLegende", self.IGF_Legende)
        wert_schreiben("IGF_RLT_Hinweis", self.hinweis)

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
            self.ab_min_labor_und_24h,
            self.ab_max_labor_und_24h,
            self.zuluft_min,
            self.zuluft_max,
            self.abluft_min,
            self.abluft_max,
            self.ABL_24h,
            self.abluft_ges,
            self.angezuluft,
            self.angeabluft,
            self.zuluft_max,
            self.abluft_ohne_24h,
            self.ABL_24h,
            self.nb_dauer,
            self.zu_nacht,
            self.ab_nacht,
            self.ab_nacht_ohne24h,
            self.tiefer_nb_dauer,
            self.tiefer_zu_nacht,
            self.tiefer_ab_nacht,
            self.tiefer_ab_nacht_ohne24h,
            self.IGF_Legende
        ]

table_data = []
mepraum_liste = []
with forms.ProgressBar(title="{value}/{max_value} Luftmengenberechnung",cancellable=True, step=10) as pb:
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
    columns=["Nummer", "Name", "Labor_min", "Lab_max", "Zuluft_min","Zuluft_max", "Abluft_min", "Abluft_max", "Abluft_24h","Abluft_Gesamt",
             "Angegebener Zuluft" ,'Angegebener Abluft', "AnlagenRaumZuluft",'AnlagenRaumAbluft','AnlagenRaum24hAbluft', "nb Dauer",
             "nb zu", "nb ab","nb ab_ohne24h","tief nb Dauer","tief nb zu","tief nb ab","tief nb ab ohne 24h","RaumdruckstufeLegende"]
)

# Werte zuückschreiben + Abfrage
if forms.alert("Berechnete Werte in Modell schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Werte schreiben",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Luftmengen')
        for n,mepraum in enumerate(mepraum_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(spaces))
            mepraum.werte_schreiben()
        t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))