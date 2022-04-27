# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit, DB
from pyrevit import script, forms
from IGF_log import getlog
from IGF_libKopie import get_value
import time
from IGF_forms import Texteingeben

start = time.time()

__title__ = "2.50 Kühllast & Kühglleistung MEP Räume"
__doc__ = """berechnet Kühllast und Kühlleistung von MEP Räume

input Parameter:
--------------------
IGF_RLT_ZuluftminRaum: Zuluftmengen

IGF_RLT_ZuluftTemperatur: Zulufttemperatur

LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE: Zulufttemperatur falls IGF_RLT_ZuluftTemperatur nicht eingegeben wird

LIN_BA_DESIGN_COOLING_TEMPERATURE: Raumtemperatur

IGF_K_KühllastLaborRaum: Kühllast Labor Raum

IGF_S_KühllastLaborPWK: Kühllast für Laboreinrichtung über PKW

LIN_BA_CALCULATED_COOLING_LOAD: Kühllast Gebäude

IGF_K_DeS_Leistung: Kühlleistung DeS

IGF_K_ULK_Leistung: Kühlleistung ULK

IGF_K_KA_Leistung: Kühlleistung Kälte
--------------------

output Parameter:
--------------------
IGF_RLT_ZuluftKühlleistung: Kühlleistung Zuluft, Zuluftfaktor * (Vol_zu * 1000 * 1.2 * 1.006 * (Temp_Raum - Temp_Zu) / 3600)

IGF_K_KühlleistungRaum: Summe von Zuluft- & DeS- & ULK- & Kältekühlleistung

IGF_K_KühllastGesamt: Summe von Kühllast Gebäude und Kühllast Labor Raum

IGF_K_KühlleistungULK: KühllastGesamt - ZuluftKühlleistung

IGF_K_KühlBilanzProzent: gesamte Kühlleistung / gesamte Kühllast
--------------------

[Version: 1.2]
[2021.11.18]
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()

faktor = Texteingeben(text='0.8',label='Faktor der Zuluftleistung')
faktor.Title = __title__
try:
   faktor.ShowDialog()
except Exception as e:
    logger.error(e)
    script.exit()

Zuluftfaktor = faktor.text.Text
if Zuluftfaktor.find(','):
    Zuluftfaktor.replace(',','.')
try:
    if Zuluftfaktor == '0':
        Zuluftfaktor = 0.8
    Zuluftfaktor = float(Zuluftfaktor)
    

except:
    logger.error('Falsche Faktor')
    script.exit()

class MEPRaum:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.name = self.get_parameter('Name')
        self.nummer = self.get_parameter('Nummer')
        self.Vol_zu = self.get_parameter('IGF_RLT_ZuluftminRaum')
        self.T_raum = self.get_parameter('LIN_BA_DESIGN_COOLING_TEMPERATURE')
        self.Kuehllast_Labor_Raum = self.get_parameter('IGF_K_KühllastLaborRaum')
        self.Kuehllast_Labor_PWK = self.get_parameter('IGF_S_KühllastLaborPWK')
        self.Kuehllast_Gebaeude = self.get_parameter('LIN_BA_CALCULATED_COOLING_LOAD')
        self.KL_DeS = self.get_parameter('IGF_K_DeS_Leistung')
        self.KL_ULK = self.get_parameter('IGF_K_ULK_Leistung')
        self.KL_KA = self.get_parameter('IGF_K_KA_Leistung')
        self.raumtyp = self.element.LookupParameter('Bedingungstyp').AsValueString()

        try:
            self.T_zu = self.get_parameter('IGF_RLT_ZuluftTemperatur')
            if self.T_zu == '0' or self.T_zu == None:
                try:
                    self.T_zu = self.get_parameter('LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE')
                except:
                    self.T_zu = -273.15
                    logger.error('kein Zulufttemperatur eingegeben')
            else:
                pass
        except:
            try:
                self.T_zu = self.get_parameter('LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE')
            except:
                self.T_zu = -273.15
                logger.error('kein Zulufttemperatur eingegeben')

        if self.raumtyp in ['Gekühlt','Beheizt und gekühlt']:
            self.P_zu = self.zuluft_kuelleistung_berechnen()
        else:
            self.P_zu = 0

        self.Kuehllast_gesamt = self.KuehllastGesamt_Berechnen()
        self.P_ULK = self.KuelleistungULK_Berechnen()
        self.KL_gesamt = self.KL_gesamt_berechnen()
        self.Prozent = self.Prozent_Berechnen()


    def Werte_Pruefen(self,wert):
        if not wert:
            wert = 0
        return wert

    def get_parameter(self, param_name):
        param = self.element.LookupParameter(param_name)
        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return get_value(param)

    def zuluft_kuelleistung_berechnen(self):
        Kuelhlleistung = 0
        if self.Vol_zu and self.T_zu > -273.15 and self.T_raum > -273.15:
            Kuelhlleistung = Zuluftfaktor * (self.Vol_zu * 1000 * 1.2 * 1.006 * (self.T_raum - self.T_zu) / 3600)
        return round(Kuelhlleistung, 2)
    
    def KL_gesamt_berechnen(self):
        Kuelhlleistung = 0
        self.KL_DeS = self.Werte_Pruefen(self.KL_DeS)
        self.KL_ULK = self.Werte_Pruefen(self.KL_ULK)
        self.KL_KA = self.Werte_Pruefen(self.KL_KA)
        self.P_zu = self.Werte_Pruefen(self.P_zu)

        Kuelhlleistung = self.KL_DeS + self.KL_ULK + self.KL_KA + self.P_zu

        return round(Kuelhlleistung, 2)

    def KuehllastGesamt_Berechnen(self):
        KuehllastGesamt = 0

        try:
            KuehllastGesamt = self.Kuehllast_Gebaeude + self.Kuehllast_Labor_Raum
        except:
            try:
                KuehllastGesamt = self.Kuehllast_Gebaeude + 0.0
            except:
                try:
                    KuehllastGesamt = self.Kuehllast_Labor_Raum + 0.0
                except:
                    KuehllastGesamt = 0.0

        KuehllastGesamt = self.Kuehllast_Gebaeude + self.Kuehllast_Labor_Raum

        return KuehllastGesamt

    def KuelleistungULK_Berechnen(self):

        KuelleistungULK = 0
        KuelleistungULK = self.Kuehllast_gesamt - self.P_zu
        
        return KuelleistungULK
    
    def Prozent_Berechnen(self):
        prozent = 0
        if not self.Kuehllast_gesamt:
            return 0

        prozent = float(self.KL_gesamt) / self.Kuehllast_gesamt
        return round(prozent,3)
        


    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                # logger.info(
                #     "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                self.element.LookupParameter(
                    param_name).SetValueString(str(wert))

        wert_schreiben("IGF_RLT_ZuluftKühlleistung", self.P_zu)
        wert_schreiben("IGF_K_KühllastGesamt", self.Kuehllast_gesamt)
        wert_schreiben("IGF_K_KühlleistungULK", self.P_ULK)
        wert_schreiben("IGF_K_KühlleistungRaum", self.KL_gesamt)
        self.element.LookupParameter('IGF_K_KühlBilanzProzent').Set(self.Prozent)



mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} MEP-Räume',cancellable=True, step=10) as pb:
    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)


# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Werte schreiben')

        for n,mepraum in enumerate(mepraum_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mepraum_liste))
            mepraum.werte_schreiben()
        t.Commit()