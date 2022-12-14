# coding: utf8
import clr
from rpw import revit,DB
from Autodesk.Revit.UI import ExternalEvent
from pyrevit import script, forms
from IGF_log import getlog
from System.Collections.ObjectModel import ObservableCollection
from eventhandler import ResizeAnsicht,ResetAnsicht,Luftauslass,ShowElement,HighLightElement,\
    VSR,DICT_MEP_ITEMSSOIRCE,VISIBLE,HIDDEN,ModifierKeys,Keyboard,Key,ChangeParameterWert,\
    RED,BLACK,UNDO,UPDATEMODEL,VERTEILEN,VERTEILEN_MEP,BERECHNEN

__title__ = "Raumluftbilanz"
__doc__ = """

imput Parameter:
-------------------------
Fläche: Raumfläche
Volumen: Raumvolumen
Personenzahl: Anzahl der Personen für Luftmengenberechnung
TGA_RLT_VolumenstromProFaktor: Volumenstromfaktor pro [Faktor oder m3/h]
IGF_RLT_RaumDruckstufeEingabe: RaumDruckstufe
IGF_RLT_ÜberströmungSummeIn: Überstromluft einströmend
IGF_RLT_ÜberströmungSummeAus: Überstromluft ausströmend
TGA_RLT_RaumÜberströmungMenge: Menge der Überströmung
IGF_RLT_AbluftSumme24h: 24h Abluft
IGF_RLT_AbluftminSummeLabor: min. Laborabluft
IGF_RLT_AbluftmaxSummeLabor: max. Laborabluft
IGF_RLT_Nachtbetrieb: Nachtbetrieb
IGF_RLT_NachtbetriebVon: Beginn des Nachbetriebs
IGF_RLT_NachtbetriebBis: Ende des Nachbetriebs
IGF_RLT_NachtbetriebLW: Luftwechsel für Nacht [-1/h]
IGF_RLT_TieferNachtbetrieb: Tiefnachtbetrieb
IGF_RLT_TieferNachtbetriebVon: Beginn des Tiefnachtbetrieb
IGF_RLT_TieferNachtbetriebBis: Ende des Tiefnachtbetrieb
IGF_RLT_TieferNachtbetriebLW: Luftwechsel für Tiefnacht [-1/h]
TGA_RLT_VolumenstromProNummer: Kennziffer für Luftmengenberechnung
TGA_RLT_RaumÜberströmungAus: Überströmung aus Raum
IGF_RLT_AbluftminSummeLabor24h: IGF_RLT_AbluftSumme24h + IGF_RLT_AbluftminSummeLabor
IGF_RLT_AbluftmaxSummeLabor24h: IGF_RLT_AbluftSumme24h + IGF_RLT_AbluftmaxSummeLabor
IGF_RLT_AbluftminRaum: min. Raumabluft, ohne Anteil der Überströmung
IGF_RLT_AbluftmaxRaum: max. Raumabluft, ohne Anteil der Überströmung
IGF_RLT_ZuluftminRaum: min. Raumzuluft
IGF_RLT_ZuluftmaxRaum: max. Raumzuluft
TGA_RLT_VolumenstromProEinheit: Einheit
Angegebener Zuluftstrom:
Angegebener Abluftluftstrom:
IGF_RLT_AbluftminRaumGes:
IGF_RLT_AnlagenRaumAbluft: Raumabluft für Anlagenberechnung. ohne Anteil der 24h Abluft
IGF_RLT_AnlagenRaumZuluft: Raumzuluft für Anlagenberechnung
IGF_RLT_AnlagenRaum24hAbluft: 24h Abluft für Anlagenberechnung
IGF_RLT_RaumBilanz: Bilanz der aller Anschlüsse im Raum
IGF_RLT_RaumDruckstufeLegende: Raum Druckstufe Legende
IGF_RLT_Hinweis: Hinweis
IGF_RLT_NachtbetriebDauer: Dauer des Nachbetriebs
IGF_RLT_ZuluftNachtRaum: Zuluftmengen des Nachbetriebs
IGF_RLT_AbluftNachtRaum: Abluftmengen des Nachbetriebs
IGF_RLT_TieferNachtbetriebDauer: Dauer des Tiefnachtbetrieb
IGF_RLT_ZuluftTieferNachtRaum: Zuluftmenegen des Tiefnachtbetrieb
IGF_RLT_AbluftTieferNachtRaum: Abluftmengen des Tiefnachtbetrieb

-------------------------

[2021.11.22]
Version: 1.2
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc


views = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.View3D)).ToElements()
for el in views:
    if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
        uidoc.ActiveView = el
        break

active_view = uidoc.ActiveView

class MEPRaum_Uebersicht(forms.WPFWindow):
    def __init__(self):
        self.red = RED
        self.black = BLACK
        self.elemId4hightlight = None
        self.resizeAnsicht = ResizeAnsicht()
        self.updateModel = UPDATEMODEL()
        self.resetAnsicht = ResetAnsicht()
        self._undo = UNDO()
        self.highlightelement = HighLightElement()
        self.showelement = ShowElement()
        self.changeparameterwert = ChangeParameterWert()
        self.verteilen = VERTEILEN()
        self.verteilenmep = VERTEILEN_MEP()
        self.berechnen = BERECHNEN()
        self.Trans = 0
        self.TransMEPRaum = []
        self.BerechnenEvent = ExternalEvent.Create(self.berechnen)
        self.verteilenmepEvent = ExternalEvent.Create(self.verteilenmep)
        self.verteilenEvent = ExternalEvent.Create(self.verteilen)
        self.updateModelEvent = ExternalEvent.Create(self.updateModel)
        self.undoEvent = ExternalEvent.Create(self._undo)
        self.changeparameterwertEvent = ExternalEvent.Create(self.changeparameterwert)
        self.resizeAnsichtEvent = ExternalEvent.Create(self.resizeAnsicht)
        self.resetAnsichtEvent = ExternalEvent.Create(self.resetAnsicht)
        self.highlightelementEvent = ExternalEvent.Create(self.highlightelement)
        self.showelementEvent = ExternalEvent.Create(self.showelement)

        self.originalboxmax = DB.XYZ(active_view.GetSectionBox().Max.X,active_view.GetSectionBox().Max.Y,active_view.GetSectionBox().Max.Z)
        self.originalboxmin = DB.XYZ(active_view.GetSectionBox().Min.X,active_view.GetSectionBox().Min.Y,active_view.GetSectionBox().Min.Z)

        self.visible = VISIBLE
        self.hidden = HIDDEN

        self.Keyboard = Keyboard
        self.ModifierKeys = ModifierKeys
        self.Key = Key

        forms.WPFWindow.__init__(self,'MEP.xaml')
        self.mepraum_liste = DICT_MEP_ITEMSSOIRCE

        self.Raumnr.ItemsSource = sorted(self.mepraum_liste.keys())

        self.mepraum = None
        self.temp_luftauslass = ObservableCollection[Luftauslass]()
        self.warnung.Visibility= self.hidden

    def set_up(self):
        self.bezugsname.IsEnabled = True
        self.bezugsname.ItemsSource = sorted(self.mepraum.berechnung_nach.values())
        self.faktor.IsEnabled = True
        self.isnachtbetrieb.IsEnabled = True
        self.LW_nacht.IsEnabled = True
        self.von_nacht.IsEnabled = True
        self.bis_nacht.IsEnabled = True
        self.istiefenachtbetrieb.IsEnabled = True
        self.LW_tnacht.IsEnabled = True
        self.von_tnacht.IsEnabled = True
        self.bis_tnacht.IsEnabled = True
        self.labmineingabe.IsEnabled = True
        self.labmaxeingabe.IsEnabled = True
        
        self.ab24heingabe.IsEnabled = True
        self.druckeingabe.IsEnabled = True


        self.isnachtbetrieb.IsChecked = True if self.mepraum.nachtbetrieb != 0 else False
        self.istiefenachtbetrieb.IsChecked = True if self.mepraum.tiefenachtbetrieb != 0 else False
        self.LW_nacht.Text = str(self.mepraum.NB_LW)
 
        self.von_nacht.Text = str(self.mepraum.nb_von.soll)
        self.bis_nacht.Text = str(self.mepraum.nb_bis.soll)
        self.LW_nacht.Text = str(self.mepraum.T_NB_LW)
        self.von_nacht.Text = str(self.mepraum.tnb_von.soll)
        self.bis_nacht.Text = str(self.mepraum.tnb_bis.soll)
        self.labmineingabe.Text = str(self.mepraum.ab_lab_min.soll)      
        self.labmaxeingabe.Text = str(self.mepraum.ab_lab_max.soll)
        self.ab24heingabe.Text = str(self.mepraum.ab_24h.soll)
        self.druckeingabe.Text = str(self.mepraum.Druckstufe.soll)


        self.ebene.Text = str(self.mepraum.ebene)
        self.bezugsname.Text = str(self.mepraum.bezugsname)
        self.faktor.Text = str(self.mepraum.faktor)
        self.einheit.Text = str(self.mepraum.einheit)
        self.flaeche.Text = str(self.mepraum.flaeche)
        self.volumen.Text = str(self.mepraum.volumen)
        self.personen.Text = str(self.mepraum.personen)
        self.hoehe.Text = str(self.mepraum.hoehe)
        self.grundinfo.ItemsSource = self.mepraum.Uebersicht
        self.grundinfo.Items.Refresh()
        self.anlageninfo.ItemsSource = self.mepraum.Anlagen_info
        self.anlageninfo.Items.Refresh()
        self.schachtinfo.ItemsSource = self.mepraum.Schacht_info
        self.schachtinfo.Items.Refresh()

        self.lv_vsr.ItemsSource = self.mepraum.list_vsr
        self.lv_auslass.ItemsSource = self.mepraum.list_auslass

        self.lv_ueber_aus.ItemsSource = self.mepraum.list_ueber['Aus']
        self.lv_ueber_in.ItemsSource = self.mepraum.list_ueber['Ein']

        self.auswertung_nachtbetrieb.Visibility = self.visible

        try:self.Auswertung_MEP()
        except:pass
        try:self.Auswertung_System()
        except:pass

    def Auswertung_System(self):
        self.mepraum.get_Anlagen_info()

        uber_in_m = self.mepraum.ueber_in_manuell.soll
        uber_aus_m = self.mepraum.ueber_aus_manuell.soll
        uber_in = 0
        uber_aus = 0
        ab24h = 0

        lab_min = 0
        lab_max = 0
        lab_nb_ab = 0
        lab_tnb_ab = 0

        min_zu_raum = 0
        max_zu_raum = 0
        nb_zu_raum = 0
        tnb_zu_raum = 0

        min_ab_raum = 0
        nb_ab_raum = 0
        max_ab_raum = 0
        tnb_ab_raum = 0

        min_zu_tier = 0
        nb_zu_tier = 0
        max_zu_tier = 0
        tnb_zu_tier = 0

        min_ab_tier = 0
        nb_ab_tier = 0
        max_ab_tier = 0
        tnb_ab_tier = 0

        self.min_auswertung_druckstufe.Text = str(int(round(self.mepraum.Druckstufe.soll)))
        self.nacht_auswertung_druckstufe.Text = str(int(round(self.mepraum.Druckstufe.soll)))
        self.max_auswertung_druckstufe.Text = str(int(round(self.mepraum.Druckstufe.soll)))
        self.tnacht_auswertung_druckstufe.Text = str(int(round(self.mepraum.Druckstufe.soll)))

        for uber in self.mepraum.list_ueber["Ein"]:
            uber_in += float(str(uber.menge).replace(',', '.'))
        for uber in self.mepraum.list_ueber["Aus"]:
            uber_aus += float(str(uber.menge).replace(',', '.'))
        for auslass in self.mepraum.list_auslass:
            if auslass.art == '24h':
                ab24h += float(str(auslass.Luftmengenmin).replace(',', '.'))

            elif auslass.art == 'LAB':
                lab_min += float(str(auslass.Luftmengenmin).replace(',', '.'))
                lab_nb_ab += float(str(auslass.Luftmengennacht).replace(',', '.'))
                lab_max += float(str(auslass.Luftmengenmax).replace(',', '.'))
                lab_tnb_ab += float(str(auslass.Luftmengentnacht).replace(',', '.'))

            elif auslass.art == 'RZU':
                min_zu_raum += float(str(auslass.Luftmengenmin).replace(',', '.'))
                nb_zu_raum += float(str(auslass.Luftmengennacht).replace(',', '.'))
                max_zu_raum += float(str(auslass.Luftmengenmax).replace(',', '.'))
                tnb_zu_raum += float(str(auslass.Luftmengentnacht).replace(',', '.'))
            elif auslass.art == 'RAB':
                min_ab_raum += float(str(auslass.Luftmengenmin).replace(',', '.'))
                nb_ab_raum += float(str(auslass.Luftmengennacht).replace(',', '.'))
                max_ab_raum += float(str(auslass.Luftmengenmax).replace(',', '.'))
                tnb_ab_raum += float(str(auslass.Luftmengentnacht).replace(',', '.'))
            elif auslass.art == 'TZU':
                min_zu_tier += float(str(auslass.Luftmengenmin).replace(',', '.'))
                nb_zu_tier += float(str(auslass.Luftmengennacht).replace(',', '.'))
                max_zu_tier += float(str(auslass.Luftmengenmax).replace(',', '.'))
                tnb_zu_tier += float(str(auslass.Luftmengentnacht).replace(',', '.'))
            elif auslass.art == 'TAB':
                min_ab_tier += float(str(auslass.Luftmengenmin).replace(',', '.'))
                nb_ab_tier += float(str(auslass.Luftmengennacht).replace(',', '.'))
                max_ab_tier += float(str(auslass.Luftmengenmax).replace(',', '.'))
                tnb_ab_tier += float(str(auslass.Luftmengentnacht).replace(',', '.'))

        min_zu_sum = min_zu_raum + min_zu_tier + uber_in + uber_in_m
        min_ab_sum = min_ab_raum + min_ab_tier + uber_aus + lab_min + ab24h + uber_aus_m
        min_abweichung = min_zu_sum - min_ab_sum

        self.b_zu_min_sum_sys.Text = str(int(round(min_zu_sum)))
        self.b_zu_min_raum_sys.Text = str(int(round(min_zu_raum)))
        self.b_zu_min_tier_sys.Text = str(int(round(min_zu_tier)))
        self.b_zu_min_ueber_sys.Text = str(int(round(uber_in)))
        self.b_zu_min_ueber_m_sys.Text = str(int(round(uber_in_m)))

        self.b_ab_min_sum_sys.Text = str(int(round(min_ab_sum)))
        self.b_ab_min_raum_sys.Text = str(int(round(min_ab_raum)))
        self.b_ab_min_tier_sys.Text = str(int(round(min_ab_tier)))
        self.b_ab_min_ueber_sys.Text = str(int(round(uber_aus)))
        self.b_ab_min_ueber_m_sys.Text = str(int(round(uber_aus_m)))
        self.b_ab_min_lab_sys.Text = str(int(round(lab_min)))
        self.b_ab_24h_min_sys.Text = str(int(round(ab24h)))

        self.min_Abweichung_sys.Text = str(int(round(min_abweichung)))
        if abs(round(min_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.min_auswertung_sys.Text = 'OK'
            self.min_Abweichung_sys.Foreground = self.black
            self.min_auswertung_sys.Foreground = self.black
        else:
            self.min_auswertung_sys.Text = 'Passt Nicht'
            self.min_auswertung_sys.Foreground = self.red
            self.min_Abweichung_sys.Foreground = self.red

        max_zu_sum = max_zu_raum + max_zu_tier + uber_in + uber_in_m
        max_ab_sum = max_ab_raum + max_ab_tier + uber_aus + lab_max + ab24h + uber_aus_m
        max_abweichung = max_zu_sum - max_ab_sum

        self.b_zu_max_sum_sys.Text = str(int(round(max_zu_sum)))
        self.b_zu_max_raum_sys.Text = str(int(round(max_zu_raum)))
        self.b_zu_max_tier_sys.Text = str(int(round(max_zu_tier)))
        self.b_zu_max_ueber_sys.Text = str(int(round(uber_in)))
        self.b_zu_max_ueber_m_sys.Text = str(int(round(uber_in_m)))

        self.b_ab_max_sum_sys.Text = str(int(round(max_ab_sum)))
        self.b_ab_max_raum_sys.Text = str(int(round(max_ab_raum)))
        self.b_ab_max_tier_sys.Text = str(int(round(max_ab_tier)))
        self.b_ab_max_ueber_sys.Text = str(int(round(uber_aus)))
        self.b_ab_max_ueber_m_sys.Text = str(int(round(uber_aus_m)))
        self.b_ab_max_lab_sys.Text = str(int(round(lab_max)))
        self.b_ab_24h_max_sys.Text = str(int(round(ab24h)))

        self.max_Abweichung_sys.Text = str(int(round(max_abweichung)))
        if abs(round(max_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.max_auswertung_sys.Text = 'OK'
            self.max_Abweichung_sys.Foreground = self.black
            self.max_auswertung_sys.Foreground = self.black
        else:
            self.max_auswertung_sys.Text = 'Passt Nicht'
            self.max_auswertung_sys.Foreground = self.red
            self.max_Abweichung_sys.Foreground = self.red


        nb_zu_sum = nb_zu_raum + uber_in + uber_in_m
        nb_ab_sum = nb_ab_raum + uber_aus + lab_nb_ab + ab24h + uber_aus_m
        nb_abweichung = nb_zu_sum - nb_ab_sum

        self.b_zu_nacht_sum_sys.Text = str(int(round(nb_zu_sum)))
        self.b_zu_nacht_raum_sys.Text = str(int(round(nb_zu_raum)))
        self.b_zu_nacht_ueber_sys.Text = str(int(round(uber_in)))
        self.b_zu_nacht_ueber_m_sys.Text = str(int(round(uber_in_m)))

        self.b_ab_nacht_sum_sys.Text = str(int(round(nb_ab_sum)))
        self.b_ab_nacht_raum_sys.Text = str(int(round(nb_ab_raum)))
        self.b_ab_nacht_ueber_sys.Text = str(int(round(uber_aus)))
        self.b_ab_nacht_ueber_m_sys.Text = str(int(round(uber_aus_m)))
        self.b_ab_nacht_lab_sys.Text = str(int(round(lab_nb_ab)))
        self.b_ab_24h_nacht_sys.Text = str(int(round(ab24h)))

        self.nacht_Abweichung_sys.Text = str(int(round(nb_abweichung)))
        if abs(round(nb_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.nacht_auswertung_sys.Text = 'OK'
            self.nacht_Abweichung_sys.Foreground = self.black
            self.nacht_auswertung_sys.Foreground = self.black
        else:
            self.nacht_auswertung_sys.Text = 'Passt Nicht'
            self.nacht_auswertung_sys.Foreground = self.red
            self.nacht_Abweichung_sys.Foreground = self.red
        
        tnb_zu_sum = tnb_zu_raum + uber_in + uber_in_m
        tnb_ab_sum = tnb_ab_raum + uber_aus + lab_tnb_ab + ab24h + uber_aus_m
        tnb_abweichung = tnb_zu_sum - tnb_ab_sum

        self.b_zu_tnacht_sum_sys.Text = str(int(round(tnb_zu_sum)))
        self.b_zu_tnacht_raum_sys.Text = str(int(round(tnb_zu_raum)))
        self.b_zu_tnacht_ueber_sys.Text = str(int(round(uber_in)))
        self.b_zu_tnacht_ueber_m_sys.Text = str(int(round(uber_in_m)))

        self.b_ab_tnacht_sum_sys.Text = str(int(round(tnb_ab_sum)))
        self.b_ab_tnacht_raum_sys.Text = str(int(round(tnb_ab_raum)))
        self.b_ab_tnacht_ueber_sys.Text = str(int(round(uber_aus)))
        self.b_ab_tnacht_ueber_m_sys.Text = str(int(round(uber_aus_m)))
        self.b_ab_tnacht_lab_sys.Text = str(int(round(lab_tnb_ab)))
        self.b_ab_24h_tnacht_sys.Text = str(int(round(ab24h)))

        self.nacht_Abweichung_sys.Text = str(int(round(tnb_abweichung)))
        if abs(round(tnb_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.tnacht_auswertung_sys.Text = 'OK'
            self.tnacht_Abweichung_sys.Foreground = self.black
            self.tnacht_auswertung_sys.Foreground = self.black
        else:
            self.tnacht_auswertung_sys.Text = 'Passt Nicht'
            self.tnacht_auswertung_sys.Foreground = self.red
            self.tnacht_Abweichung_sys.Foreground = self.red




        self.mepraum.zu_min.ist = int(round(min_zu_raum))
        self.mepraum.zu_max.ist = int(round(max_zu_raum))
        self.mepraum.ab_min.ist = int(round(min_ab_raum))
        self.mepraum.ab_max.ist = int(round(max_ab_raum))
        self.mepraum.ab_24h.ist = int(round(ab24h))
        self.mepraum.ab_lab_min.ist = int(round(lab_min))
        self.mepraum.ab_lab_max.ist = int(round(lab_max))
        self.mepraum.nb_zu.ist = int(round(nb_zu_raum))
        self.mepraum.nb_ab.ist = int(round(nb_ab_raum))
        self.mepraum.tnb_zu.ist = int(round(tnb_zu_raum))
        self.mepraum.tnb_ab.ist = int(round(tnb_ab_raum))

        if self.mepraum.IsTierRaum != 0:
            self.mepraum.tier_zu_min.ist = int(round(min_zu_tier))
            self.mepraum.tier_ab_min.ist = int(round(min_ab_tier))
            self.mepraum.tier_zu_max.ist = int(round(max_zu_tier))
            self.mepraum.tier_ab_max.ist = int(round(max_ab_tier))
        self.mepraum.ueber_in.ist = int(round(uber_in))
        self.mepraum.ueber_aus.ist = int(round(uber_aus))
        self.mepraum.ueber_sum.ist = int(round(uber_in-uber_aus))

    def Auswertung_MEP(self):
        if self.mepraum.IsTierRaum != 0:
            min_zu_sum = float(self.mepraum.zu_min.soll) + float(self.mepraum.tier_zu_min.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
            min_ab_sum = float(self.mepraum.ab_min.soll) + float(self.mepraum.ab_lab_min.soll) + float(self.mepraum.tier_ab_min.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll)
            max_zu_sum = float(self.mepraum.zu_max.soll) + float(self.mepraum.tier_zu_max.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
            max_ab_sum = float(self.mepraum.ab_max.soll) + float(self.mepraum.ab_lab_max.soll) + float(self.mepraum.tier_ab_max.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll)
        else:
            min_zu_sum = float(self.mepraum.zu_min.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
            min_ab_sum = float(self.mepraum.ab_min.soll) + float(self.mepraum.ueber_aus.soll)+ float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll) + float(self.mepraum.ab_lab_min.soll)
            max_zu_sum = float(self.mepraum.zu_max.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
            max_ab_sum = float(self.mepraum.ab_max.soll) + float(self.mepraum.ueber_aus.soll)+ float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll) + float(self.mepraum.ab_lab_max.soll)

        min_abweichung = min_zu_sum - min_ab_sum
        max_abweichung = max_zu_sum - max_ab_sum


        self.b_zu_min_sum_mep.Text = str(int(round(min_zu_sum)))
        self.b_zu_min_raum_mep.Text = str(int(round(float(self.mepraum.zu_min.soll))))
        try:self.b_zu_min_tier_mep.Text = str(int(round(float(self.mepraum.tier_zu_min.soll))))
        except:self.b_zu_min_tier_mep.Text = '0'
        self.b_zu_min_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_in.soll))))
        self.b_zu_min_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_in_manuell.soll))))

        self.b_ab_min_sum_mep.Text = str(int(round(min_ab_sum)))
        self.b_ab_min_raum_mep.Text = str(int(round(float(self.mepraum.ab_min.soll))))
        try: self.b_ab_min_tier_mep.Text = str(int(round(float(self.mepraum.tier_ab_min.soll))))
        except:self.b_ab_min_tier_mep.Text = '0'
        self.b_ab_min_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_aus.soll))))
        self.b_ab_min_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_aus_manuell.soll))))
        self.b_ab_min_lab_mep.Text = str(int(round(float(self.mepraum.ab_lab_min.soll))))
        self.b_ab_24h_min_mep.Text = str(int(round(float(self.mepraum.ab_24h.soll))))

        self.min_Abweichung_mep.Text = str(int(round(min_abweichung)))
        if abs(round(min_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.min_auswertung_mep.Text = 'OK'
            self.min_Abweichung_mep.Foreground = self.black
            self.min_auswertung_mep.Foreground = self.black
        else:
            self.min_auswertung_mep.Text = 'Passt Nicht'
            self.min_auswertung_mep.Foreground = self.red
            self.min_Abweichung_mep.Foreground = self.red
        
        self.b_zu_max_sum_mep.Text = str(int(round(max_zu_sum)))
        self.b_zu_max_raum_mep.Text = str(int(round(float(self.mepraum.zu_max.soll))))
        try:self.b_zu_max_tier_mep.Text = str(int(round(float(self.mepraum.tier_zu_max.soll))))
        except:self.b_zu_max_tier_mep.Text = '0'
        self.b_zu_max_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_in.soll))))
        self.b_zu_max_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_in_manuell.soll))))

        self.b_ab_max_sum_mep.Text = str(int(round(max_ab_sum)))
        self.b_ab_max_raum_mep.Text = str(int(round(float(self.mepraum.ab_max.soll))))
        try: self.b_ab_max_tier_mep.Text = str(int(round(float(self.mepraum.tier_ab_max.soll))))
        except:self.b_ab_max_tier_mep.Text = '0'
        self.b_ab_max_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_aus.soll))))
        self.b_ab_max_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_aus_manuell.soll))))
        self.b_ab_max_lab_mep.Text = str(int(round(float(self.mepraum.ab_lab_max.soll))))
        self.b_ab_24h_max_mep.Text = str(int(round(float(self.mepraum.ab_24h.soll))))

        self.max_Abweichung_mep.Text = str(int(round(max_abweichung)))
        if abs(round(max_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.max_auswertung_mep.Text = 'OK'
            self.max_Abweichung_mep.Foreground = self.black
            self.max_auswertung_mep.Foreground = self.black
        else:
            self.max_auswertung_mep.Text = 'Passt Nicht'
            self.max_auswertung_mep.Foreground = self.red
            self.max_Abweichung_mep.Foreground = self.red

        nb_zu_sum = float(self.mepraum.nb_zu.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
        nb_ab_sum = float(self.mepraum.nb_ab.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll) + float(self.mepraum.ab_lab_min.soll)
        nb_abweichung = nb_zu_sum - nb_ab_sum
        self.b_zu_nacht_sum_mep.Text = str(int(round(nb_zu_sum)))
        self.b_zu_nacht_raum_mep.Text = str(int(round(float(self.mepraum.nb_zu.soll))))
        self.b_zu_nacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_in.soll))))
        self.b_zu_nacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_in_manuell.soll))))

        self.b_ab_nacht_sum_mep.Text = str(int(round(nb_ab_sum)))
        self.b_ab_nacht_raum_mep.Text = str(int(round(float(self.mepraum.nb_ab.soll))))
        self.b_ab_nacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_aus.soll))))
        self.b_ab_nacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_aus_manuell.soll))))
        self.b_ab_nacht_lab_mep.Text = str(int(round(float(self.mepraum.ab_lab_min.soll))))
        self.b_ab_24h_nacht_mep.Text = str(int(round(float(self.mepraum.ab_24h.soll))))

        self.nacht_Abweichung_mep.Text = str(int(round(nb_abweichung)))

        if abs(round(nb_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.nacht_auswertung_mep.Text = 'OK'
            self.nacht_Abweichung_mep.Foreground = self.black
            self.nacht_auswertung_mep.Foreground = self.black
        else:
            self.nacht_auswertung_mep.Text = 'Passt Nicht'
            self.nacht_auswertung_mep.Foreground = self.red
            self.nacht_Abweichung_mep.Foreground = self.red
        
        tnb_zu_sum = float(self.mepraum.tnb_zu.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
        tnb_ab_sum = float(self.mepraum.tnb_ab.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll) + float(self.mepraum.ab_lab_min.soll)
        tnb_abweichung = tnb_zu_sum - tnb_ab_sum
        self.b_zu_tnacht_sum_mep.Text = str(int(round(tnb_zu_sum)))
        self.b_zu_tnacht_raum_mep.Text = str(int(round(float(self.mepraum.tnb_zu.soll))))
        self.b_zu_tnacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_in.soll))))
        self.b_zu_tnacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_in_manuell.soll))))

        self.b_ab_tnacht_sum_mep.Text = str(int(round(tnb_ab_sum)))
        self.b_ab_tnacht_raum_mep.Text = str(int(round(float(self.mepraum.tnb_ab.soll))))
        self.b_ab_tnacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_aus.soll))))
        self.b_ab_tnacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_aus_manuell.soll))))
        self.b_ab_tnacht_lab_mep.Text = str(int(round(float(self.mepraum.ab_lab_min.soll))))
        self.b_ab_24h_tnacht_mep.Text = str(int(round(float(self.mepraum.ab_24h.soll))))

        self.tnacht_Abweichung_mep.Text = str(int(round(tnb_abweichung)))

        if abs(round(tnb_abweichung-self.mepraum.Druckstufe.soll)) == 0:
            self.tnacht_auswertung_mep.Text = 'OK'
            self.tnacht_Abweichung_mep.Foreground = self.black
            self.tnacht_auswertung_mep.Foreground = self.black
        else:
            self.tnacht_auswertung_mep.Text = 'Passt Nicht'
            self.tnacht_auswertung_mep.Foreground = self.red
            self.tnacht_Abweichung_mep.Foreground = self.red

    def zeigraum(self, sender, args):
        self.resizeAnsichtEvent.Raise()

    def nachtbetriebchanged(self, sender, args):
        try:
            self.mepraum.nachtbetrieb = sender.IsChecked
        except:pass

    def tnachtbetriebchanged(self, sender, args):
        try:
            self.mepraum.nachtbetrieb = sender.IsChecked
        except:pass

    def labmineingabe_changed(self, sender, args):
        try:
            self.mepraum.ab_lab_min.soll = round(float(str(sender.Text).replace(',', '.')))
            self.Auswertung_MEP()
            self.Auswertung_System()
        except:pass
    def labmaxeingabe_changed(self, sender, args):
        try:
            self.mepraum.ab_lab_max.soll = round(float(str(sender.Text).replace(',', '.')))
            self.Auswertung_MEP()
            self.Auswertung_System()
        except:pass

    def ab24heingabe_changed(self, sender, args):
        try:
            self.mepraum.ab_24h.soll = round(float(str(sender.Text).replace(',', '.')))
            self.Auswertung_MEP()
            self.Auswertung_System()
        except:pass
    def druckeingabe_changed(self, sender, args):
        try:
            self.mepraum.Druckstufe.soll = round(float(str(sender.Text).replace(',', '.')))
            self.Auswertung_MEP()
            self.Auswertung_System()
        except:pass

    def LW_nacht_changed(self, sender, args):
        try:
            self.mepraum.NB_LW = float(str(sender.Text).replace(',', '.'))
        except:pass

    def LW_tnacht_changed(self, sender, args):
        try:
            self.mepraum.T_NB_LW = float(str(sender.Text).replace(',', '.'))
        except:pass

    def von_nacht_changed(self, sender, args):
        try:
            self.mepraum.nb_von.soll = float(str(sender.Text).replace(',', '.'))
        except:pass

    def bis_nacht_changed(self, sender, args):
        try:
            self.mepraum.nb_bis.soll = float(str(sender.Text).replace(',', '.'))
        except:pass

    def von_tnacht_changed(self, sender, args):
        try:
            self.mepraum.tnb_von.soll = float(str(sender.Text).replace(',', '.'))
        except:pass

    def bis_tnacht_changed(self, sender, args):
        try:
            self.mepraum.tnb_bis.soll = float(str(sender.Text).replace(',', '.'))
        except:pass

    def tiefernachtbetriebchanged(self, sender, args):
        return
        self.resizeAnsichtEvent.Raise()

    def LW_tiefernacht_changed(self, sender, args):
        return
        self.resizeAnsichtEvent.Raise()

    def faktorchanged(self, sender, args):
        try:
            self.mepraum.faktor = sender.Text
        except:pass

    def bezugsnameselectionchanged(self, sender, args):
        self.mepraum.bezugsname = self.bezugsname.SelectedItem.ToString()
        for el in self.mepraum.berechnung_nach.keys():
            if self.mepraum.berechnung_nach[el] == self.bezugsname.SelectedItem.ToString():
                self.mepraum.bezugsnummer = el
        self.einheit.Text = self.mepraum.einheit = self.mepraum.einheit_liste[self.mepraum.bezugsnummer]

    def MEP_changed(self, sender, args):
        text = str(self.Raumnr.SelectedItem)
        self.rauman.IsEnabled = True
        self.temp_luftauslass.Clear()
        try:
            self.mepraum = self.mepraum_liste[text]

            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden

            for el in self.lv_vsr.Items:
                el.checked = False
            try:self.set_up()
            except Exception as e:print(e)
        except:pass

    def MEP_Suche_changed(self, sender, args):
        try:
            temp = []
            text = self.suche.Text
            if text in [None,'']:
                self.suche.Text = ''
                self.Raumnr.ItemsSource = sorted(self.mepraum_liste.keys())
            else:
                for el in sorted(self.mepraum_liste.keys()):
                    if el.upper().find(text.upper()) != -1:
                        temp.append(el)
                self.Raumnr.ItemsSource = temp
        except:pass

    def VSR_MEPRaum_changed(self, sender, args):
        text = str(self.auswahl.SelectedItem)
        text_original = str(self.Raumnr.SelectedItem)
        if text != text_original:
            self.rauman.IsEnabled = False
            self.Rauminfo.Visibility = self.visible
            self.Rauminfo.Text = "Die angezeigte Rauminfomationen beziehen sich auf Raum => " + text
        else:
            self.rauman.IsEnabled = True
            self.Rauminfo.Visibility = self.hidden
        try:
            temp = None
            for el in self.lv_vsr.Items:
                if el.checked:
                    temp = el
                    break
            self.mepraum = self.mepraum_liste[text]
            self.set_up()
            self.temp_luftauslass.Clear()
            for el in temp.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass

        except:pass

    def VSR_checked_changed(self, sender, args):
        Checked = sender.IsChecked

        item = sender.DataContext

        self.temp_luftauslass.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            item.checked = True

            self.lv_vsr.Items.Refresh()

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass

            if len(item.liste_Raum) > 1:
                self.raumwechsel_hinweis.Visibility= self.visible
                self.warnung.Visibility= self.visible
                self.auswahl.Visibility= self.visible
                self.auswahl.ItemsSource = sorted(item.liste_Raum)
            else:
                self.raumwechsel_hinweis.Visibility= self.hidden
                self.warnung.Visibility= self.hidden
                self.auswahl.Visibility= self.hidden

        else:
            self.raumwechsel_hinweis.Visibility= self.hidden
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
            self.lv_auslass.ItemsSource = self.mepraum.list_auslass

    def Setkey(self, sender, args):
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        elif args.Key == self.Key.OemComma or args.Key == self.Key.Decimal:
            if sender.Text.Length == 0:
                args.Handled = True
            else:
                zahl_new = sender.Text.replace(',', '.') + '.'
                try:
                    float(zahl_new)
                    args.Handled = False
                except:
                    args.Handled = True
        else:
            args.Handled = True

    def Auslass_Volumen_changed(self, sender, args):
        try:
            auslass = sender.DataContext
            vsr = auslass.VSR_Class
            vsr.Luftmengenermitteln_new()
            self.lv_vsr.Items.Refresh()
        except:pass
        try:
            self.Auswertung_System()
            self.anlageninfo.Items.Refresh()
            self.grundinfo.Items.Refresh()
        except:pass

    def Auslass_Selected_Changed(self, sender, args):
        if self.lv_auslass.SelectedIndex != -1:
            self.highlightelement.elemId = self.lv_auslass.SelectedItem.elemid
            self.highlightelementEvent.Raise()
            self.lv_vsr.SelectedIndex = -1
            self.lv_ueber_aus.SelectedIndex = -1
            self.lv_ueber_in.SelectedIndex = -1
        else:
            return

    def VSR_Selected_Changed(self, sender, args):
        if self.lv_vsr.SelectedIndex != -1:
            self.highlightelement.elemId = self.lv_vsr.SelectedItem.elemid
            self.highlightelementEvent.Raise()
            self.lv_auslass.SelectedIndex = -1
            self.lv_ueber_aus.SelectedIndex = -1
            self.lv_ueber_in.SelectedIndex = -1
        else:
            return
    
    def Ueber_aus_Selected_Changed(self, sender, args):
        if self.lv_ueber_aus.SelectedIndex != -1:
            self.highlightelement.elemId = self.lv_ueber_aus.SelectedItem.elemid
            self.highlightelementEvent.Raise()
            self.lv_vsr.SelectedIndex = -1
            self.lv_auslass.SelectedIndex = -1
            self.lv_ueber_in.SelectedIndex = -1
        else:
            return

    def Ueber_in_Selected_Changed(self, sender, args):
        if self.lv_ueber_in.SelectedIndex != -1:
            self.highlightelement.elemId = self.lv_ueber_in.SelectedItem.elemid
            self.highlightelementEvent.Raise()
            self.lv_vsr.SelectedIndex = -1
            self.lv_auslass.SelectedIndex = -1
            self.lv_ueber_aus.SelectedIndex = -1
        else:
            return

    def anzeigen_VSR(self, sender, args):
        if self.lv_vsr.SelectedItem:
            self.showelement.elemId = self.lv_vsr.SelectedItem.elemid
            self.showelement.Raum_Einbau = self.lv_vsr.SelectedItem.raum
            self.showelement.Raum = self.mepraum.Raumnr
            self.showelementEvent.Raise()

        else:
            return

    def anzeigen_auslass(self, sender, args):
        if self.lv_auslass.SelectedItem:
            self.showelement.elemId = self.lv_auslass.SelectedItem.elemid
            self.showelement.Raum_Einbau = self.lv_auslass.SelectedItem.raum
            self.showelement.Raum = self.mepraum.Raumnr
            self.showelementEvent.Raise()

        else:
            return

    def anzeigen_ueberaus(self, sender, args):
        if self.lv_ueber_aus.SelectedItem:
            self.showelement.elemId = self.lv_ueber_aus.SelectedItem.elemid
            self.showelement.Raum_Einbau = self.lv_ueber_aus.SelectedItem.raum
            self.showelement.Raum = self.mepraum.Raumnr
            self.showelementEvent.Raise()

        else:
            return

    def anzeigen_ueberin(self, sender, args):
        if self.lv_ueber_in.SelectedItem:
            self.showelement.elemId = self.lv_ueber_in.SelectedItem.elemid
            self.showelement.Raum_Einbau = self.lv_ueber_in.SelectedItem.raum
            self.showelement.Raum = self.mepraum.Raumnr
            self.showelementEvent.Raise()

        else:
            return

    def uebertragen_VSR(self, sender, args):
        self.changeparameterwert.bauteile = self.lv_vsr.Items
        if self.changeparameterwert.bauteile.Count > 0:
            self.changeparameterwertEvent.Raise()
        else:
            return

    def uebertragen_auslass(self, sender, args):
        self.changeparameterwert.bauteile = self.lv_auslass.Items
        if self.changeparameterwert.bauteile.Count > 0:
            self.changeparameterwertEvent.Raise()
        else:
            return
        self.rueckspielen.IsEnabled = True

    # def resetview(self):
    #     if self.raumanzeigen:
    #         self.resetAnsicht.boxmax = self.originalboxmax
    #         self.resetAnsicht.boxmin = self.originalboxmin
    #         self.raumanzeigen = False
    #         self.resetAnsichtEvent.Raise()

    def Berechnen(self, sender, args):
        try:
            self.BerechnenEvent.Raise()
        except:pass

    def update(self, sender, args):
        try:
            # self.mepraum.update()
            self.mepraum.TiefeNachtbetrieb_Berechnen()
        except:pass
        try:
            self.mepraum.Nachtbetrieb_Berechnen()
        except:pass
        try:
            self.mepraum.Tagesbetrieb_Berechnen()
        except:pass
        for el in self.mepraum.list_auslass:
            el.set_up()

        for el in self.mepraum.list_vsr:
            el.set_up()
            el.Luftmengenermitteln_new()

        self.lv_auslass.Items.Refresh()
        self.lv_vsr.Items.Refresh()
        try:
            self.Auswertung_System()
            self.Auswertung_MEP()
            self.anlageninfo.Items.Refresh()
            self.grundinfo.Items.Refresh()
            self.schachtinfo.Items.Refresh()
        except Exception as e:print(e)

    def updaterevit(self, sender, args):
        try:
            self.updateModelEvent.Raise()
        except:pass

    def undo(self, sender, args):
        try:
            self.undoEvent.Raise()
        except:pass
    def doverteilen_mep(self, sender, args):
        try:
            self.verteilenmepEvent.Raise()
        except Exception as e:print(e)
    def doverteilen(self, sender, args):
        try:
            self.verteilenEvent.Raise()
        except Exception as e:print(e)

    def abbrechen(self, sender, args):
        if self.Trans == 0:
            self.Close()
            return
        for el in range(0,self.Trans):
            try:self.undoEvent.Raise()
            except:pass
        self.Close()

mepWPF = MEPRaum_Uebersicht()
mepWPF.changeparameterwert.GUI_MEPRaum = mepWPF
mepWPF._undo.GUI_MEPRaum = mepWPF
mepWPF.updateModel.GUI_MEPRaum = mepWPF
mepWPF.resizeAnsicht.GUI_MEPRaum = mepWPF
mepWPF.verteilen.GUI_MEPRaum = mepWPF
mepWPF.verteilenmep.GUI_MEPRaum = mepWPF
mepWPF.berechnen.GUI_MEPRaum = mepWPF
try:
    mepWPF.Show()
except Exception as e:
    logger.error(e)
    mepWPF.Close()
    script.exit()
