# coding: utf8
from pyrevit import script, forms
from IGF_log import getlog
from System.Collections.ObjectModel import ObservableCollection
# from eventhandler import *
from System.Text.RegularExpressions import Regex
from eventhandler import Luftauslass,VSR,DICT_MEP_ITEMSSOIRCE,VISIBLE,HIDDEN,RED,BLACK,ExtenalEventListe,ExternalEvent,LISTE_MEP

__title__ = "Raumluftbilanz_MFC"
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

class MEPRaum_Uebersicht(forms.WPFWindow):
    def __init__(self):
        self.red = RED
        self.black = BLACK
        self.regex1 = Regex("[^0-9,]+")
        self.regex2 = Regex("[^0-9]+")
        self.regex3 = Regex("[^0-9-]+")
        self.externaleventliste = ExtenalEventListe()
        self.ListeMEP = LISTE_MEP
        self.path = ''
        self.externalevent = ExternalEvent.Create(self.externaleventliste)
        self.visible = VISIBLE
        self.hidden = HIDDEN

        forms.WPFWindow.__init__(self,'MEP.xaml',handle_esc=False)
        self.mepraum_liste = DICT_MEP_ITEMSSOIRCE

        self.Raumnr.ItemsSource = sorted(self.mepraum_liste.keys())

        self.mepraum = None
        self.temp_luftauslass = ObservableCollection[Luftauslass]()
        self.temp_vsr = ObservableCollection[VSR]()
        self.temp_Iris = ObservableCollection[VSR]()
        self.warnung.Visibility= self.hidden
    
    def TextBox_GotFocus(self,sender,e):
        tb = sender
        tb.SelectAll()
    
    def vsr_selectedfabrikat_changed(self,sender,e):
        if self.lv_vsr1_getrennt.SelectedItem:
            try:
                item = sender.DataContext
                item.vsrauswaelten()
                item.vsrueberpruefen()
            except Exception as e:
                pass
    
    def vsr_selectedtyp_changed(self,sender,e):
        if self.lv_vsr1_getrennt.SelectedItem:
            try:
                item = sender.DataContext
                item.VSR_Hersteller = item.DICT_DatenBank[item.Liste_Herstellertyp[item.Herstellertypindex]]
                item.vsrueberpruefen()
            except Exception as e:
                pass

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


        self.isnachtbetrieb.IsChecked = True if self.mepraum.nachtbetrieb  else False
        self.istiefenachtbetrieb.IsChecked = True if self.mepraum.tiefenachtbetrieb  else False
        self.LW_nacht.Text = str(self.mepraum.NB_LW)
        self.von_nacht.Text = str(self.mepraum.nb_von.soll)
        self.bis_nacht.Text = str(self.mepraum.nb_bis.soll)

        self.LW_tnacht.Text = str(self.mepraum.T_NB_LW)
        self.von_tnacht.Text = str(self.mepraum.tnb_von.soll)
        self.bis_tnacht.Text = str(self.mepraum.tnb_bis.soll)

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
        self.detail_min.ItemsSource = self.mepraum.Detail_Min
        self.detail_min.Items.Refresh()

        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr1.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
        self.lv_vsr2.ItemsSource = self.mepraum.list_vsr2


        self.lv_auslass.ItemsSource = self.mepraum.list_auslass
        self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass

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

        min_zu_sum = min_zu_raum + uber_in + uber_in_m
        min_ab_sum = min_ab_raum + uber_aus + lab_min + ab24h + uber_aus_m
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
        if abs(round(min_abweichung)) <= 3:
        #if abs(round(min_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
            self.min_auswertung_sys.Text = 'OK'
            self.min_Abweichung_sys.Foreground = self.black
            self.min_auswertung_sys.Foreground = self.black
        else:
            self.min_auswertung_sys.Text = 'Passt Nicht'
            self.min_auswertung_sys.Foreground = self.red
            self.min_Abweichung_sys.Foreground = self.red

        max_zu_sum = max_zu_raum + uber_in + uber_in_m
        max_ab_sum = max_ab_raum + uber_aus + lab_max + ab24h + uber_aus_m
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
        if abs(round(max_abweichung)) <= 3:
        #if abs(round(max_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
            self.max_auswertung_sys.Text = 'OK'
            self.max_Abweichung_sys.Foreground = self.black
            self.max_auswertung_sys.Foreground = self.black
        else:
            self.max_auswertung_sys.Text = 'Passt Nicht'
            self.max_auswertung_sys.Foreground = self.red
            self.max_Abweichung_sys.Foreground = self.red

        if self.mepraum.nachtbetrieb:
            nb_zu_sum = nb_zu_raum + uber_in + uber_in_m
            nb_ab_sum = nb_ab_raum + uber_aus + lab_nb_ab + ab24h + uber_aus_m
            nb_abweichung = nb_zu_sum - nb_ab_sum
        else:
            nb_zu_sum = nb_zu_raum
            nb_ab_sum = nb_ab_raum + ab24h 
            nb_abweichung = nb_zu_sum - nb_ab_sum

        self.b_zu_nacht_sum_sys.Text = str(int(round(nb_zu_sum)))
        self.b_zu_nacht_raum_sys.Text = str(int(round(nb_zu_raum)))
        if self.mepraum.nachtbetrieb:
            self.b_zu_nacht_ueber_sys.Text = str(int(round(uber_in)))
            self.b_zu_nacht_ueber_m_sys.Text = str(int(round(uber_in_m)))
            self.b_ab_nacht_ueber_sys.Text = str(int(round(uber_aus)))
            self.b_ab_nacht_ueber_m_sys.Text = str(int(round(uber_aus_m)))
            self.b_ab_nacht_lab_sys.Text = str(int(round(lab_nb_ab)))
        else:
            self.b_zu_nacht_ueber_sys.Text = '0'
            self.b_zu_nacht_ueber_m_sys.Text = '0'
            self.b_ab_nacht_ueber_sys.Text = '0'
            self.b_ab_nacht_ueber_m_sys.Text = '0'
            self.b_ab_nacht_lab_sys.Text = '0'


        self.b_ab_nacht_sum_sys.Text = str(int(round(nb_ab_sum)))
        self.b_ab_nacht_raum_sys.Text = str(int(round(nb_ab_raum)))
        self.b_ab_24h_nacht_sys.Text = str(int(round(ab24h)))

        self.nacht_Abweichung_sys.Text = str(int(round(nb_abweichung)))
        if abs(round(nb_abweichung)) <= 3:
        #if abs(round(nb_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
            self.nacht_auswertung_sys.Text = 'OK'
            self.nacht_Abweichung_sys.Foreground = self.black
            self.nacht_auswertung_sys.Foreground = self.black
        else:
            self.nacht_auswertung_sys.Text = 'Passt Nicht'
            self.nacht_auswertung_sys.Foreground = self.red
            self.nacht_Abweichung_sys.Foreground = self.red
        if self.mepraum.tiefenachtbetrieb:
        
            tnb_zu_sum = tnb_zu_raum + uber_in + uber_in_m
            tnb_ab_sum = tnb_ab_raum + uber_aus + lab_tnb_ab + ab24h + uber_aus_m
            tnb_abweichung = tnb_zu_sum - tnb_ab_sum
            self.b_zu_tnacht_ueber_sys.Text = str(int(round(uber_in)))
            self.b_zu_tnacht_ueber_m_sys.Text = str(int(round(uber_in_m)))
            self.b_ab_tnacht_ueber_sys.Text = str(int(round(uber_aus)))
            self.b_ab_tnacht_ueber_m_sys.Text = str(int(round(uber_aus_m)))
            self.b_ab_tnacht_lab_sys.Text = str(int(round(lab_tnb_ab)))
            self.b_ab_24h_tnacht_sys.Text = str(int(round(ab24h)))
            #if abs(round(tnb_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
            if abs(round(tnb_abweichung)) <= 3:
                self.tnacht_auswertung_sys.Text = 'OK'
                self.tnacht_Abweichung_sys.Foreground = self.black
                self.tnacht_auswertung_sys.Foreground = self.black
            else:
                self.tnacht_auswertung_sys.Text = 'Passt Nicht'
                self.tnacht_auswertung_sys.Foreground = self.red
                self.tnacht_Abweichung_sys.Foreground = self.red
        else:
            tnb_zu_sum = tnb_zu_raum 
            tnb_ab_sum = tnb_ab_raum 
            tnb_abweichung = tnb_zu_sum - tnb_ab_sum
            self.b_zu_tnacht_ueber_sys.Text = '0'
            self.b_zu_tnacht_ueber_m_sys.Text = '0'
            self.b_ab_tnacht_ueber_sys.Text = '0'
            self.b_ab_tnacht_ueber_m_sys.Text = '0'
            self.b_ab_tnacht_lab_sys.Text = '0'
            self.b_ab_24h_tnacht_sys.Text = '0'
            
            self.tnacht_auswertung_sys.Text = 'OK'
            self.tnacht_Abweichung_sys.Foreground = self.black
            self.tnacht_auswertung_sys.Foreground = self.black
            


        self.b_zu_tnacht_sum_sys.Text = str(int(round(tnb_zu_sum)))
        self.b_zu_tnacht_raum_sys.Text = str(int(round(tnb_zu_raum)))
        self.b_ab_tnacht_sum_sys.Text = str(int(round(tnb_ab_sum)))
        self.b_ab_tnacht_raum_sys.Text = str(int(round(tnb_ab_raum)))
        self.tnacht_Abweichung_sys.Text = str(int(round(tnb_abweichung)))
        

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
        # if self.mepraum.IsTierRaum != 0:
        #     min_zu_sum = float(self.mepraum.zu_min.soll) + float(self.mepraum.tier_zu_min.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
        #     min_ab_sum = float(self.mepraum.ab_min.soll) + float(self.mepraum.ab_lab_min.soll) + float(self.mepraum.tier_ab_min.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll)
        #     max_zu_sum = float(self.mepraum.zu_max.soll) + float(self.mepraum.tier_zu_max.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
        #     max_ab_sum = float(self.mepraum.ab_max.soll) + float(self.mepraum.ab_lab_max.soll) + float(self.mepraum.tier_ab_max.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll)
        # else:
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
        #if abs(round(min_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
        if abs(round(min_abweichung)) <= 3:
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
        #if abs(round(max_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
        if abs(round(max_abweichung)) <= 3:
            self.max_auswertung_mep.Text = 'OK'
            self.max_Abweichung_mep.Foreground = self.black
            self.max_auswertung_mep.Foreground = self.black
        else:
            self.max_auswertung_mep.Text = 'Passt Nicht'
            self.max_auswertung_mep.Foreground = self.red
            self.max_Abweichung_mep.Foreground = self.red
        
        if not self.mepraum.nachtbetrieb:
            nb_zu_sum = float(self.mepraum.nb_zu.soll) 
            nb_ab_sum = float(self.mepraum.nb_ab.soll) + float(self.mepraum.ab_24h.soll) 
        else:
            nb_zu_sum = float(self.mepraum.nb_zu.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
            nb_ab_sum = float(self.mepraum.nb_ab.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll) + float(self.mepraum.ab_lab_min.soll)
        nb_abweichung = nb_zu_sum - nb_ab_sum
        self.b_zu_nacht_sum_mep.Text = str(int(round(nb_zu_sum)))
        self.b_zu_nacht_raum_mep.Text = str(int(round(float(self.mepraum.nb_zu.soll))))
        if self.mepraum.nachtbetrieb:
            self.b_zu_nacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_in.soll))))
        else:self.b_zu_nacht_ueber_mep.Text = '0'
        if self.mepraum.nachtbetrieb:self.b_zu_nacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_in_manuell.soll))))
        else:self.b_zu_nacht_ueber_m_mep.Text = '0'

        self.b_ab_nacht_sum_mep.Text = str(int(round(nb_ab_sum)))
        self.b_ab_nacht_raum_mep.Text = str(int(round(float(self.mepraum.nb_ab.soll))))
        if self.mepraum.nachtbetrieb:self.b_ab_nacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_aus.soll))))
        else:self.b_ab_nacht_ueber_mep.Text = '0'
        if self.mepraum.nachtbetrieb:self.b_ab_nacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_aus_manuell.soll))))
        else:self.b_ab_nacht_ueber_m_mep.Text = '0'
        if self.mepraum.nachtbetrieb:self.b_ab_nacht_lab_mep.Text = str(int(round(float(self.mepraum.ab_lab_min.soll))))
        else:self.b_ab_nacht_lab_mep.Text = '0'
        self.b_ab_24h_nacht_mep.Text = str(int(round(float(self.mepraum.ab_24h.soll))))

        self.nacht_Abweichung_mep.Text = str(int(round(nb_abweichung)))

        #if abs(round(nb_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
        if abs(round(nb_abweichung)) <= 3:
            self.nacht_auswertung_mep.Text = 'OK'
            self.nacht_Abweichung_mep.Foreground = self.black
            self.nacht_auswertung_mep.Foreground = self.black
        else:
            self.nacht_auswertung_mep.Text = 'Passt Nicht'
            self.nacht_auswertung_mep.Foreground = self.red
            self.nacht_Abweichung_mep.Foreground = self.red
        if self.mepraum.tiefenachtbetrieb:
            tnb_zu_sum = float(self.mepraum.tnb_zu.soll) + float(self.mepraum.ueber_in.soll) + float(self.mepraum.ueber_in_manuell.soll)
            tnb_ab_sum = float(self.mepraum.tnb_ab.soll) + float(self.mepraum.ueber_aus.soll) + float(self.mepraum.ab_24h.soll) + float(self.mepraum.ueber_aus_manuell.soll) + float(self.mepraum.ab_lab_min.soll)
        else:
            tnb_zu_sum = float(self.mepraum.tnb_zu.soll)
            tnb_ab_sum = float(self.mepraum.tnb_ab.soll) + float(self.mepraum.ab_24h.soll) 
        
        tnb_abweichung = tnb_zu_sum - tnb_ab_sum
        self.b_zu_tnacht_sum_mep.Text = str(int(round(tnb_zu_sum)))
        self.b_zu_tnacht_raum_mep.Text = str(int(round(float(self.mepraum.tnb_zu.soll))))
        if self.mepraum.tiefenachtbetrieb:self.b_zu_tnacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_in.soll))))
        else:self.b_zu_tnacht_ueber_mep.Text = '0'
        if self.mepraum.tiefenachtbetrieb:self.b_zu_tnacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_in_manuell.soll))))
        else:self.b_zu_tnacht_ueber_m_mep.Text = '0'
        
        self.b_ab_tnacht_sum_mep.Text = str(int(round(tnb_ab_sum)))
        self.b_ab_tnacht_raum_mep.Text = str(int(round(float(self.mepraum.tnb_ab.soll))))
        if self.mepraum.tiefenachtbetrieb:
            self.b_ab_tnacht_ueber_mep.Text = str(int(round(float(self.mepraum.ueber_aus.soll))))
            self.b_ab_tnacht_ueber_m_mep.Text = str(int(round(float(self.mepraum.ueber_aus_manuell.soll))))
            self.b_ab_tnacht_lab_mep.Text = str(int(round(float(self.mepraum.ab_lab_min.soll))))
            self.b_ab_24h_tnacht_mep.Text = str(int(round(float(self.mepraum.ab_24h.soll))))
            self.tnacht_Abweichung_mep.Text = str(int(round(tnb_abweichung)))
            #if abs(round(tnb_abweichung-self.mepraum.Druckstufe.soll)) <= 3:
            if abs(round(tnb_abweichung)) <= 3:
                self.tnacht_auswertung_mep.Text = 'OK'
                self.tnacht_Abweichung_mep.Foreground = self.black
                self.tnacht_auswertung_mep.Foreground = self.black
            else:
                self.tnacht_auswertung_mep.Text = 'Passt Nicht'
                self.tnacht_auswertung_mep.Foreground = self.red
                self.tnacht_Abweichung_mep.Foreground = self.red
        else:
            self.b_ab_tnacht_ueber_mep.Text = '0'
            self.b_ab_tnacht_ueber_m_mep.Text = '0'
            self.b_ab_tnacht_lab_mep.Text = '0'
            self.b_ab_24h_tnacht_mep.Text = '0'
            self.tnacht_Abweichung_mep.Text = '0'
            self.tnacht_auswertung_mep.Text = 'OK'
            self.tnacht_Abweichung_mep.Foreground = self.black
            self.tnacht_auswertung_mep.Foreground = self.black

    def nachtbetriebchanged(self, sender, args):
        try:
            self.mepraum.nachtbetrieb = 1 if sender.IsChecked else 0
            self.mepraum.Nachtbetrieb_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except Exception as e:print(e)

    def tnachtbetriebchanged(self, sender, args):
        try:
            self.mepraum.tiefenachtbetrieb = 1 if sender.IsChecked else 0
            self.mepraum.Nachtbetrieb_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except Exception as e:print(e)

    def labmineingabe_changed(self, sender, args):
        try:
            self.mepraum.ab_lab_min.soll = round(float(str(sender.Text).replace(',', '.')))
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass
    def labmaxeingabe_changed(self, sender, args):
        try:
            self.mepraum.ab_lab_max.soll = round(float(str(sender.Text).replace(',', '.')))
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass

    def ab24heingabe_changed(self, sender, args):
        try:
            self.mepraum.ab_24h.soll = round(float(str(sender.Text).replace(',', '.')))
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass
    def druckeingabe_changed(self, sender, args):
        try:
            self.mepraum.Druckstufe.soll = round(float(str(sender.Text).replace(',', '.')))
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass
    
    def hoehechanged(self, sender, args):
        try:
            self.mepraum.hoehe = int(float(str(sender.Text)))
            self.mepraum.volumen = round(self.mepraum.flaeche * self.mepraum.hoehe / 1000,2)
            self.volumen.Text = str(self.mepraum.volumen)
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except Exception as e:print(e)

    def LW_nacht_changed(self, sender, args):
        try:
            self.mepraum.NB_LW = float(str(sender.Text).replace(',', '.'))
            self.mepraum.Nachtbetrieb_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass

    def LW_tnacht_changed(self, sender, args):
        try:
            self.mepraum.T_NB_LW = float(str(sender.Text).replace(',', '.'))
            self.mepraum.Nachtbetrieb_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass

    def von_nacht_changed(self, sender, args):
        try:
            self.mepraum.nb_von.soll = float(str(sender.Text).replace(',', '.'))
            self.mepraum.Nachtbetrieb_Berechnen()
            # self.grundinfo.Items.Refresh()
        except:pass

    def bis_nacht_changed(self, sender, args):
        try:
            self.mepraum.nb_bis.soll = float(str(sender.Text).replace(',', '.'))
            self.mepraum.Nachtbetrieb_Berechnen()
            # self.grundinfo.Items.Refresh()
        except:pass

    def von_tnacht_changed(self, sender, args):
        try:
            self.mepraum.tnb_von.soll = float(str(sender.Text).replace(',', '.'))
            self.mepraum.Nachtbetrieb_Berechnen()
            # self.grundinfo.Items.Refresh()
        except:pass

    def bis_tnacht_changed(self, sender, args):
        try:
            self.mepraum.tnb_bis.soll = float(str(sender.Text).replace(',', '.'))
            self.mepraum.TiefeNachtbetrieb_Berechnen()
            # self.grundinfo.Items.Refresh()
        except:pass

    def faktorchanged(self, sender, args):
        try:
            self.mepraum.faktor = float(str(sender.Text).replace(',', '.'))
        except:pass
        try:
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
            # self.grundinfo.Items.Refresh()
            # self.anlageninfo.Items.Refresh()
            # self.schachtinfo.Items.Refresh()
        except:pass

    def bezugsnameselectionchanged(self, sender, args):
        self.mepraum.bezugsname = self.bezugsname.SelectedItem.ToString()
        for el in self.mepraum.berechnung_nach.keys():
            if self.mepraum.berechnung_nach[el] == self.bezugsname.SelectedItem.ToString():
                self.mepraum.bezugsnummer = el
        self.einheit.Text = self.mepraum.einheit = self.mepraum.einheit_liste[self.mepraum.bezugsnummer]
        try:
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Nachtbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
            self.Auswertung_MEP()
        except:pass

    def MEP_changed(self, sender, args):
        try:
            text = str(self.Raumnr.SelectedItem)
            
            self.rauman.IsEnabled = True
            self.rauman1.IsEnabled = True
            
            self.temp_luftauslass.Clear()
            self.temp_vsr.Clear()
            self.temp_Iris.Clear()
            
            try:
                self.mepraum = self.mepraum_liste[text]
                self.warnung.Visibility= self.hidden
                self.auswahl.Visibility= self.hidden
           
                for el in self.mepraum.list_vsr:
                    el.checked = False
                try:self.set_up()
                except Exception as e:print(e)
            except Exception as ex: print(ex)
        except Exception as exx:
            print(exx)

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
            self.temp_Iris.Clear()
            self.temp_vsr.Clear()
            for el in temp.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass

            if temp.IsHaupt:
                for el_vsr in temp.List_VSR:
                    if self.mepraum.list_vsr1.Contains(el_vsr):
                        self.temp_vsr.Add(el)
                self.lv_vsr1_getrennt.ItemsSource = self.temp_vsr
                self.lv_vsr1.ItemsSource = self.temp_vsr

                for el_iris in temp.List_Iris:
                    if self.mepraum.list_vsr2.Contains(el_iris):
                        self.temp_Iris.Add(el)
                self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
                self.lv_vsr2.ItemsSource = self.temp_Iris
            elif temp.IsIris:
                pass
            else:
                for el_iris in temp.List_Iris:
                    if self.mepraum.list_vsr2.Contains(el_iris):
                        self.temp_Iris.Add(el)
                self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
                self.lv_vsr2.ItemsSource = self.temp_Iris

        except Exception as e:print(e)

    def VSR_checked_changed(self, sender, args):
        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        Checked = sender.IsChecked
        item = sender.DataContext

        self.temp_luftauslass.Clear()
        self.temp_vsr.Clear()
        self.temp_Iris.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            for el in self.lv_vsr1.Items:
                el.checked = False
            for el in self.lv_vsr2.Items:
                el.checked = False
            item.checked = True

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass
            

            for el_vsr in item.List_VSR:
                if self.mepraum.list_vsr1.Contains(el_vsr):
                    self.temp_vsr.Add(el_vsr)
            self.lv_vsr1_getrennt.ItemsSource = self.temp_vsr
            self.lv_vsr1.ItemsSource = self.temp_vsr

            for el_iris in item.List_Iris:
                if self.mepraum.list_vsr2.Contains(el_iris):
                    self.temp_Iris.Add(el_iris)
            self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
            self.lv_vsr2.ItemsSource = self.temp_Iris

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
            self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass
            self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
            self.lv_vsr2.ItemsSource = self.mepraum.list_vsr2
            self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
            self.lv_vsr1.ItemsSource = self.mepraum.list_vsr1
    
    def VSR1_checked_changed(self, sender, args):
        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2.ItemsSource = self.mepraum.list_vsr1
        Checked = sender.IsChecked
        item = sender.DataContext
        self.temp_luftauslass.Clear()
        self.temp_Iris.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            for el in self.lv_vsr1.Items:
                el.checked = False
            for el in self.lv_vsr2.Items:
                el.checked = False
            item.checked = True

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass
            
            for el_iris in item.List_Iris:
                if self.mepraum.list_vsr2.Contains(el_iris):
                    self.temp_Iris.Add(el_iris)
            self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
            self.lv_vsr2.ItemsSource = self.temp_Iris

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
            self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass
            self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
            self.lv_vsr2.ItemsSource = self.mepraum.list_vsr2
    
    def VSR2_checked_changed(self, sender, args):
        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
        self.lv_vsr1.ItemsSource = self.mepraum.list_vsr2
        Checked = sender.IsChecked
        item = sender.DataContext
        self.temp_luftauslass.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            for el in self.lv_vsr1.Items:
                el.checked = False
            for el in self.lv_vsr2.Items:
                el.checked = False
            item.checked = True

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass
            

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
            self.lv_auslass.ItemsSource = self.mepraum.list_auslass
            self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass
            self.raumwechsel_hinweis.Visibility= self.hidden
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
                  
    def Auslass_Volumen_changed_max(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    if el.art != '24h':
                        el.Luftmengenmax = item.Luftmengenmax
                        self.Auslass_Volumen_changed(el)
        else:
            self.Auslass_Volumen_changed(item)
    
    def Auslass_Volumen_changed_tnacht(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    if el.art != '24h':
                        el.Luftmengentnacht = item.Luftmengentnacht
                        self.Auslass_Volumen_changed(el)
        else:
            self.Auslass_Volumen_changed(item)
    
    def Auslass_Volumen_changed_min(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    el.Luftmengenmin = item.Luftmengenmin
                    if el.art == '24h':
                        el.Luftmengentnacht = item.Luftmengenmin
                        el.Luftmengenmax = item.Luftmengenmin
                        if self.mepraum.tiefenachtbetrieb:el.Luftmengennacht = item.Luftmengenmin
                    self.Auslass_Volumen_changed(el)
        else:
            if item.art == '24h':
                item.Luftmengentnacht = item.Luftmengenmin
                item.Luftmengenmax = item.Luftmengenmin
                if self.mepraum.tiefenachtbetrieb:item.Luftmengennacht = item.Luftmengenmin
            self.Auslass_Volumen_changed(item)
    
    def Auslass_Volumen_changed_nacht(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    if el.art != '24h':
                        el.Luftmengennacht = item.Luftmengennacht
                        self.Auslass_Volumen_changed(el)
        else:
            self.Auslass_Volumen_changed(item)

    def Auslass_Volumen_changed(self, auslass):
        try:
            vsr = auslass.VSR_Class
            if vsr:
                vsr.Luftmengenermitteln_new()
                vsr.vsrauswaelten()
                vsr.vsrueberpruefen()
            iris = auslass.IRIS_Class
            if iris:
                iris.Luftmengenermitteln_new()
            haupt = auslass.Haupt_Class
            if haupt:
                haupt.Luftmengenermitteln_new()
                haupt.vsrauswaelten()
                haupt.vsrueberpruefen()
        except:pass
        try:
            self.Auswertung_System()
        except:pass

    def Auslass_Selected_Changed(self, sender, args):
        if self.lv_auslass.SelectedIndex != -1 and self.lv_auslass_getrennt.SelectedIndex != self.lv_auslass.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_auslass.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Auslass')
            self.lv_auslass_getrennt.SelectedIndex = self.lv_auslass.SelectedIndex
        else:
            return
    
    def Auslass_Selected_Changed_getrennt(self, sender, args):
        if self.lv_auslass_getrennt.SelectedIndex != -1 and self.lv_auslass_getrennt.SelectedIndex != self.lv_auslass.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_auslass_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Auslass')
            self.lv_auslass.SelectedIndex = self.lv_auslass_getrennt.SelectedIndex
        else:
            return
    
    def VSR_Selected_Changed(self, sender, args):
        if self.lv_vsr.SelectedIndex != -1 and self.lv_vsr_getrennt.SelectedIndex != self.lv_vsr.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Haupt')
            self.lv_vsr_getrennt.SelectedIndex = self.lv_vsr.SelectedIndex
        else:
            return
    
    def VSR_Selected_Changed_getrennt(self, sender, args):
        if self.lv_vsr_getrennt.SelectedIndex != -1 and self.lv_vsr_getrennt.SelectedIndex != self.lv_vsr.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Haupt')
            self.lv_vsr.SelectedIndex = self.lv_vsr_getrennt.SelectedIndex
        else:
            return   
    
    def VSR1_Selected_Changed(self, sender, args):
        if self.lv_vsr1.SelectedIndex != -1 and self.lv_vsr1_getrennt.SelectedIndex != self.lv_vsr1.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr1.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('VSR')
            self.lv_vsr1_getrennt.SelectedIndex = self.lv_vsr1.SelectedIndex
        else:
            return
    
    def VSR1_Selected_Changed_getrennt(self, sender, args):
        if self.lv_vsr1_getrennt.SelectedIndex != -1 and self.lv_vsr1_getrennt.SelectedIndex != self.lv_vsr1.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr1_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('VSR')
            self.lv_vsr1.SelectedIndex = self.lv_vsr1_getrennt.SelectedIndex
        else:
            return
    
    def VSR2_Selected_Changed(self, sender, args):
        if self.lv_vsr2.SelectedIndex != -1 and self.lv_vsr2_getrennt.SelectedIndex != self.lv_vsr2.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr2.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Iris')
            self.lv_vsr2_getrennt.SelectedIndex = self.lv_vsr2.SelectedIndex
        else:
            return
   
    def VSR2_Selected_Changed_getrennt(self, sender, args):
        if self.lv_vsr2_getrennt.SelectedIndex != -1 and self.lv_vsr2_getrennt.SelectedIndex != self.lv_vsr2.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr2_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Iris')
            self.lv_vsr2.SelectedIndex = self.lv_vsr2_getrennt.SelectedIndex
        else:
            return

    def Ueber_aus_Selected_Changed(self, sender, args):
        if self.lv_ueber_aus.SelectedIndex != -1:
            self.externaleventliste.Liste = [self.lv_ueber_aus.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Üeberaus')
        else:
            return

    def Ueber_in_Selected_Changed(self, sender, args):
        if self.lv_ueber_in.SelectedIndex != -1:
            self.externaleventliste.Liste = [self.lv_ueber_in.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Üeberin')
        else:
            return
    
    def SelectedIndex(self,Art):
        if Art != 'Haupt':
            self.lv_vsr.SelectedIndex = -1
            self.lv_vsr_getrennt.SelectedIndex = -1
        if Art != 'VSR':
            self.lv_vsr1.SelectedIndex = -1
            self.lv_vsr1_getrennt.SelectedIndex = -1
        if Art != 'Iris':
            self.lv_vsr2.SelectedIndex = -1
            self.lv_vsr2_getrennt.SelectedIndex = -1
        if Art != 'Auslass':
            self.lv_auslass.SelectedIndex = -1
            self.lv_auslass_getrennt.SelectedIndex = -1
        if Art != 'Üeberaus':
            self.lv_ueber_aus.SelectedIndex = -1
        if Art != 'Üeberin':
            self.lv_ueber_in.SelectedIndex = -1

    def abbrechen(self, sender, args):
        self.externalevent.Dispose()
        self.Close()
    
    def exportieren(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.ExportRaumluftbilanz
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def exportieren_schacht(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.ExportLuftmengenInSchacht
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def raumanzeigen(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.RaumAnzeigen
            self.externalevent.Raise()
        except Exception as e:print(e)

    def bauteilundraumanzeigen(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.RaumundBauteileanzeigen
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def Berechnen(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.RaumluftBerechnen

            self.externalevent.Raise()
        except Exception as e:print(e)

    def changevsrtype(self, sender, args):
        try:
            
            self.externaleventliste.Executeapp = self.externaleventliste.vsranpassen
            self.externalevent.Raise()
        except Exception as e:print(e)

    def luftmenge_gleich_verteilen_mep(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.LuftmengeverteilenMEP
            self.externalevent.Raise()
        except Exception as e:print(e)

    def luftmenge_gleich_verteilen_pro(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.LuftmengeverteilenProjekt
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def aenderung_uebernehmen_mep(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.AederungUebernehmenMEP
            self.externalevent.Raise()
        except Exception as e:print(e)

    def aenderung_uebernehmen_pro(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.AederungUebernehmenProjekt
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def aus_revit_uebernehmen(self, sender, args):
        try:
            for el in self.mepraum.list_auslass:
                el.set_up()

            for el in self.mepraum.list_vsr:
                el.set_up()
                el.Luftmengenermitteln_new()
            
            self.Auswertung_System()
            self.Auswertung_MEP()
        except Exception as e:print(e)
    
    def Textinput0(self, sender, args):
        try:args.Handled = self.regex2.IsMatch(args.Text)
        except:args.Handled = True
    
    def Textinput1(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find(',') != -1:
                if args.Text == ',':args.Handled = True
                return
        try:args.Handled = self.regex1.IsMatch(args.Text)
        except:args.Handled = True
    
    def Textinput2(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find('-') != -1:
                if args.Text == '-':args.Handled = True
                return
        try:args.Handled = self.regex1.IsMatch(args.Text)
        except:args.Handled = True
    
    def buero(self,sender,a):
        self.LW_nacht.Text = '0'
        self.LW_tnacht.Text = '0'
        self.isnachtbetrieb.IsChecked = False
        self.istiefenachtbetrieb.IsChecked = False
        
    
    def lab1(self,sender,a):
        self.LW_nacht.Text = '4'
        self.LW_tnacht.Text = '0'
        self.isnachtbetrieb.IsChecked = True
        self.istiefenachtbetrieb.IsChecked = False
        
    
    def lab2(self,sender,a):
        self.LW_nacht.Text = '4'
        self.LW_tnacht.Text = '2'
        self.isnachtbetrieb.IsChecked = True
        self.istiefenachtbetrieb.IsChecked = True
        
    def lab3(self,sender,a):
        self.LW_nacht.Text = '2'
        self.LW_tnacht.Text = '2'
        self.isnachtbetrieb.IsChecked = True
        self.istiefenachtbetrieb.IsChecked = True
        


mepWPF = MEPRaum_Uebersicht()
mepWPF.externaleventliste.GUI = mepWPF 
mepWPF.Show()


# from eventhandler import Raumdaten,IGF_LOGO,xlsxwriter
# path = r'C:\Users\Zhang\Desktop\test12.xlsx'
# e = xlsxwriter.Workbook(path)
# worksheet = e.add_worksheet()

# n = 0
# rowstart = 0
# Liste_Pagebreaks = []
# for mepname in sorted(DICT_MEP_ITEMSSOIRCE.keys()):  
#     mep = DICT_MEP_ITEMSSOIRCE[mepname]
#     # raumdaten = Raumdaten(mep,rowstart)
#     raumdaten = Raumdaten(mep,rowstart)
#     raumdaten.book = e
#     raumdaten.sheet = worksheet
#     # raumdaten.exportheader1()
#     raumdaten.GetFinalExportdaten()
#             # while (raumdaten.rowende - rowstart > 36):
#             #     Liste_Pagebreaks.append(rowstart+36)
#             #     rowstart += 36
#     rowstart = raumdaten.rowende

#     Liste_Pagebreaks.extend(raumdaten.rowbreaks)

# worksheet.set_landscape()
# worksheet.set_column(0,1,17)
# worksheet.set_column(2,5,13)
# worksheet.set_column(6,6,7)
# worksheet.set_column(7,7,13)
# worksheet.set_column(8,10,7)
# worksheet.set_column(11,18,11)
# worksheet.set_column(19,19,18)
# worksheet.set_column(20,20,25)
# worksheet.set_paper(9)
# worksheet.set_h_pagebreaks(Liste_Pagebreaks)   
# header2 = '&L&G'  
# worksheet.set_footer('&C&p / &N')   
# worksheet.set_margins(top=1)
# worksheet.set_header(header2, {'image_left': IGF_LOGO})
# e.close() 
