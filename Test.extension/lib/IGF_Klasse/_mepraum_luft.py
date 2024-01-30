# coding: utf8
from IGF_Klasse import DB,ItemTemplate,TemplateItemBase
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from IGF_Funktionen._Parameter import get_value
from System import Guid
from IGF_Allgemein import Parameter_Dict,Luftberechnung_NachName,Luftberechnung_NachNummer,Luftberechnung_Einheit
from rpw import db
from IGF_Funktionen._Parameter import get_value,get_Parameter,wert_schreibenbase
from pyrevit import script

class MEPRaum(object):
    def __init__(self, elem, list_vsr,LISTE_SCHACHT,DICT_MEP_AUSLASS,DICT_MEP_UEBERSTROM,Dict_Ueber_Manuell,DICT_MEP_UN_AUNLASS,_Einbauteile,_VSR):
        self.logger = script.get_logger()
        self.elem = elem
        




        self.Dict_Ueber_Manuell = Dict_Ueber_Manuell
        self.DICT_MEP_UN_AUNLASS = DICT_MEP_UN_AUNLASS
        self.DICT_MEP_AUSLASS = DICT_MEP_AUSLASS
        self.DICT_MEP_UEBERSTROM = DICT_MEP_UEBERSTROM
        self.Einbauteile = _Einbauteile
        self.VSR = _VSR
        self.berechnung_nach = {
                        '1': "Fläche",
                        '2': "Luftwechsel",
                        '3': "Person",
                        '4': "manuell",
                        '5': "nurZUMa",
                        '6': "nurABMa",
                        '5.1': "nurZU_Fläche",
                        '5.2': "nurZU_Luftwechsel",
                        '5.3': "nurZU_Person",
                        '6.1': "nurAB_Fläche",
                        '6.2': "nurAB_Luftwechsel",
                        '6.3': "nurAB_Person",
                        '9': "keine",
                    }
        self.einheit_liste = {
                    '1': 'm³/h pro m²',
                    '2': '-1/h',
                    '3': 'm3/h pro P',
                    '4': 'm³/h ',
                    '5': 'm³/h ',
                    '6': 'm³/h' ,
                    '5.1': "m³/h pro m²",
                    '5.2': '-1/h',
                    '5.3': 'm3/h pro P',
                    '6.1': "m³/h pro m²",
                    '6.2': '-1/h',
                    '6.3': 'm3/h pro P',
                    '9': '',
                }
        self.liste_schacht = LISTE_SCHACHT
        self.elemid = elem.Id
        self.elem = elem
        self.Raumnr = self.elem.Number + ' - ' + self.elem.LookupParameter('Name').AsString()

        self.list_vsr = list_vsr
        self.list_vsr0 = ObservableCollection[VSR]() # Haupt
        self.list_vsr1 = ObservableCollection[VSR]() # VSR
        self.list_vsr2 = ObservableCollection[VSR]() # Iris

        self.Liste_RaumluftUnrelevant = ObservableCollection[Luftauslass]()
        
        self.Liste_RZU = []
        self.Liste_RAB = []
        self.Liste_TZU = []
        self.Liste_TAB = []
        self.Liste_H24 = []
        self.Liste_LAB = []
        self.Liste_UIN = []
        self.Liste_UAU = []
        self.Liste_ZU = []
        self.Liste_AB = []

        self.Dict_RZU = []
        self.Dict_RAB = []

        self.Dict_Einbauteile = {}
        self.Dict_VSR = {}

        self.IGF_Legende = ''     


        # Übersicht
        self.Uebersicht = ObservableCollection[MEPGrundInfo]()
        # Anlagen
        self.Anlagen_info = ObservableCollection[MEPAnlagenInfo]()
        # Schacht
        self.Schacht_info = ObservableCollection[MEPSchachtInfo]()
        # Detail
        self.Detail_Min = ObservableCollection[MEPAuswertung]()
        self.Detail_Max = ObservableCollection[MEPAuswertung]()
        self.Detail_Nacht = ObservableCollection[MEPAuswertung]()
        self.Detail_Tnacht = ObservableCollection[MEPAuswertung]()

        try:self.list_ueber = self.DICT_MEP_UEBERSTROM[self.elemid.ToString()]
        except:self.list_ueber = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        
        # Grundinfo.
        self.bezugsnummer = self.elem.LookupParameter('IGF_RLT_VolumenstromProNummer').AsValueString()

        try:self.bezugsname = self.berechnung_nach[self.bezugsnummer]
        except:self.bezugsname = 'keine'

        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''

        self.flaeche = round(self.get_element('Fläche'), 2)
        self.volumen = round(self.get_element('Volumen'),2)
        self.hoehe = int(self.get_element('Lichte Höhe'))
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('IGF_RLT_VolumenstromProFaktor')

        if self.flaeche == 0 and self.volumen == 0:
            return

        try:self.einheit = self.einheit_liste[self.bezugsnummer]
        except:self.einheit = ''

        # self.hoehe = int(self.get_element('IGF_A_Lichte_Höhe'))
        # self.volumen = round((self.hoehe*self.flaeche/1000),2)
        # self.hoehe = round((self.volumen/self.flaeche),2)
        self.nachtbetrieb = self.get_element('IGF_RLT_Nachtbetrieb')
        self.tiefenachtbetrieb = self.get_element('IGF_RLT_TieferNachtbetrieb')

        self.ueber_Zu_Tag = self.get_element('IGF_RLT_ÜberströmungIst_Tag_ZU')
        self.ueber_Zu_Nacht = self.get_element('IGF_RLT_ÜberströmungIst_Nacht_ZU')

        self.NB_LW = self.get_element('IGF_RLT_NachtbetriebLW')
        self.T_NB_LW = self.get_element('IGF_RLT_TieferNachtbetriebLW')
        
        self.istreduziert = self.elem.LookupParameter('IGF_RLT_istReduziert').AsInteger()
        self.vol_faktorreduzierung = self.get_element('IGF_RLT_Raum-ReduziertFaktor')
        self.vol_faktorreduzierung_berechnen = 1
        if not self.vol_faktorreduzierung or self.vol_faktorreduzierung == 0:
            self.vol_faktorreduzierung = 1
        else:
            self.vol_faktorreduzierung_berechnen = self.vol_faktorreduzierung

        self.IsTierRaum = (self.get_element('IGF_RLT_Tierkäfig_raumunabhängig') != 0)
        self.IsSchacht = (self.get_element('IGF_RLT_InstallationsSchacht') != 0)

        self.schachtname = self.elem.LookupParameter('IGF_RLT_InstallationsSchachtName').AsString()
        self.list_auslass = ObservableCollection[Luftauslass]()

        self.ueber_in_m_class = UeberStromAuslass_Manuell(round(self.get_element('IGF_RLT_RaumÜberströmungMenge')))
        self.ueber_aus_m_class = UeberStromAuslass_Manuell()

        self._vsr_bearbeiten()
        self._auslass_bearbeiten()
        self._auslass_analyse()
    
        # Übersicht 
        self.angezuluft = MEPGrundInfo('Ange.Zuluft','Angegebener Zuluftstrom')
        self.angeabluft = MEPGrundInfo('Ange.Abluft','Angegebener Abluftluftstrom')
        self.ab_24h = MEPGrundInfo('24h Ab','IGF_RLT_AbluftSumme24h',round(self.get_element('IGF_RLT_AbluftSumme24h')),self.Liste_H24)
        self.ab_min = MEPGrundInfo('min.RAB','IGF_RLT_AbluftminRaum-IGF_RLT_AbluftminSummeLabor-IGF_RLT_AbluftSumme24h,self.Liste_RZU','',self.Liste_RAB)
        self.ab_lab_min = MEPGrundInfo('min.LAB','IGF_RLT_AbluftminSummeLabor',round(self.get_element('IGF_RLT_AbluftminSummeLabor')),self.Liste_LAB)
        self.zu_lab_min = MEPGrundInfo('min.LZU','IGF_RLT_ZuluftminSummeLabor',round(self.get_element('IGF_RLT_ZuluftminSummeLabor')),self.Liste_LAB)
        self.zu_min = MEPGrundInfo('min.RZU','IGF_RLT_ZuluftminRaum','',self.Liste_RZU)
        self.ab_max = MEPGrundInfo('max.RAB','IGF_RLT_AbluftmaxRaum-IGF_RLT_AbluftmaxSummeLabor-IGF_RLT_AbluftSumme24h',0,self.Liste_RAB)
        self.ab_lab_max = MEPGrundInfo('max.LAB','IGF_RLT_AbluftmaxSummeLabor',round(self.get_element('IGF_RLT_AbluftmaxSummeLabor')),self.Liste_LAB)
        self.zu_lab_max = MEPGrundInfo('max.LZU','IGF_RLT_ZuluftmaxSummeLabor',round(self.get_element('IGF_RLT_ZuluftmaxSummeLabor')),self.Liste_LAB)
        self.zu_max = MEPGrundInfo('max.RZU','IGF_RLT_ZuluftmaxRaum',0,self.Liste_RZU)
        self.nb_von = MEPGrundInfo('NB von','IGF_RLT_NachtbetriebVon',self.get_element('IGF_RLT_NachtbetriebVon'))
        self.nb_bis = MEPGrundInfo('NB bis','IGF_RLT_NachtbetriebBis',self.get_element('IGF_RLT_NachtbetriebBis'))
        self.nb_dauer = MEPGrundInfo('NB Dauer','IGF_RLT_NachtbetriebDauer')
        self.tnb_von = MEPGrundInfo('TNB von','IGF_RLT_TieferNachtbetriebVon',self.get_element('IGF_RLT_TieferNachtbetriebVon'))
        self.tnb_bis = MEPGrundInfo('TNB bis','IGF_RLT_TieferNachtbetriebBis',self.get_element('IGF_RLT_TieferNachtbetriebBis'))
        self.tnb_dauer = MEPGrundInfo('TNB Dauer','IGF_RLT_TieferNachtbetriebDauer')
        self.tier_zu_min = MEPGrundInfo('min.TZU','IGF_RLT_Luftmenge_min_TZU',self.get_element('IGF_RLT_Luftmenge_min_TZU'),self.Liste_TZU)
        self.tier_ab_min = MEPGrundInfo('min.TAB','IGF_RLT_Luftmenge_min_TAB',self.get_element('IGF_RLT_Luftmenge_min_TAB'),self.Liste_TAB)
        self.tier_zu_max = MEPGrundInfo('max.TZU','IGF_RLT_Luftmenge_max_TZU',self.get_element('IGF_RLT_Luftmenge_max_TZU'),self.Liste_TZU)
        self.tier_ab_max = MEPGrundInfo('max.TAB','IGF_RLT_Luftmenge_max_TAB',self.get_element('IGF_RLT_Luftmenge_max_TAB'),self.Liste_TAB)
        self.ueber_in = MEPGrundInfo('Über. Ein.','IGF_RLT_ÜberströmungSummeIn',round(self.get_element('IGF_RLT_ÜberströmungSummeIn'),1),self.Liste_UIN)
        self.ueber_aus = MEPGrundInfo('Über. Aus.','IGF_RLT_ÜberströmungSummeAus',round(self.get_element('IGF_RLT_ÜberströmungSummeAus'),1),self.Liste_UAU)
        self.ueber_in_manuell = MEPGrundInfo('Über.Ein.M.','IGF_RLT_RaumÜberströmungMenge',round(self.get_element('IGF_RLT_RaumÜberströmungMenge')))
        self.ueber_aus_manuell = MEPGrundInfo('Über.Aus.M.','IGF_RLT_RaumÜberströmungMenge',round(self.get_element('IGF_RLT_RaumÜberströmungMenge')))
        self.Druckstufe = MEPGrundInfo('Druckstufe','IGF_RLT_RaumDruckstufeEingabe',round(self.get_element('IGF_RLT_RaumDruckstufeEingabe')))

        self.ab_minsum = MEPGrundInfo('min.AB_SUM','RAB,LAB,24h,Über,Über_M','',[self.ab_min,self.ab_lab_min,self.ab_24h,self.ueber_aus,self.ueber_aus_manuell])
        self.zu_minsum = MEPGrundInfo('min.ZU_SUM','RZU,Über,Über_M','',[self.zu_min,self.ueber_in,self.ueber_in_manuell])
        self.ab_maxsum = MEPGrundInfo('max.AB_SUM','RAB,LAB,24h,Über,Über_M','',[self.ab_max,self.ab_lab_max,self.ab_24h,self.ueber_aus,self.ueber_aus_manuell])
        self.zu_maxsum = MEPGrundInfo('max.ZU_SUM','RZU,Über,Über_M','',[self.zu_max,self.ueber_in,self.ueber_in_manuell])
        self.nb_zu = MEPGrundInfo('NB Zu','IGF_RLT_ZuluftNachtRaum',0,self.Liste_ZU)
        self.nb_ab = MEPGrundInfo('NB Ab','IGF_RLT_AbluftNachtRaum',0,self.Liste_AB)
        self.tnb_zu = MEPGrundInfo('TNB Zu','IGF_RLT_ZuluftTieferNachtRaum',0,self.Liste_ZU)
        self.tnb_ab = MEPGrundInfo('TNB Ab','IGF_RLT_AbluftTieferNachtRaum',0,self.Liste_AB)
        self.ueber_sum = MEPGrundInfo('Überstrom','IGF_RLT_ÜberströmungRaum',round(self.get_element('IGF_RLT_ÜberströmungRaum')),[self.ueber_in,self.ueber_aus,self.ueber_aus_manuell,self.ueber_in_manuell])

        self.rzu_Schacht = MEPSchachtInfo('RZU',self.get_element('IGF_RLT_Schacht_RZU'),self.zu_max,self.liste_schacht)
        self.rab_Schacht = MEPSchachtInfo('RAB',self.get_element('IGF_RLT_Schacht_RAB'),self.ab_max,self.liste_schacht)
        self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.tier_zu_max,self.liste_schacht)
        self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.tier_ab_max,self.liste_schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('IGF_RLT_Schacht_24h'),self.ab_24h,self.liste_schacht)
        self.lab_Schacht = MEPSchachtInfo('LAB',self.get_element('IGF_RLT_Schacht_LAB'),self.ab_lab_max,self.liste_schacht)


        # Grundinfo_für Detailberechnen
        self.nb_ueber_in =  MEPGrundInfo('nb.Über_in','','',self.Liste_UIN)
        self.nb_ueber_in_m =  MEPGrundInfo('nb.Über_in_m','','',self.ueber_in_m_class)
        self.nb_rab =  MEPGrundInfo('nb.RAB','','',self.Liste_RAB)
        self.nb_ueber_aus =  MEPGrundInfo('nb.Über_aus','','',self.Liste_UAU)
        self.nb_ueber_aus_m =  MEPGrundInfo('nb.Über_aus_m','','',self.ueber_aus_m_class)
        self.nb_lab =  MEPGrundInfo('nb.LAB','','',self.Liste_LAB)
        self.nb_24h =  MEPGrundInfo('nb.24h','','',self.Liste_H24)
        self.nb_druckstufe =  MEPGrundInfo('nb.Druck','','',[])
        self.tnb_ueber_in =  MEPGrundInfo('tnb.Über_in','','',self.Liste_UIN)
        self.tnb_ueber_in_m =  MEPGrundInfo('tnb.Über_in_m','','',self.ueber_in_m_class)
        self.tnb_rab =  MEPGrundInfo('tnb.RAB','','',self.Liste_RAB)
        self.tnb_ueber_aus =  MEPGrundInfo('tnb.Über_aus','','',self.Liste_UAU)
        self.tnb_ueber_aus_m =  MEPGrundInfo('tnb.Über_aus_m','','',self.ueber_aus_m_class)
        self.tnb_lab =  MEPGrundInfo('tnb.LAB','','',self.Liste_LAB)
        self.tnb_24h =  MEPGrundInfo('tnb.24h','','',self.Liste_H24)
        self.tnb_druckstufe =  MEPGrundInfo('tnb.Druck','','',[])

        self.LEER = MEPAuswertung('','')
        self.LEER._set_up_soll()
        self.LEER._set_up_ist()

        self.D_T_MIN_Rzu   = MEPAuswertung('Zu-Raum',self.zu_min)
        self.D_T_MIN_ZuUe  = MEPAuswertung('Zu-Über',self.ueber_in)
        self.D_T_MIN_ZuUeM = MEPAuswertung('Zu-Über-Manuel',self.ueber_in_manuell)
        self.D_T_MIN_Rab   = MEPAuswertung('Ab-Raum',self.ab_min)
        self.D_T_MIN_AbUe  = MEPAuswertung('Ab-Über',self.ueber_aus)
        self.D_T_MIN_AbUeM = MEPAuswertung('Ab-Über-Manuel',self.ueber_aus_manuell)
        self.D_T_MIN_Lab   = MEPAuswertung('Ab-Labor',self.ab_lab_min)
        self.D_T_MIN_24h   = MEPAuswertung('Ab-24h',self.ab_24h)
        self.D_T_MIN_ZuS   = MEPAuswertung('Zuluft-Summe',[self.D_T_MIN_Rzu,self.D_T_MIN_ZuUe,self.D_T_MIN_ZuUeM])
        self.D_T_MIN_AbS   = MEPAuswertung('Abluft-Summe',[self.D_T_MIN_Rab,self.D_T_MIN_AbUe,self.D_T_MIN_AbUeM,self.D_T_MIN_24h,self.D_T_MIN_Lab])
        self.D_T_MIN_BLZ   = MEPAuswertung('Luftbilanz',[self.D_T_MIN_ZuS,self.D_T_MIN_AbS])
        self.D_T_Druckstufe= MEPAuswertung('Druckstufe',self.Druckstufe)
        self.D_T_MIN_Aus   = MEPAuswertung('Auswertung',[self.D_T_MIN_BLZ,self.D_T_Druckstufe])
        self.D_T_MIN_Tier  = MEPAuswertung('Tierhaltung',[])
        self.D_T_MIN_TZu   = MEPAuswertung('Zu-TH',self.tier_zu_min)
        self.D_T_MIN_TAb   = MEPAuswertung('Ab-TH',self.tier_ab_min)
        
        self.D_T_MAX_Rzu   = MEPAuswertung('Zu-Raum',self.zu_max)
        self.D_T_MAX_ZuUe  = MEPAuswertung('Zu-Über',self.ueber_in)
        self.D_T_MAX_ZuUeM = MEPAuswertung('Zu-Über-Manuel',self.ueber_in_manuell)
        self.D_T_MAX_Rab   = MEPAuswertung('Ab-Raum',self.ab_max)
        self.D_T_MAX_AbUe  = MEPAuswertung('Ab-Über',self.ueber_aus)
        self.D_T_MAX_AbUeM = MEPAuswertung('Ab-Über-Manuel',self.ueber_aus_manuell)
        self.D_T_MAX_Lab   = MEPAuswertung('Ab-Labor',self.ab_lab_max)
        self.D_T_MAX_24h   = MEPAuswertung('Ab-24h',self.ab_24h)
        self.D_T_MAX_ZuS   = MEPAuswertung('Zuluft-Summe',[self.D_T_MAX_Rzu,self.D_T_MAX_ZuUe,self.D_T_MAX_ZuUeM])
        self.D_T_MAX_AbS   = MEPAuswertung('Abluft-Summe',[self.D_T_MAX_Rab,self.D_T_MAX_AbUe,self.D_T_MAX_AbUeM,self.D_T_MAX_24h,self.D_T_MAX_Lab])
        self.D_T_MAX_BLZ   = MEPAuswertung('Luftbilanz',[self.D_T_MAX_ZuS,self.D_T_MAX_AbS])
        self.D_T_MAX_Aus   = MEPAuswertung('Auswertung',[self.D_T_MAX_BLZ,self.D_T_Druckstufe])
        self.D_T_MAX_TZu   = MEPAuswertung('Zu-TH',self.tier_zu_max)
        self.D_T_MAX_TAb   = MEPAuswertung('Ab-TH',self.tier_ab_max)

        self.D_T_NB_Rzu   = MEPAuswertung('Zu-Raum',self.nb_zu)
        self.D_T_NB_ZuUe  = MEPAuswertung('Zu-Über',self.nb_ueber_in)
        self.D_T_NB_ZuUeM = MEPAuswertung('Zu-Über-Manuel',self.nb_ueber_in_m)
        self.D_T_NB_Rab   = MEPAuswertung('Ab-Raum',self.nb_rab)
        self.D_T_NB_AbUe  = MEPAuswertung('Ab-Über',self.nb_ueber_aus)
        self.D_T_NB_AbUeM = MEPAuswertung('Ab-Über-Manuel',self.nb_ueber_aus_m)
        self.D_T_NB_Lab   = MEPAuswertung('Ab-Labor',self.nb_lab)
        self.D_T_NB_24h   = MEPAuswertung('Ab-24h',self.nb_24h)
        self.D_T_NB_ZuS   = MEPAuswertung('Zuluft-Summe',[self.D_T_NB_Rzu,self.D_T_NB_ZuUe,self.D_T_NB_ZuUeM])
        self.D_T_NB_AbS   = MEPAuswertung('Abluft-Summe',[self.D_T_NB_Rab,self.D_T_NB_AbUe,self.D_T_NB_AbUeM,self.D_T_NB_24h,self.D_T_NB_Lab])
        self.D_T_NB_BLZ   = MEPAuswertung('Luftbilanz',[self.D_T_NB_ZuS,self.D_T_NB_AbS])
        self.D_N_Druckstufe= MEPAuswertung('Druckstufe',self.nb_druckstufe)
        self.D_T_NB_Aus   = MEPAuswertung('Auswertung',[self.D_T_NB_BLZ,self.D_N_Druckstufe])

        self.D_T_TNB_Rzu   = MEPAuswertung('Zu-Raum',self.tnb_zu)
        self.D_T_TNB_ZuUe  = MEPAuswertung('Zu-Über',self.tnb_ueber_in)
        self.D_T_TNB_ZuUeM = MEPAuswertung('Zu-Über-Manuel',self.tnb_ueber_in_m)
        self.D_T_TNB_Rab   = MEPAuswertung('Ab-Raum',self.tnb_rab)
        self.D_T_TNB_AbUe  = MEPAuswertung('Ab-Über',self.tnb_ueber_aus)
        self.D_T_TNB_AbUeM = MEPAuswertung('Ab-Über-Manuel',self.tnb_ueber_aus_m)
        self.D_T_TNB_Lab   = MEPAuswertung('Ab-Labor',self.tnb_lab)
        self.D_T_TNB_24h   = MEPAuswertung('Ab-24h',self.tnb_24h)
        self.D_T_TNB_ZuS   = MEPAuswertung('Zuluft-Summe',[self.D_T_TNB_Rzu,self.D_T_TNB_ZuUe,self.D_T_TNB_ZuUeM])
        self.D_T_TNB_AbS   = MEPAuswertung('Abluft-Summe',[self.D_T_TNB_Rab,self.D_T_TNB_AbUe,self.D_T_TNB_AbUeM,self.D_T_TNB_24h,self.D_T_TNB_Lab])
        self.D_T_TNB_BLZ   = MEPAuswertung('Luftbilanz',[self.D_T_TNB_ZuS,self.D_T_TNB_AbS])
        self.D_TN_Druckstufe= MEPAuswertung('Druckstufe',self.tnb_druckstufe)
        self.D_T_TNB_Aus   = MEPAuswertung('Auswertung',[self.D_T_TNB_BLZ,self.D_N_Druckstufe])

        self.Grundinfo_detail = []
        self.Grundinfo_Summe0 = []
        self.Grundinfo_Summe1 = []

        self.Detail_detail = []
        self.Detail_summe0 = []
        self.Detail_summe1 = []
        self.Detail_summe2 = []

        self._uebersicht_bearbeiten()
        self._ueberstrom_manuel_bearbeiten()
        self._ueberstrom_bearbeiten()
        self._detail_bearbeiten()
        self._schacht_bearbeiten()
        self._anlagen_bearbeiten()
        self._liste_bearbeiten()
        self._unrelevant_bearbeiten()

        self.labnacht = 0
        self.labtnacht = 0
        self.ab24nacht = 0
        self.ab24tnacht = 0        
            
        self.Tagesbetrieb_Berechnen()
        self.Nachtbetrieb_Berechnen()
        
        self.update_default()
        self.update_ist_alle() 
        self.update_soll_alle() 
        self.Druckstufe_Berechnen()
             
    def update_default(self):
        self.D_T_MIN_AbUe._set_up_soll()
        self.D_T_MIN_ZuUe._set_up_soll()
        self.D_T_MAX_AbUe._set_up_soll()
        self.D_T_MAX_ZuUe._set_up_soll()
        self.D_T_MIN_24h._set_up_soll()
        self.D_T_MIN_Lab._set_up_soll()
        self.D_T_MAX_Lab._set_up_soll()
        self.D_T_MAX_24h._set_up_soll()
        self.D_T_NB_AbUe._set_up_soll()
        self.D_T_NB_ZuUe._set_up_soll()
        self.D_T_TNB_AbUe._set_up_soll()
        self.D_T_TNB_ZuUe._set_up_soll()
        self.D_T_MIN_AbUeM._set_up_soll()
        self.D_T_MIN_ZuUeM._set_up_soll()
        self.D_T_MAX_AbUeM._set_up_soll()
        self.D_T_MAX_ZuUeM._set_up_soll()
        self.D_T_NB_AbUeM._set_up_soll()
        self.D_T_NB_ZuUeM._set_up_soll()
        self.D_T_TNB_AbUeM._set_up_soll()
        self.D_T_TNB_ZuUeM._set_up_soll()
        self._lab_bearbeiten()
        self._24h_bearbeiten()
        self._druckstufe_bearbeiten()
        self.D_T_MIN_Tier._set_up_soll()

    def update(self):        
        self.ab_24h.soll = self.get_element('IGF_RLT_AbluftSumme24h')
        self.ab_lab_min.soll = self.get_element('IGF_RLT_AbluftminSummeLabor') 
        self.ab_lab_max.soll = self.get_element('IGF_RLT_AbluftmaxSummeLabor') 
        self.Druckstufe.soll = self.get_element('IGF_RLT_RaumDruckstufeEingabe')
    
    def update_soll(self,Liste):
        for el in Liste:
            el._set_up_soll()
    
    def update_ist(self,Liste):
        for el in Liste:
            el._set_up_ist()

    def update_soll_alle(self):
        self.update_soll(self.Grundinfo_detail)
        self.update_soll(self.Grundinfo_Summe0)
        self.update_soll(self.Grundinfo_Summe1)
        self.update_soll(self.Detail_detail)
        self.update_soll(self.Detail_summe0)
        self.update_soll(self.Detail_summe1)
        self.update_soll(self.Detail_summe2)
    
    def update_ist_alle(self):
        self.update_ist(self.Grundinfo_detail)
        self.update_ist(self.Grundinfo_Summe0)
        self.update_ist(self.Grundinfo_Summe1)
        self.update_ist(self.Detail_detail)
        self.update_ist(self.Detail_summe0)
        self.update_ist(self.Detail_summe1)
        self.update_ist(self.Detail_summe2)
    
    def update_soll_Tag(self):
        self.angezuluft._set_up_soll()
        self.angeabluft._set_up_soll()
        self.ab_minsum._set_up_soll()
        self.zu_minsum._set_up_soll()
        self.ab_maxsum._set_up_soll()
        self.zu_maxsum._set_up_soll()
        self.D_T_MIN_Rzu._set_up_soll()
        self.D_T_MIN_Rab._set_up_soll()
        self.D_T_MIN_ZuS._set_up_soll()
        self.D_T_MIN_AbS._set_up_soll()
        self.D_T_MIN_BLZ._set_up_soll()
        self.D_T_MIN_Aus._set_up_soll()
        self.D_T_MAX_Rzu._set_up_soll()
        self.D_T_MAX_Rab._set_up_soll()
        self.D_T_MAX_ZuS._set_up_soll()
        self.D_T_MAX_AbS._set_up_soll()
        self.D_T_MAX_BLZ._set_up_soll()
        self.D_T_MAX_Aus._set_up_soll()

    def update_Nacht_Ueber_Lab(self):
        self._ueberstrom_bearbeiten()
        self._lab_bearbeiten()
        self._druckstufe_bearbeiten()
        self._24h_bearbeiten()
        self.D_T_NB_ZuUe._set_up_soll()
        self.D_T_NB_ZuUeM._set_up_soll()
        self.D_T_NB_AbUe._set_up_soll()
        self.D_T_NB_AbUeM._set_up_soll()
        self.D_T_NB_Lab._set_up_soll()
        self.D_T_NB_24h._set_up_soll()
        self.D_T_TNB_ZuUe._set_up_soll()
        self.D_T_TNB_ZuUeM._set_up_soll()
        self.D_T_TNB_AbUe._set_up_soll()
        self.D_T_TNB_AbUeM._set_up_soll()
        self.D_T_TNB_Lab._set_up_soll()
        self.D_T_TNB_24h._set_up_soll()
        self.D_N_Druckstufe._set_up_ist()
        self.D_TN_Druckstufe._set_up_ist()

    def update_Tnacht_Ueber_Lab(self):
        self._ueberstrom_bearbeiten()
        self._lab_bearbeiten()
        self._druckstufe_bearbeiten()
        self._24h_bearbeiten()
        self.D_T_TNB_ZuUe._set_up_soll()
        self.D_T_TNB_ZuUeM._set_up_soll()
        self.D_T_TNB_AbUe._set_up_soll()
        self.D_T_TNB_AbUeM._set_up_soll()
        self.D_T_TNB_Lab._set_up_soll()
        self.D_T_TNB_24h._set_up_soll()
        self.D_TN_Druckstufe._set_up_ist()

    def update_lab_min(self):
        self._lab_bearbeiten()
        self.D_T_MIN_Lab._set_up_soll()
        self.D_T_NB_Lab._set_up_soll()
        self.D_T_TNB_Lab._set_up_soll()
    
    def update_lab_max(self):
        self.D_T_MAX_Lab._set_up_soll()
    
    def update_24h(self):
        self._24h_bearbeiten()
        self.D_T_MAX_24h._set_up_soll()
        self.D_T_MIN_24h._set_up_soll()
        self.D_T_NB_24h._set_up_soll()
        self.D_T_TNB_24h._set_up_soll()
    
    def update_druck(self):
        self._druckstufe_bearbeiten()
        self.Druckstufe._set_up_soll()
        self.D_T_Druckstufe._set_up_soll()
        self.D_N_Druckstufe._set_up_soll()
        self.D_TN_Druckstufe._set_up_soll()
        self.Druckstufe._set_up_ist()
        self.D_T_Druckstufe._set_up_ist()
        self.D_N_Druckstufe._set_up_ist()
        self.D_TN_Druckstufe._set_up_ist()

    def update_soll_Nacht(self):
        self.nb_von._set_up_soll()
        self.nb_bis._set_up_soll()
        self.nb_dauer._set_up_soll()
        self.tnb_von._set_up_soll()
        self.tnb_bis._set_up_soll()
        self.tnb_dauer._set_up_soll()
        self.nb_zu._set_up_soll()
        self.nb_ab._set_up_soll()
        self.tnb_zu._set_up_soll()
        self.tnb_ab._set_up_soll()

        self.D_T_NB_Rzu._set_up_soll()
        self.D_T_NB_Rab._set_up_soll()
        self.D_T_NB_ZuS._set_up_soll()
        self.D_T_NB_AbS._set_up_soll()
        self.D_T_NB_BLZ._set_up_soll()
        self.D_T_NB_Aus._set_up_soll()
        self.D_T_TNB_Rzu._set_up_soll()
        self.D_T_TNB_Rab._set_up_soll()
        self.D_T_TNB_ZuS._set_up_soll()
        self.D_T_TNB_AbS._set_up_soll()
        self.D_T_TNB_BLZ._set_up_soll()
        self.D_T_TNB_Aus._set_up_soll()
    
    def update_soll_summe(self):
        self.update_soll(self.Grundinfo_Summe0)
        self.update_soll(self.Grundinfo_Summe1)
        self.update_soll(self.Detail_summe0)
        self.update_soll(self.Detail_summe1)
        self.update_soll(self.Detail_summe2)
     
    def daten_update(self):
        def Detail_update(Liste):
            for el in Liste:
                el._set_up_soll()
                el._set_up_ist()
        Detail_update(self.Grundinfo_detail)
        Detail_update(self.Grundinfo_Summe0)
        Detail_update(self.Grundinfo_Summe1)
        Detail_update(self.Detail_detail)
        Detail_update(self.Detail_summe0)
        Detail_update(self.Detail_summe1)
        Detail_update(self.Detail_summe2)
        
    def luft_round(self,luft):
        zahl = luft % 5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Druckstufe_Berechnen(self):
        n = abs(int(self.Druckstufe.soll/10)) if abs(int(self.Druckstufe.soll/10)) < 6 else 5
        if self.Druckstufe.soll > 0:self.IGF_Legende = n*'+'
        else:self.IGF_Legende = n*'-'
          
    def Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self,zuluftmin):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll
        
        if self.zu_lab_min.soll > zuluftmin:
            zuluftmin = self.zu_lab_min.soll
        

        return self.luft_round(zuluftmin)

    def Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self,zuluftmin):

        zuluftmin =  zuluftmin - self.ueber_in_manuell.soll - self.ueber_in.soll
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll > zuluftmin:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        
        if zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll < self.zu_lab_min.soll:
            zuluftmin = self.zu_lab_min.soll -self.ueber_in_manuell.soll -self.ueber_in.soll

        if zuluftmin < 0:
            zuluftmin = 0


        return self.luft_round(zuluftmin)

    def Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self,zuluftmax):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax:
            zuluftmax = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll
        
        if self.zu_lab_max.soll > zuluftmax:
            zuluftmax = self.zu_lab_max.soll
        return zuluftmax

    def Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self,zuluftmax):
        zuluftmax =  zuluftmax - self.ueber_in_manuell.soll - self.ueber_in.soll
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll:
            zuluftmax = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        
        if zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll < self.zu_lab_max.soll:
            zuluftmax = self.zu_lab_max.soll -self.ueber_in_manuell.soll -self.ueber_in.soll
            
        if zuluftmax < 0:
            zuluftmax = 0

        return self.luft_round(zuluftmax)
    
    def Labmax_24h_Druckstufe_Pruefen(self,zuluftmax):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax :
            zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
        

        return self.luft_round(zuluftmax)

    def Nachtbetrieb_Berechnen(self):
        if self.bezugsname in ['Fläche',"Luftwechsel","Person","manuell"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll < 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = self.luft_round(self.T_NB_LW * self.volumen)
                
                    if self.ueber_Zu_Nacht:self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.tnb_zu.soll)
                    else:self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.tnb_zu.soll)
                    self.tnb_rab.soll = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
                    self.tnb_ab.soll = self.tnb_rab.soll + self.ab_lab_min.soll + self.ab_24h.soll #- self.Druckstufe.soll
                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0
                    self.tnb_rab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll
                if self.nb_dauer.soll < 0:
                    self.nb_dauer.soll += 24
                
                self.nb_dauer.soll -= self.tnb_dauer.soll

                self.nb_zu.soll = self.luft_round(self.NB_LW * self.volumen)
                if self.ueber_Zu_Nacht:self.nb_zu.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.nb_zu.soll)
                else:self.nb_zu.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.nb_zu.soll)
                self.nb_rab.soll = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
                self.nb_ab.soll = self.nb_rab.soll + self.ab_lab_min.soll + self.ab_24h.soll #- self.Druckstufe.soll
            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.nb_rab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_rab.soll = 0
                self.tnb_zu.soll = 0   

        elif self.bezugsname in ['nurZU_Fläche',"nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = 0-(self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll )#- self.Druckstufe.soll)
                    if self.tnb_zu.soll < 0:
                        self.logger.error("Raum {}: Achtung: Bitte Überströmung-Aus um {} m³/h erhöhen".format(self.Raumnr,0-self.tnb_zu.soll))
                        self.tnb_zu.soll = 0
                    # self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = 0
                    self.tnb_rab.soll = 0

                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0
                    self.tnb_rab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll
                if self.nb_dauer.soll < 0:
                    self.nb_dauer.soll += 24
                
                self.nb_dauer.soll -= self.tnb_dauer.soll

                self.nb_zu.soll = 0-(self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll )#- self.Druckstufe.soll)
                # self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = 0
                self.nb_rab.soll = 0
                if self.nb_zu.soll < 0:
                    self.logger.error("Raum {}: Achtung: Bitte Überströmung-Aus um {} m³/h erhöhen".format(self.Raumnr,0-self.nb_zu.soll))
                    self.nb_zu.soll = 0

            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.nb_rab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_rab.soll = 0
                self.tnb_zu.soll = 0   

        elif self.bezugsname in ['nurAB_Fläche',"nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = 0
                    # self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_rab.soll = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll# - self.Druckstufe.soll
                    if self.tnb_rab.soll < 0:
                        self.logger.error("Raum {}: Achtung: Bitte Überströmung-Ein um {} m³/h erhöhen".format(self.Raumnr,0-self.tnb_rab.soll))
                        self.tnb_rab.soll = 0
                    self.tnb_ab.soll = self.tnb_rab.soll + self.ab_lab_min.soll + self.ab_24h.soll# - self.Druckstufe.soll
                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0
                    self.tnb_rab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll
                if self.nb_dauer.soll < 0:
                    self.nb_dauer.soll += 24
                
                self.nb_dauer.soll -= self.tnb_dauer.soll
                
                self.nb_zu.soll = 0

                # self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_rab.soll = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
                if self.nb_rab.soll < 0:
                    self.logger.error("Raum {}: Achtung: Bitte Überströmung-Ein um {} m³/h erhöhen".format(self.Raumnr,0-self.tnb_rab.soll))
                    self.nb_rab.soll = 0
                self.nb_ab.soll = self.nb_rab.soll + self.ab_lab_min.soll + self.ab_24h.soll
            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.nb_rab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_rab.soll = 0
                self.tnb_zu.soll = 0   
        else:
            self.nb_dauer.soll = 0
            self.nb_zu.soll = 0
            self.nb_ab.soll = 0
            self.nb_rab.soll = 0
            self.tnb_dauer.soll = 0
            self.tnb_ab.soll = 0
            self.tnb_rab.soll = 0
            self.tnb_zu.soll = 0   
        
        self.update_soll_Nacht()
        
    def Tagesbetrieb_Berechnen(self):
        if self.flaeche == 0:
            return
        if self.bezugsname in ['nurZU_Fläche',"nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
            if (self.ab_lab_min.soll + self.ab_24h.soll> 0) or (self.ab_lab_max.soll + self.ab_24h.soll> 0):
                self.logger.error("Berechnungsprinzip von Raum {} ist Falsch. Der Raum ist nur über Überströmung ausströmt aber hat Laborabluft min: {}, max: {} m³/h und 24h-Abluft: {} m³/h".format(self.Raumnr,self.ab_lab_min.soll,self.ab_lab_max.soll,self.ab_24h.soll))
                return
            if self.bezugsname == "nurZU_Fläche":
                self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "nurZU_Luftwechsel":
                self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "nurZU_Person":
                self.zu_min.soll = self.luft_round(self.personen * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "nurZUMa":
                self.zu_min.soll = self.luft_round(float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            
            self.ab_min.soll = 0
            self.ab_max.soll = 0

            if self.ueber_Zu_Tag:
                self.zu_max.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.zu_min.soll)
                self.zu_min.soll = self.zu_max.soll
            else:
                self.zu_max.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.zu_min.soll)
                self.zu_min.soll = self.zu_max.soll


            self.angeabluft.soll = self.zu_min.soll
            self.angezuluft.soll = self.zu_min.soll
            
            abweichung = self.ueber_aus_manuell.soll + self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll #+ self.Druckstufe.soll
            if abweichung >= 0:
                self.zu_min.soll += abweichung
                self.zu_max.soll += abweichung

            else:
                self.hinweis = "Achtung: Bitte Überströmung-Aus um {} m³/h erhöhen".format(int(0 - abweichung))
                self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))

        elif self.bezugsname in ['nurAB_Fläche',"nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
            if self.bezugsname == "nurAB_Fläche":
                self.angezuluft.soll = self.luft_round(self.flaeche * float(self.faktor) * self.vol_faktorreduzierung_berechnen )
            elif self.bezugsname == "nurAB_Luftwechsel":
                self.angezuluft.soll = self.luft_round(self.volumen * float(self.faktor) * self.vol_faktorreduzierung_berechnen )
            elif self.bezugsname == "nurAB_Person":
                self.angezuluft.soll = self.luft_round(self.personen * float(self.faktor) * self.vol_faktorreduzierung_berechnen )
            elif self.bezugsname == "nurABMa":
                self.angezuluft.soll = self.luft_round(float(self.faktor) * self.vol_faktorreduzierung_berechnen)
                        
            self.angeabluft.soll = self.angezuluft.soll
            self.zu_min.soll = 0
            self.zu_max.soll = 0
            
            
            abweichung_max = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll #- self.Druckstufe.soll
            abweichung_min = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            abweichung = self.ueber_in.soll + self.ueber_in_manuell.soll - self.angezuluft.soll
            if abweichung_max >= 0:
                self.ab_min.soll = abweichung_min
                self.ab_max.soll = abweichung_max
                if abweichung < 0:
                    self.hinweis = "Achtung: Bitte Überströmung-Ein um {} m³/h erhöhen".format(int(0 - abweichung))
                    self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))
            else:
                self.ab_min.soll = 0
                self.ab_max.soll = 0

  
                if abweichung < 0:
                    self.hinweis = "Achtung: Bitte Überströmung-Ein um {} m³/h erhöhen".format(max(int(0 - abweichung),int(0 - abweichung_max)))
                    self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))
      
        elif self.bezugsname in ['Fläche',"Luftwechsel","Person","manuell"]:
            if self.bezugsname == "Fläche":
                self.zu_max.soll = self.luft_round(self.flaeche * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
                
            elif self.bezugsname == "Luftwechsel":
                self.zu_max.soll = self.luft_round(self.volumen * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "Person":
                self.zu_max.soll = self.luft_round(self.personen * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "manuell":
                self.zu_max.soll = self.luft_round(float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            
            self.zu_min.soll = self.zu_max.soll
            if self.ueber_Zu_Tag:
                self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.zu_max.soll)
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.zu_min.soll)
            else:
                self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.zu_max.soll)
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.zu_min.soll)
            
            self.ab_max.soll = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll #- self.Druckstufe.soll
            self.ab_min.soll = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll

            self.angeabluft.soll = self.zu_max.soll
            self.angezuluft.soll = self.zu_max.soll

            if self.ab_max.soll <= 0:
                self.zu_max.soll -= self.ab_max.soll
                self.ab_max.soll = 0
            if self.ab_min.soll <= 0:
                self.zu_max.soll -= self.ab_min.soll
                self.ab_min.soll = 0
                 
        elif self.bezugsname == 'keine':
            self.zu_min.soll = 0
            self.angezuluft.soll = 0
            self.angeabluft.soll = 0
            self.ab_min.soll = 0
            self.zu_max.soll = 0
            self.ab_max.soll = 0
            
        self.update_schacht()
        self.update_anlagen()
        self.update_soll_Tag()
  
    def get_element(self, param_name):
        param = self.elem.LookupParameter(param_name)
        if not param:
            self.logger.info("Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return self.get_value(param)
    
    def get_value(self,param):
       
        """gibt den gesuchten Wert ohne Einheit zurück"""
        if not param:return ''
        if param.StorageType.ToString() == 'ElementId':
            return param.AsValueString()
        elif param.StorageType.ToString() == 'Integer':
            value = param.AsInteger()
        elif param.StorageType.ToString() == 'Double':
            value = param.AsDouble()
        elif param.StorageType.ToString() == 'String':
            value = param.AsString()
            return value

        try:
            # in Revit 2020
            unit = param.DisplayUnitType
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            try:
                # in Revit 2021/2022
                unit = param.GetUnitTypeId()
                value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
            except:
                pass

        return value
    
    def werte_schreiben_Einbauteile(self):
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.elem.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)

        VSR_Params = ['IGF_RLT_MEP-Raum_VRs_00', 'IGF_RLT_MEP-Raum_VRs_01', 'IGF_RLT_MEP-Raum_VRs_02', 'IGF_RLT_MEP-Raum_VRs_03', 'IGF_RLT_MEP-Raum_VRs_04',
                      'IGF_RLT_MEP-Raum_VRs_05', 'IGF_RLT_MEP-Raum_VRs_06', 'IGF_RLT_MEP-Raum_VRs_07', 'IGF_RLT_MEP-Raum_VRs_08', 'IGF_RLT_MEP-Raum_VRs_09',
                      'IGF_RLT_MEP-Raum_VRs_10', 'IGF_RLT_MEP-Raum_VRs_11', 'IGF_RLT_MEP-Raum_VRs_12', 'IGF_RLT_MEP-Raum_VRs_13', 'IGF_RLT_MEP-Raum_VRs_14',
                      'IGF_RLT_MEP-Raum_VRs_15', 'IGF_RLT_MEP-Raum_VRs_16', 'IGF_RLT_MEP-Raum_VRs_17', 'IGF_RLT_MEP-Raum_VRs_18', 'IGF_RLT_MEP-Raum_VRs_19',
                      'IGF_RLT_MEP-Raum_VRs_20', 'IGF_RLT_MEP-Raum_VRs_21', 'IGF_RLT_MEP-Raum_VRs_22', 'IGF_RLT_MEP-Raum_VRs_23']

        Bauteile_Params = ['IGF_RLT_MEP-Raum_Einbau_00', 'IGF_RLT_MEP-Raum_Einbau_01', 'IGF_RLT_MEP-Raum_Einbau_02', 'IGF_RLT_MEP-Raum_Einbau_03', 'IGF_RLT_MEP-Raum_Einbau_04',
                           'IGF_RLT_MEP-Raum_Einbau_05', 'IGF_RLT_MEP-Raum_Einbau_06', 'IGF_RLT_MEP-Raum_Einbau_07', 'IGF_RLT_MEP-Raum_Einbau_08', 'IGF_RLT_MEP-Raum_Einbau_09',
                           'IGF_RLT_MEP-Raum_Einbau_10', 'IGF_RLT_MEP-Raum_Einbau_11', 'IGF_RLT_MEP-Raum_Einbau_12', 'IGF_RLT_MEP-Raum_Einbau_13', 'IGF_RLT_MEP-Raum_Einbau_14',
                           'IGF_RLT_MEP-Raum_Einbau_15', 'IGF_RLT_MEP-Raum_Einbau_16', 'IGF_RLT_MEP-Raum_Einbau_17', 'IGF_RLT_MEP-Raum_Einbau_18', 'IGF_RLT_MEP-Raum_Einbau_19',
                           'IGF_RLT_MEP-Raum_Einbau_20', 'IGF_RLT_MEP-Raum_Einbau_21', 'IGF_RLT_MEP-Raum_Einbau_22', 'IGF_RLT_MEP-Raum_Einbau_23', 'IGF_RLT_MEP-Raum_Einbau_24',
                           'IGF_RLT_MEP-Raum_Einbau_25', 'IGF_RLT_MEP-Raum_Einbau_26', 'IGF_RLT_MEP-Raum_Einbau_27', 'IGF_RLT_MEP-Raum_Einbau_28', 'IGF_RLT_MEP-Raum_Einbau_29',
                           'IGF_RLT_MEP-Raum_Einbau_30', 'IGF_RLT_MEP-Raum_Einbau_31']
        self._dict_analyse()

        Liste_VAR = self.Dict_VSR.keys()[:]
        Liste_VAR.sort()
        Liste_Bauteil = self.Dict_Einbauteile.keys()[:]
        Liste_Bauteil.sort()
        for n,param in enumerate(VSR_Params):
            if n < len(Liste_VAR):
                try:
                    anzahl = str(self.Dict_VSR[Liste_VAR[n]])
                    while ((len(anzahl)) < 2):
                        anzahl = '0' + anzahl
                    werte = anzahl + '_' + Liste_VAR[n]
                    wert_schreiben(param,werte)
                except Exception as e:
                    self.logger.error(e)
                    self.logger.error("Index: {}, Art: {}, Anzahl: {}".format(n,Liste_VAR[n],self.Dict_VSR[Liste_VAR[n]]))
                    self.logger.error('MEP Raum: {}, ElementId: {}'.format(self.Raumnr,self.elemid))
            else:
                wert_schreiben(param,'')
        
        for n,param in enumerate(Bauteile_Params):
            if n < len(Liste_Bauteil):
                try:
                    anzahl = str(self.Dict_Einbauteile[Liste_Bauteil[n]])
                    while ((len(anzahl)) < 2):
                        anzahl = '0' + anzahl
                    werte = anzahl + '_' + Liste_Bauteil[n]
                    wert_schreiben(param,werte)
                except Exception as e:
                    self.logger.error(e)
                    self.logger.error("Index: {}, Art: {}, Anzahl: {}".format(n,Liste_Bauteil[n],self.Dict_Einbauteile[Liste_VAR[n]]))
                    self.logger.error('MEP Raum: {}, ElementId: {}'.format(self.Raumnr,self.elemid))
            else:
                wert_schreiben(param,'')
    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""


        def wert_schreiben(param_name, wert):
            try:
                if not wert is None:
                #  logger.info(
                #      "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                    if self.elem.LookupParameter(param_name):
                        if self.elem.LookupParameter(param_name).IsReadOnly is True:
                            self.logger.error(self.elemid)
                            self.logger.error(param_name)
                            return
                        self.elem.LookupParameter(param_name).SetValueString(str(wert))
                    else:
                        print(param_name)
            except Exception as e:print(e,0)
        def wert_schreiben2(param_name, wert):
            try:
                if self.elem.LookupParameter(param_name):
                    if self.elem.LookupParameter(param_name).IsReadOnly is True:
                        self.logger.error(self.elemid)
                        self.logger.error(param_name)
                        return
                    self.elem.LookupParameter(param_name).Set(wert)
                else:
                    print(param_name)
            except Exception as e:print(e,1)  
        def wert_schreiben3(param_name, wert):
            '''für Schacht'''
            try:
                if self.elem.LookupParameter(param_name):
                    if wert.schachtindex != -1:
                        self.elem.LookupParameter(param_name).Set(wert.SchachtListe[wert.schachtindex])
                else:
                    print(param_name)
                  
            except Exception as e:pass#print(e)
        try:self.werte_schreiben_Einbauteile()
        except Exception as e:self.logger.error(e)
        wert_schreiben("Angegebener Zuluftstrom", self.angezuluft.soll)
        wert_schreiben("Angegebener Abluftluftstrom", self.angeabluft.soll)
        wert_schreiben("IGF_RLT_AbluftminRaum", self.ab_min.soll+self.ab_lab_min.soll+self.ab_24h.soll)
        wert_schreiben("IGF_RLT_AbluftmaxRaum", self.ab_max.soll+self.ab_lab_max.soll+self.ab_24h.soll)
        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zu_min.soll)
        wert_schreiben("IGF_RLT_ZuluftmaxRaum", self.zu_max.soll)
        # wert_schreiben("IGF_A_Lichte_Höhe", self.hoehe)
        # wert_schreiben("IGF_A_Volumen", self.volumen)

        wert_schreiben2("IGF_RLT_VolumenstromProName", self.bezugsname)
        wert_schreiben("IGF_RLT_VolumenstromProEinheit", self.einheit)
        wert_schreiben("IGF_RLT_VolumenstromProNummer", self.bezugsnummer)
        wert_schreiben("IGF_RLT_VolumenstromProFaktor", float(self.faktor))

        wert_schreiben2("IGF_RLT_Nachtbetrieb", self.nachtbetrieb)
        wert_schreiben("IGF_RLT_NachtbetriebLW", self.NB_LW)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.nb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.nb_ab.soll)
        wert_schreiben("IGF_RLT_NachtbetriebVon", self.nb_von.soll)
        wert_schreiben("IGF_RLT_NachtbetriebBis", self.nb_bis.soll)

        wert_schreiben2("IGF_RLT_TieferNachtbetrieb", self.tiefenachtbetrieb)
        wert_schreiben("IGF_RLT_TieferNachtbetriebLW", self.T_NB_LW)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer", self.tnb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tnb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tnb_ab.soll +self.ab_lab_max.soll+self.ab_24h.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebVon", self.tnb_von.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebBis", self.tnb_bis.soll)

        wert_schreiben("IGF_RLT_AbluftSumme24h", self.ab_24h.soll)
        # wert_schreiben("IGF_RLT_AbluftminRaumL24h", self.ab_24h.soll)    
        wert_schreiben("IGF_RLT_AbluftminSummeLabor", self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftminSummeLabor24h", self.ab_24h.soll + self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor24h", self.ab_24h.soll + self.ab_lab_max.soll)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor", self.ab_lab_max.soll)

        wert_schreiben("IGF_RLT_RaumDruckstufeEingabe", self.Druckstufe.soll)
        wert_schreiben2("IGF_RLT_RaumDruckstufeLegende", self.IGF_Legende)
        
        # wert_schreiben("IGF_RLT_AnlagenRaumAbluft", self.ab_min.soll+self.ab_lab_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaumZuluft", self.zu_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaum24hAbluft", self.ab_24h.soll)
        
        wert_schreiben("IGF_RLT_Luftmenge_RAB", self.ab_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_RZU", self.zu_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_Luftmenge_LAB", self.ab_lab_max.soll)
        # wert_schreiben("IGF_RLT_Luftmenge_min_TAB", self.tier_ab_min.soll)
        # wert_schreiben("IGF_RLT_Luftmenge_min_TZU", self.tier_zu_min.soll)
        
        wert_schreiben3('IGF_RLT_Schacht_RZU',self.rzu_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_RAB',self.rab_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TZU',self.tzu_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TAB',self.tab_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_24h',self._24h_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_LAB',self.lab_Schacht)

        for el in self.Anlagen_info:
            if el.name == 'RZU':
                wert_schreiben2('IGF_RLT_AnlagenNr_RZU',int(el.mep_nr))
            elif el.name == 'RAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_RAB',int(el.mep_nr))
            elif el.name == 'TZU':
                wert_schreiben2('IGF_RLT_AnlagenNr_TZU',int(el.mep_nr))
            elif el.name == 'TAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_TAB',int(el.mep_nr))
            elif el.name == 'LAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_LAB',int(el.mep_nr))
            elif el.name == '24h':
                wert_schreiben2('IGF_RLT_AnlagenNr_24h',int(el.mep_nr))
    
    def update_anlagen(self):
        for anlage in self.Anlagen_info:
            anlage._set_up()
        
    def update_schacht(self):
        for schacht in self.Schacht_info:
            schacht._set_up()
  
    def update_ist_von_auslass(self,auslass,zustand):
        if zustand == 'min':
            self.update_ist_min(auslass)
        elif zustand == 'max':
            self.update_ist_max(auslass)
        elif zustand == 'nacht':
            self.update_ist_nacht(auslass)

        elif zustand == 'tnacht':
            self.update_ist_tnacht(auslass)
        self.update_anlagen()

    def update_ist_min(self,auslass):
        if auslass in self.Liste_RZU:
            self.zu_min._set_up_ist()
            self.zu_minsum._set_up_ist()
            self.D_T_MIN_Rzu._set_up_ist()
            self.D_T_MIN_ZuS._set_up_ist()
            self.D_T_MIN_BLZ._set_up_ist()
            self.D_T_MIN_Aus._set_up_ist()
        elif auslass in self.Liste_RAB:
            self.ab_min._set_up_ist()
            self.ab_minsum._set_up_ist()
            self.D_T_MIN_Rab._set_up_ist()
            self.D_T_MIN_AbS._set_up_ist()
            self.D_T_MIN_BLZ._set_up_ist()
            self.D_T_MIN_Aus._set_up_ist()
        elif auslass in self.Liste_LAB:
            self.ab_lab_min._set_up_ist()
            self.ab_minsum._set_up_ist()
            self.D_T_MIN_Lab._set_up_ist()
            self.D_T_MIN_AbS._set_up_ist()
            self.D_T_MIN_BLZ._set_up_ist()
            self.D_T_MIN_Aus._set_up_ist()
        elif auslass in self.Liste_H24:
            self.ab_24h._set_up_ist()
            self.ab_minsum._set_up_ist()
            self.D_T_MIN_24h._set_up_ist()
            self.D_T_MIN_AbS._set_up_ist()
            self.D_T_MIN_BLZ._set_up_ist()
            self.D_T_MIN_Aus._set_up_ist()
        elif auslass in self.Liste_TZU:
            self.tier_zu_min._set_up_ist()
            self.D_T_MIN_TZu._set_up_ist()
        elif auslass in self.Liste_TAB:
            self.tier_ab_min._set_up_ist()
            self.D_T_MIN_TAb._set_up_ist()
        
    def update_ist_max(self,auslass):
        if auslass in self.Liste_RZU:
            self.zu_max._set_up_ist()
            self.zu_maxsum._set_up_ist()
            self.D_T_MAX_Rzu._set_up_ist()
            self.D_T_MAX_ZuS._set_up_ist()
            self.D_T_MAX_BLZ._set_up_ist()
            self.D_T_MAX_Aus._set_up_ist()
        elif auslass in self.Liste_RAB:
            self.ab_max._set_up_ist()
            self.ab_maxsum._set_up_ist()
            self.D_T_MAX_Rab._set_up_ist()
            self.D_T_MAX_AbS._set_up_ist()
            self.D_T_MAX_BLZ._set_up_ist()
            self.D_T_MAX_Aus._set_up_ist()
        elif auslass in self.Liste_LAB:
            self.ab_lab_max._set_up_ist()
            self.ab_maxsum._set_up_ist()
            self.D_T_MAX_Lab._set_up_ist()
            self.D_T_MAX_AbS._set_up_ist()
            self.D_T_MAX_BLZ._set_up_ist()
            self.D_T_MAX_Aus._set_up_ist()
        elif auslass in self.Liste_H24:
            self.ab_24h._set_up_ist()
            self.ab_maxsum._set_up_ist()
            self.D_T_MAX_24h._set_up_ist()
            self.D_T_MAX_AbS._set_up_ist()
            self.D_T_MAX_BLZ._set_up_ist()
            self.D_T_MAX_Aus._set_up_ist()
        elif auslass in self.Liste_TZU:
            self.tier_zu_max._set_up_ist()
            self.D_T_MAX_TZu._set_up_ist()
        elif auslass in self.Liste_TAB:
            self.tier_ab_max._set_up_ist()
            self.D_T_MAX_TAb._set_up_ist()
    
    def update_ist_nacht(self,auslass):
        if auslass in self.Liste_RZU:
            self.nb_zu._set_up_ist()
            self.D_T_NB_Rzu._set_up_ist()
            self.D_T_NB_ZuS._set_up_ist()
            self.D_T_NB_BLZ._set_up_ist()
            self.D_T_NB_Aus._set_up_ist()
        elif auslass in self.Liste_RAB:
            self.nb_rab._set_up_ist()
            self.nb_ab._set_up_ist()
            self.D_T_NB_Rab._set_up_ist()
            self.D_T_NB_AbS._set_up_ist()
            self.D_T_NB_BLZ._set_up_ist()
            self.D_T_NB_Aus._set_up_ist()
        elif auslass in self.Liste_LAB:
            self.nb_lab._set_up_ist()
            self.nb_ab._set_up_ist()
            self.D_T_NB_Lab._set_up_ist()
            self.D_T_NB_AbS._set_up_ist()
            self.D_T_NB_BLZ._set_up_ist()
            self.D_T_NB_Aus._set_up_ist()
        elif auslass in self.Liste_H24:
            self.nb_24h._set_up_ist()
            self.nb_ab._set_up_ist()
            self.D_T_NB_24h._set_up_ist()
            self.D_T_NB_AbS._set_up_ist()
            self.D_T_NB_BLZ._set_up_ist()
            self.D_T_NB_Aus._set_up_ist()
    
    def update_ist_bei_verteilen(self):
        self.ab_min._set_up_ist()
        self.ab_max._set_up_ist()
        self.zu_max._set_up_ist()
        self.zu_max._set_up_ist()
        self.nb_zu._set_up_ist()
        self.tnb_zu._set_up_ist()
        self.tnb_rab._set_up_ist()
        self.nb_rab._set_up_ist()
        self.zu_minsum._set_up_ist()
        self.zu_maxsum._set_up_ist()
        self.ab_minsum._set_up_ist()
        self.ab_maxsum._set_up_ist()
        self.nb_ab._set_up_ist()
        self.tnb_ab._set_up_ist()
        self.D_T_MIN_Rzu._set_up_ist()
        self.D_T_MIN_Rab._set_up_ist()
        self.D_T_MIN_ZuS._set_up_ist()
        self.D_T_MIN_AbS._set_up_ist()
        self.D_T_MIN_BLZ._set_up_ist()
        self.D_T_MIN_Aus._set_up_ist()
        self.D_T_MAX_Rzu._set_up_ist()
        self.D_T_MAX_Rab._set_up_ist()
        self.D_T_MAX_ZuS._set_up_ist()
        self.D_T_MAX_AbS._set_up_ist()
        self.D_T_MAX_BLZ._set_up_ist()
        self.D_T_MAX_Aus._set_up_ist()
        self.D_T_NB_Rzu._set_up_ist()
        self.D_T_NB_Rab._set_up_ist()
        self.D_T_NB_ZuS._set_up_ist()
        self.D_T_NB_AbS._set_up_ist()
        self.D_T_NB_BLZ._set_up_ist()
        self.D_T_NB_Aus._set_up_ist()
        self.D_T_TNB_Rzu._set_up_ist()
        self.D_T_TNB_Rab._set_up_ist()
        self.D_T_TNB_ZuS._set_up_ist()
        self.D_T_TNB_AbS._set_up_ist()
        self.D_T_TNB_BLZ._set_up_ist()
        self.D_T_TNB_Aus._set_up_ist()
        self.update_anlagen()
    
    def update_ist_tnacht(self,auslass):
        if auslass in self.Liste_RZU:
            self.tnb_zu._set_up_ist()
            self.D_T_TNB_Rzu._set_up_ist()
            self.D_T_TNB_ZuS._set_up_ist()
            self.D_T_TNB_BLZ._set_up_ist()
            self.D_T_TNB_Aus._set_up_ist()
        elif auslass in self.Liste_RAB:
            self.tnb_rab._set_up_ist()
            self.tnb_ab._set_up_ist()
            self.D_T_TNB_Rab._set_up_ist()
            self.D_T_TNB_AbS._set_up_ist()
            self.D_T_TNB_BLZ._set_up_ist()
            self.D_T_TNB_Aus._set_up_ist()
        elif auslass in self.Liste_LAB:
            self.tnb_lab._set_up_ist()
            self.tnb_ab._set_up_ist()
            self.D_T_TNB_Lab._set_up_ist()
            self.D_T_TNB_AbS._set_up_ist()
            self.D_T_TNB_BLZ._set_up_ist()
            self.D_T_TNB_Aus._set_up_ist()
        elif auslass in self.Liste_H24:
            self.nb_24h._set_up_ist()
            self.nb_ab._set_up_ist()
            self.D_T_TNB_24h._set_up_ist()
            self.D_T_TNB_AbS._set_up_ist()
            self.D_T_TNB_BLZ._set_up_ist()
            self.D_T_TNB_Aus._set_up_ist()

    def _vsr_bearbeiten(self):
        for el in self.list_vsr:
            if el.IsIris == True:
                self.list_vsr2.Add(el)
            elif el.IsHaupt == True:
                self.list_vsr0.Add(el) 
            else:
                self.list_vsr1.Add(el)

        if self.list_vsr0.Count > 0:
            for el in self.list_vsr0:
                for auslass in el.Auslass:
                    iris = auslass.IRIS_Class
                    vsr = auslass.VSR_Class
                    if iris is not None:
                        if el.List_Iris.Contains(iris) == False:
                            el.List_Iris.Add(iris)
                            iris.List_Haupt.Add(el)
                    if vsr is not None:
                        if el.List_VSR.Contains(vsr) == False:
                            el.List_VSR.Add(vsr)
                            vsr.List_Haupt.Add(el)
                        if iris is not None:
                            if vsr.List_Iris.Contains(iris) == False:
                                vsr.List_Iris.Add(iris)
                                iris.List_VSR.Add(el)

        else:
            if self.list_vsr1.Count > 0:
                for el in self.list_vsr1:
                    for auslass in el.Auslass:
                        iris = auslass.IRIS_Class
                        if iris is not None:
                            if el.List_Iris.Contains(iris) == False:
                                el.List_Iris.Add(iris)
                                iris.List_VSR.Add(el)

        if self.list_vsr0.Count > 0:
            for mainvsr in self.list_vsr0:
                for subvsr in mainvsr.List_VSR:
                    if self.list_vsr1.Contains(subvsr):
                        if subvsr.slavevon != '-':
                            if subvsr.slavevon != mainvsr.vsrid:
                                logger.error('VSR {} {} hat zwei Haupt VSR {}, {}. Bitte überprüfen.'.format(subvsr.elemid,subvsr.vsrid,mainvsr.vsrid,subvsr.slavevon))
                        subvsr.slavevon = mainvsr.vsrid
                        if subvsr.vsrid.find('-->') == -1:
                            subvsr.vsrid = '--> ' + subvsr.vsrid

    def _auslass_bearbeiten(self):
        if self.elemid.ToString() in self.DICT_MEP_AUSLASS.keys():
             for art in sorted(self.DICT_MEP_AUSLASS[self.elemid.ToString()].keys()):
                for fam in sorted(self.DICT_MEP_AUSLASS[self.elemid.ToString()][art].keys()):
                    for terminal in self.DICT_MEP_AUSLASS[self.elemid.ToString()][art][fam]:
                        self.list_auslass.Add(terminal)

    def _auslass_analyse(self):
        Liste_RZU = []
        Liste_RAB = []
        Liste_TZU = []
        Liste_TAB = []
        Liste_H24 = []
        Liste_LAB = []

        Liste_UIN = self.list_ueber["Ein"]
        Liste_UAU = self.list_ueber["Aus"]

        Dict_RZU = {}
        Dict_RAB = {}


        for auslass in self.list_auslass:
            if auslass.art == '24h':Liste_H24.append(auslass)
            elif auslass.art == 'LAB':Liste_LAB.append(auslass)
            elif auslass.art == 'RZU':
                Liste_RZU.append(auslass)
                if auslass.familyandtyp not in Dict_RZU.keys():
                    Dict_RZU[auslass.familyandtyp] = []
                Dict_RZU[auslass.familyandtyp].append(auslass)

            elif auslass.art == 'RAB':
                Liste_RAB.append(auslass)
                if auslass.familyandtyp not in Dict_RAB.keys():
                    Dict_RAB[auslass.familyandtyp] = []
                Dict_RAB[auslass.familyandtyp].append(auslass)
            elif auslass.art == 'TZU':Liste_TZU.append(auslass)
            elif auslass.art == 'TAB':Liste_TAB.append(auslass)
        Liste_ZU = Liste_RZU[:]
        Liste_ZU.extend(Liste_UIN)
        Liste_ZU.append(self.ueber_in_m_class)
        Liste_AB = Liste_RAB[:]
        Liste_AB.append(self.ueber_aus_m_class)
        Liste_AB.extend(Liste_LAB)
        Liste_AB.extend(Liste_H24)
        Liste_AB.extend(Liste_UAU)

        self.Liste_UAU = Liste_UAU
        self.Liste_UIN = Liste_UIN
        self.Liste_LAB = Liste_LAB
        self.Liste_H24 = Liste_H24
        self.Liste_TAB = Liste_TAB
        self.Liste_TZU = Liste_TZU
        self.Liste_RAB = Liste_RAB
        self.Liste_RZU = Liste_RZU
        self.Liste_AB = Liste_AB
        self.Liste_ZU = Liste_ZU
        self.Dict_RAB = Dict_RAB
        self.Dict_RZU = Dict_RZU

    def _uebersicht_bearbeiten(self):
        self.Uebersicht.Add(self.angezuluft)
        self.Uebersicht.Add(self.angeabluft)
        self.Uebersicht.Add(self.ab_24h)
        self.Uebersicht.Add(self.ab_minsum)
        self.Uebersicht.Add(self.ab_min)
        self.Uebersicht.Add(self.ab_lab_min)
        self.Uebersicht.Add(self.zu_minsum)
        self.Uebersicht.Add(self.zu_min)
        self.Uebersicht.Add(self.ab_maxsum)
        self.Uebersicht.Add(self.ab_max)
        self.Uebersicht.Add(self.ab_lab_max)
        self.Uebersicht.Add(self.zu_maxsum)
        self.Uebersicht.Add(self.zu_max)
        self.Uebersicht.Add(self.nb_von)
        self.Uebersicht.Add(self.nb_bis)
        self.Uebersicht.Add(self.nb_dauer)
        self.Uebersicht.Add(self.nb_zu)
        self.Uebersicht.Add(self.nb_ab)
        self.Uebersicht.Add(self.tnb_von)
        self.Uebersicht.Add(self.tnb_bis)
        self.Uebersicht.Add(self.tnb_dauer)
        self.Uebersicht.Add(self.tnb_zu)
        self.Uebersicht.Add(self.tnb_ab)
        self.Uebersicht.Add(self.tier_zu_min)
        self.Uebersicht.Add(self.tier_ab_min)
        self.Uebersicht.Add(self.tier_zu_max)
        self.Uebersicht.Add(self.tier_ab_max)
        self.Uebersicht.Add(self.ueber_sum)
        self.Uebersicht.Add(self.ueber_in)
        self.Uebersicht.Add(self.ueber_aus)
        self.Uebersicht.Add(self.ueber_in_manuell)
        self.Uebersicht.Add(self.ueber_aus_manuell)
        self.Uebersicht.Add(self.Druckstufe)

    def _ueberstrom_bearbeiten(self):
        if self.nachtbetrieb:
            for el in self.Liste_UIN:
                el.Luftmengennacht = el.menge
            for el in self.Liste_UAU:
                el.Luftmengennacht = el.menge
            self.ueber_aus_m_class.Luftmengennacht = self.ueber_aus_m_class.menge
            self.ueber_in_m_class.Luftmengennacht = self.ueber_in_m_class.menge
            self.nb_ueber_in.soll = self.ueber_in.soll
            self.nb_ueber_in_m.soll = self.ueber_in_manuell.soll
            self.nb_ueber_aus.soll = self.ueber_aus.soll
            self.nb_ueber_aus_m.soll = self.ueber_aus_manuell.soll

            if self.tiefenachtbetrieb:
                for el in self.Liste_UIN:
                    el.Luftmengentnacht = el.menge
                for el in self.Liste_UAU:
                    el.Luftmengentnacht = el.menge
                self.ueber_aus_m_class.Luftmengentnacht = self.ueber_aus_m_class.menge
                self.ueber_in_m_class.Luftmengentnacht = self.ueber_in_m_class.menge
                self.tnb_ueber_in.soll = self.ueber_in.soll
                self.tnb_ueber_in_m.soll = self.ueber_in_manuell.soll
                self.tnb_ueber_aus.soll = self.ueber_aus.soll
                self.tnb_ueber_aus_m.soll = self.ueber_aus_manuell.soll
            else:
                for el in self.Liste_UIN:
                    el.Luftmengentnacht = 0
                for el in self.Liste_UAU:
                    el.Luftmengentnacht = 0
                self.ueber_aus_m_class.Luftmengentnacht = 0
                self.ueber_in_m_class.Luftmengentnacht = 0
                self.tnb_ueber_in.soll = 0
                self.tnb_ueber_in_m.soll = 0
                self.tnb_ueber_aus.soll = 0
                self.tnb_ueber_aus_m.soll = 0
        else:
            for el in self.Liste_UIN:
                el.Luftmengennacht = 0
                el.Luftmengentnacht = 0
            for el in self.Liste_UAU:
                el.Luftmengennacht = 0
                el.Luftmengentnacht = 0

            self.ueber_aus_m_class.Luftmengennacht = 0
            self.ueber_in_m_class.Luftmengennacht = 0
            self.ueber_aus_m_class.Luftmengentnacht = 0
            self.ueber_in_m_class.Luftmengentnacht = 0
            self.tnb_ueber_in.soll = 0
            self.tnb_ueber_in_m.soll = 0
            self.tnb_ueber_aus.soll = 0
            self.tnb_ueber_aus_m.soll = 0
            self.nb_ueber_in.soll = 0
            self.nb_ueber_in_m.soll = 0
            self.nb_ueber_aus.soll = 0
            self.nb_ueber_aus_m.soll = 0

    def _lab_bearbeiten(self):
        if self.nachtbetrieb:
            self.nb_lab.soll = self.ab_lab_min.soll
            if self.tiefenachtbetrieb:
                self.tnb_lab.soll = self.ab_lab_min.soll
            else:
                self.tnb_lab.soll = 0
        else:
            self.tnb_lab.soll = 0
            self.nb_lab.soll = 0
    
    def _24h_bearbeiten(self):
        if self.nachtbetrieb:
            self.nb_24h.soll = self.ab_24h.soll
            if self.tiefenachtbetrieb:
                self.tnb_24h.soll = self.ab_24h.soll
            else:
                self.tnb_24h.soll = 0
        else:
            self.tnb_24h.soll = 0
            self.nb_24h.soll = 0

    def _druckstufe_bearbeiten(self):
        if self.nachtbetrieb:
            self.nb_druckstufe.soll = self.nb_druckstufe.ist = self.Druckstufe.soll
            if self.tiefenachtbetrieb:
                self.tnb_druckstufe.soll = self.tnb_druckstufe.ist = self.Druckstufe.soll
                
            else:
                self.tnb_druckstufe.soll = self.tnb_druckstufe.ist = 0
                
        else:
            self.nb_druckstufe.soll = self.nb_druckstufe.ist = 0
            self.tnb_druckstufe.soll = self.tnb_druckstufe.ist = 0

    def _ueberstrom_manuel_bearbeiten(self):
        if self.elem.Number in self.Dict_Ueber_Manuell.keys():
            self.ueber_aus_manuell.soll = self.Dict_Ueber_Manuell[self.elem.Number]
            self.ueber_aus_m_class.menge = self.Dict_Ueber_Manuell[self.elem.Number]
            self.ueber_aus_m_class.Luftmengenmin = self.Dict_Ueber_Manuell[self.elem.Number]
            self.ueber_aus_m_class.Luftmengenmax = self.Dict_Ueber_Manuell[self.elem.Number]

    def _detail_bearbeiten(self):
        self.Detail_Min.Add(self.D_T_MIN_ZuS)
        self.Detail_Min.Add(self.D_T_MIN_Rzu)
        self.Detail_Min.Add(self.D_T_MIN_ZuUe)
        self.Detail_Min.Add(self.D_T_MIN_ZuUeM)
        self.Detail_Min.Add(self.LEER)
        self.Detail_Min.Add(self.D_T_MIN_AbS)
        self.Detail_Min.Add(self.D_T_MIN_Rab)
        self.Detail_Min.Add(self.D_T_MIN_AbUe)
        self.Detail_Min.Add(self.D_T_MIN_AbUeM)
        self.Detail_Min.Add(self.D_T_MIN_Lab)
        self.Detail_Min.Add(self.D_T_MIN_24h)
        self.Detail_Min.Add(self.LEER)
        self.Detail_Min.Add(self.D_T_MIN_BLZ)
        self.Detail_Min.Add(self.D_T_Druckstufe)
        self.Detail_Min.Add(self.LEER)
        self.Detail_Min.Add(self.D_T_MIN_Aus)
        self.Detail_Min.Add(self.LEER)
        self.Detail_Min.Add(self.D_T_MIN_Tier)
        self.Detail_Min.Add(self.D_T_MIN_TZu)
        self.Detail_Min.Add(self.D_T_MIN_TAb)

        self.Detail_Max.Add(self.D_T_MAX_ZuS)
        self.Detail_Max.Add(self.D_T_MAX_Rzu)
        self.Detail_Max.Add(self.D_T_MAX_ZuUe)
        self.Detail_Max.Add(self.D_T_MAX_ZuUeM)
        self.Detail_Max.Add(self.LEER)
        self.Detail_Max.Add(self.D_T_MAX_AbS)
        self.Detail_Max.Add(self.D_T_MAX_Rab)
        self.Detail_Max.Add(self.D_T_MAX_AbUe)
        self.Detail_Max.Add(self.D_T_MAX_AbUeM)
        self.Detail_Max.Add(self.D_T_MAX_Lab)
        self.Detail_Max.Add(self.D_T_MAX_24h)
        self.Detail_Max.Add(self.LEER)
        self.Detail_Max.Add(self.D_T_MAX_BLZ)
        self.Detail_Max.Add(self.D_T_Druckstufe)
        self.Detail_Max.Add(self.LEER)
        self.Detail_Max.Add(self.D_T_MAX_Aus)
        self.Detail_Max.Add(self.LEER)
        self.Detail_Max.Add(self.D_T_MIN_Tier)
        self.Detail_Max.Add(self.D_T_MAX_TZu)
        self.Detail_Max.Add(self.D_T_MAX_TAb)

        self.Detail_Nacht.Add(self.D_T_NB_ZuS)
        self.Detail_Nacht.Add(self.D_T_NB_Rzu)
        self.Detail_Nacht.Add(self.D_T_NB_ZuUe)
        self.Detail_Nacht.Add(self.D_T_NB_ZuUeM)
        self.Detail_Nacht.Add(self.LEER)
        self.Detail_Nacht.Add(self.D_T_NB_AbS)
        self.Detail_Nacht.Add(self.D_T_NB_Rab)
        self.Detail_Nacht.Add(self.D_T_NB_AbUe)
        self.Detail_Nacht.Add(self.D_T_NB_AbUeM)
        self.Detail_Nacht.Add(self.D_T_NB_Lab)
        self.Detail_Nacht.Add(self.D_T_NB_24h)
        self.Detail_Nacht.Add(self.LEER)
        self.Detail_Nacht.Add(self.D_T_NB_BLZ)
        self.Detail_Nacht.Add(self.D_N_Druckstufe)
        self.Detail_Nacht.Add(self.LEER)
        self.Detail_Nacht.Add(self.D_T_NB_Aus)

        self.Detail_Tnacht.Add(self.D_T_TNB_ZuS)
        self.Detail_Tnacht.Add(self.D_T_TNB_Rzu)
        self.Detail_Tnacht.Add(self.D_T_TNB_ZuUe)
        self.Detail_Tnacht.Add(self.D_T_TNB_ZuUeM)
        self.Detail_Tnacht.Add(self.LEER)
        self.Detail_Tnacht.Add(self.D_T_TNB_AbS)
        self.Detail_Tnacht.Add(self.D_T_TNB_Rab)
        self.Detail_Tnacht.Add(self.D_T_TNB_AbUe)
        self.Detail_Tnacht.Add(self.D_T_TNB_AbUeM)
        self.Detail_Tnacht.Add(self.D_T_TNB_Lab)
        self.Detail_Tnacht.Add(self.D_T_TNB_24h)
        self.Detail_Tnacht.Add(self.LEER)
        self.Detail_Tnacht.Add(self.D_T_TNB_BLZ)
        self.Detail_Tnacht.Add(self.D_TN_Druckstufe)
        self.Detail_Tnacht.Add(self.LEER)
        self.Detail_Tnacht.Add(self.D_T_TNB_Aus)

    def _schacht_bearbeiten(self):
        self.Schacht_info.Add(self.rzu_Schacht)
        self.Schacht_info.Add(self.rab_Schacht)
        self.Schacht_info.Add(self.tzu_Schacht)
        self.Schacht_info.Add(self.tab_Schacht)
        self.Schacht_info.Add(self._24h_Schacht)
        self.Schacht_info.Add(self.lab_Schacht)
    
    def _anlagen_bearbeiten(self):
        Dict = {}
        for el in self.list_auslass:
            if not el.art in Dict.keys():
                Dict[el.art] = {}
            if not el.AnlNr in Dict[el.art].keys():
                Dict[el.art][el.AnlNr] = []
            
            Dict[el.art][el.AnlNr].append(el)
        
        if 'RZU' in Dict.keys():
            for anl in sorted(Dict['RZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                    self.zu_max,anl,Dict['RZU'][anl]))
        else:
            self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                self.zu_max,'',''))
        if 'RAB' in Dict.keys():
            for anl in sorted(Dict['RAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                    self.ab_max,anl,Dict['RAB'][anl]))
        else:
            self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                self.ab_max,'',''))
        if 'TZU' in Dict.keys():
            for anl in sorted(Dict['TZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                    self.tier_zu_max,anl,Dict['TZU'][anl]))
        else:
            self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                self.tier_zu_max,'',''))
        
        if 'TAB' in Dict.keys():
            for anl in sorted(Dict['TAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                    self.tier_ab_max,anl,Dict['TAB'][anl]))
        else:
            self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                self.tier_ab_max,'',''))
        
        if '24h' in Dict.keys():
            for anl in sorted(Dict['24h'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                    self.ab_24h,anl,Dict['24h'][anl]))
        else:
            self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                self.ab_24h,'',''))
           
        if 'LAB' in Dict.keys():
            for anl in sorted(Dict['LAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                    self.ab_lab_max,anl,Dict['LAB'][anl]))  
        else:
            self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                self.ab_lab_max,'',''))

    def _liste_bearbeiten(self):
        self.Grundinfo_detail = [self.ab_24h,self.ab_min,self.ab_lab_min,self.zu_min,self.ab_max,self.ab_lab_max,self.zu_max,self.angezuluft,self.angeabluft,\
                                 self.nb_bis,self.nb_dauer,self.nb_von,self.tnb_bis,self.tnb_dauer,self.tnb_von,self.ueber_in,self.ueber_in_manuell,\
                                 self.tier_ab_min,self.tier_ab_max,self.tier_zu_max,self.tier_zu_min,self.ueber_aus,self.ueber_aus_manuell,\
                                 self.nb_ueber_in,self.nb_ueber_in_m,self.nb_rab,self.nb_ueber_aus,self.nb_ueber_aus_m,self.nb_lab,self.nb_24h,self.nb_druckstufe,\
                                 self.tnb_ueber_in,self.tnb_ueber_in_m,self.tnb_rab,self.tnb_ueber_aus,self.tnb_ueber_aus_m,self.tnb_lab,self.tnb_24h,self.tnb_druckstufe]
        self.Grundinfo_Summe0 = [self.nb_ab,self.nb_zu,self.tnb_ab,self.tnb_zu]
        self.Grundinfo_Summe1 = [self.ab_minsum,self.zu_minsum,self.ueber_sum,self.Druckstufe,self.ab_maxsum,self.zu_maxsum]

        self.Detail_detail = [self.D_T_MIN_Rzu,self.D_T_MIN_ZuUe,self.D_T_MIN_ZuUeM,self.D_T_MIN_Rab,self.D_T_MIN_Lab,self.D_T_MIN_24h,\
                              self.D_T_MIN_AbUe,self.D_T_MIN_AbUeM,self.D_T_Druckstufe,self.D_T_MIN_TZu,self.D_T_MIN_TAb,self.D_T_MIN_Tier,\
                              self.D_T_MAX_Rzu,self.D_T_MAX_ZuUe,self.D_T_MAX_ZuUeM,self.D_T_MAX_Rab,self.D_T_MAX_Lab,self.D_T_MAX_24h,\
                              self.D_T_MAX_AbUe,self.D_T_MAX_AbUeM,self.D_T_MAX_TZu,self.D_T_MAX_TAb,\
                              self.D_T_NB_Rzu,self.D_T_NB_ZuUe,self.D_T_NB_ZuUeM,self.D_T_NB_Rab,self.D_T_NB_Lab,self.D_T_NB_24h,\
                              self.D_T_NB_AbUe,self.D_T_NB_AbUeM,self.D_N_Druckstufe,\
                              self.D_T_TNB_Rzu,self.D_T_TNB_ZuUe,self.D_T_TNB_ZuUeM,self.D_T_TNB_Rab,self.D_T_TNB_Lab,self.D_T_TNB_24h,\
                              self.D_T_TNB_AbUe,self.D_T_TNB_AbUeM,self.D_TN_Druckstufe]
        self.Detail_summe0 = [self.D_T_MIN_ZuS,self.D_T_MIN_AbS,\
                              self.D_T_MAX_ZuS,self.D_T_MAX_AbS,\
                              self.D_T_NB_ZuS,self.D_T_NB_AbS,\
                              self.D_T_TNB_ZuS,self.D_T_TNB_AbS]
        self.Detail_summe1 = [self.D_T_MIN_BLZ,\
                              self.D_T_MAX_BLZ,\
                              self.D_T_NB_BLZ,\
                              self.D_T_TNB_BLZ]
        self.Detail_summe2 = [self.D_T_MIN_Aus,\
                              self.D_T_MAX_Aus,\
                              self.D_T_NB_Aus,\
                              self.D_T_TNB_Aus]
    
    def _unrelevant_bearbeiten(self):
        if self.elemid.ToString() in self.DICT_MEP_UN_AUNLASS.keys():
            liste = self.DICT_MEP_UN_AUNLASS[self.elemid.ToString()]
            _dict = {}
            for el in liste:
                if el.familyandtyp not in _dict.keys():
                    _dict[el.familyandtyp] = []
                _dict[el.familyandtyp].append(el)
            for key in sorted(_dict.keys()):
                for elem in sorted(_dict[key]):
                    self.Liste_RaumluftUnrelevant.Add(elem) 
    
    def _dict_analyse(self):
        dict_vsr = {}
        dict_einbauteile = {}
        for el in self.Einbauteile:
            el.get_Art()
            el.Luftmengenermitteln()

            if not el.Text in dict_einbauteile.keys():
                dict_einbauteile[el.Text] = 1
            else:
                dict_einbauteile[el.Text] += 1

        for el in self.VSR:
            if not el.Text in dict_vsr.keys():
                dict_vsr[el.Text] = 1
            else:
                dict_vsr[el.Text] += 1
        self.Dict_Einbauteile = dict_einbauteile
        self.Dict_VSR = dict_vsr


