# coding: utf8
import System
from System import Windows
from System.Collections.ObjectModel import ObservableCollection

from Luftbilanz_Config import config,doc,DB,uidoc
from Luftbilanz_Herstellerdaten import Liste_Fabrikat,DICT_DatenBank,Liste_Datenbank,Liste_Datenbank1
from Luftbilanz_Forms import Familien,forms,RED,BLACK,HIDDEN,VISIBLE
from Luftbilanz_Uebersicht import MEPRaum_Uebersicht
from Luftbilanz_Klasse import *
from clr import GetClrType
from IGF_lib import get_value
from IGF_log import getlog

__title__ = "Raumluftbilanz_NA"
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

[2023.02.27]
Version: 1.2
"""
__authors__ = "Menghui Zhang"

# try:
#     getlog(__title__)
# except:
#     pass

DICT_MEP_NUMMER_NAME = {} # 100.103 : 100.103 - Büro
DICT_MEP_AUSLASS = {} ## MEPID: OB(Auslässe)
DICT_MEP_VSR = {} ## MEPID: [VSRID]
DICT_MEP_UEBERSTROM = {} ## MEPID: {'Ein': ..., 'Aus':...}

ELEMID_UEBER = [] ## ElememntId Überstrom
LISTE_SCHACHT = []

DICT_VSR_MEP_NUR_NUMMER = {} ## VSRID: [MEPNr,MEPNr]
DICT_MEP_UN_AUNLASS = {} #Raumluftbilanzunrelevante Auslässe, Raumid: [Auslass] 
DICT_VSR_MEP = {} ## VSRID: [MEPNr-Name,MEPNr-Name]
DICT_VSR_VSRLISTE = {} ## VSRID: [VSRID,VSRID]
DICT_VSR_AUSLASS = {} ## VSRID: OB(Auslässe)
DICT_VSR = {} ## VSRID: Class VSR
LISTE_VSR = []
LISTE_IRIS = []
DICT_MEP_EINBAUTEILE = {}
DICT_EINBAUTEILE_AUSLASS = {}
DICT_MEP_VSRKLASSE = {}

views = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.View3D)).ToElements()
for el in views:
    if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
        uidoc.ActiveView = el
        break

active_view = uidoc.ActiveView
if active_view.Name != 'Berechnungsmodell_L_KG4xx_IGF':
    logger.error('die Ansicht "Berechnungsmodell_L_KG4xx_IGF "nicht gefunden!')
    script.exit()

IS_AUSLASS = ObservableCollection[Itemtemplate]()
IS_AUSLASS_D = ObservableCollection[Itemtemplate]()
IS_AUSLASS_LAB = ObservableCollection[Itemtemplate]()
IS_AUSLASS_24H = ObservableCollection[Itemtemplate]()
IS_VSR = ObservableCollection[Itemtemplate]()
IS_DOSSEL = ObservableCollection[Itemtemplate]()

def get_IS():
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Liste_Auslass = []
    Liste_Auslass_Lab = []
    Liste_Zubehoer = []
    for el in Families:
        FamName = el.Name
        if el.FamilyCategoryId.IntegerValue == -2008013:
            if FamName not in Liste_Auslass:
                Liste_Auslass.append(FamName)
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Liste_Auslass_Lab:
                    Liste_Auslass_Lab.append(typname)
        elif el.FamilyCategoryId.IntegerValue == -2008016:
            if FamName not in Liste_Auslass:
                Liste_Zubehoer.append(FamName)

    for el in sorted(Liste_Auslass):
        IS_AUSLASS_D.Add(Itemtemplate(el))
    for el in sorted(Liste_Zubehoer):
        IS_VSR.Add(Itemtemplate(el))
        IS_DOSSEL.Add(Itemtemplate(el))
    for el in sorted(Liste_Auslass_Lab):
        IS_AUSLASS.Add(Itemtemplate(el,True))
        IS_AUSLASS_LAB.Add(Itemtemplate(el))
        IS_AUSLASS_24H.Add(Itemtemplate(el))

get_IS()

familien = Familien(IS_AUSLASS,IS_AUSLASS_D,IS_AUSLASS_LAB,IS_AUSLASS_24H,IS_VSR,IS_DOSSEL)
try:familien.ShowDialog()
except Exception as e:
    logger.error(e)
    familien.Close()
    script.exit()

VSR_AUSLASS_LISTE = config.get_value('auslass')
DRO_AUSLASS_LISTE = config.get_value('auslassd')
LAB_AUSLASS_LISTE = config.get_value('auslasslab')
H24_AUSLASS_LISTE = config.get_value('auslass24h')
VSR_FAMILIE_LISTE = config.get_value('vsr')
DRO_FAMILIE_LISTE = config.get_value('drossel')

print(LAB_AUSLASS_LISTE)
# script.exit()
# VSR_AUSLASS_LISTE = config.auslass
# DRO_AUSLASS_LISTE = config.auslassd
# LAB_AUSLASS_LISTE = config.auslasslab
# H24_AUSLASS_LISTE = config.auslass24h
# VSR_FAMILIE_LISTE = config.vsr
# DRO_FAMILIE_LISTE = config.drossel

# def get_IS_VSR():
#     Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
#     Dict = {}
#     for el in Families:
#         FamName = el.Name
#         if FamName in VSR_FAMILIE_LISTE:
#             for typid in el.GetFamilySymbolIds():
#                 typ = doc.GetElement(typid)
#                 typname = typ.LookupParameter('Typenkommentare').AsString()
#                 if typname not in Dict.keys():
#                     Dict[typname] = typid
#                 else:
#                     logger.error('{} und {} haben gleiche Herstellertypname {}.'.format(typid,Dict[typname],typname))
#     return Dict


# Dict_Herstellertyp = get_IS_VSR()

# def get_MEP_NUMMER_NAMR():
#     spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
#     for el in spaces:
#         DICT_MEP_NUMMER_NAME[el.Number] = [el.Number + ' - ' + el.LookupParameter('Name').AsString(),el.Id.ToString()]
#         schacht = el.LookupParameter('TGA_RLT_InstallationsSchacht').AsInteger()
#         if schacht == 1:
#             name = el.LookupParameter('TGA_RLT_InstallationsSchachtName').AsString()
#             if name not in LISTE_SCHACHT:
#                 LISTE_SCHACHT.append(name)

# get_MEP_NUMMER_NAMR() 
# LISTE_SCHACHT.sort()

# def Get_Ueberstrom_Info():
#     """
#     DICT_MEP_UEBERSTROM:  {[raumid:{'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}}
#     ELEMID_UEBER:   [lemid.ToString()]
#     """
#     Baugruppen = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToElements()
#     with forms.ProgressBar(title = "{value}/{max_value} Überströmungsbaugruppen",cancellable=True,step=10) as pb:
#         for n, BGid in enumerate(Baugruppen):
#             if pb.cancelled:
#                 script.exit()
#             pb.update_progress(n + 1, len(Baugruppen))
#             baugruppe = Baugruppe(BGid,logger,DICT_MEP_NUMMER_NAME)
#             if not baugruppe.Pruefen():
#                 continue
#             if not baugruppe.Eingang.raumid in DICT_MEP_UEBERSTROM.keys():
#                 DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
#             if not baugruppe.Ausgang.raumid in DICT_MEP_UEBERSTROM.keys():
#                 DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
#             DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid][baugruppe.Eingang.typ].Add(baugruppe.Eingang)
#             DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid][baugruppe.Ausgang.typ].Add(baugruppe.Ausgang)
#             ELEMID_UEBER.append(baugruppe.Eingang.elemid.ToString())
#             ELEMID_UEBER.append(baugruppe.Ausgang.elemid.ToString())

# Get_Ueberstrom_Info()

# def Get_Auslass_Info():
#     """
#     DICT_MEP_AUSLASS: {raumid:{auslass.art:{auslass.familyandtyp:[auslass(Klasse)]}}}
#     DICT_VSR_MEP: {auslass.vsr: ["auslass.raumnr-auslass.raumname"],auslass.iris: ["auslass.raumnr-auslass.raumname"]}
#     DICT_VSR_MEP_NUR_NUMMER: {auslass.vsr: [auslass.raumnummer],auslass.iris: [auslass.raumnummer]}
#     DICT_VSR_VSRLISTE: {auslass.vsr: [auslass.VSR_Liste]} elemid: [elemid]
#     DICT_MEP_VSR: {auslass.raumid: [auslass.iris,auslass.vsr]}
#     DICT_VSR_AUSLASS: {auslass.vsr:{auslass.art:{auslass.familyandtyp:[auslass]}}}
#     DICT_VSR: {vsrid:vsr}
#     LISTE_VSR: [elemid]
#     LISTE_IRIS: [elemid]

#     """
#     Ductterminalids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
#     with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe", cancellable=True,step=10) as pb:
#         for n, ductid in enumerate(Ductterminalids):
#             if pb.cancelled:
#                 script.exit()
#             pb.update_progress(n + 1, len(Ductterminalids))
#             if ductid.Id.ToString() in ELEMID_UEBER:continue
#             auslass = Luftauslass(ductid,DICT_MEP_NUMMER_NAME,logger,LAB_AUSLASS_LISTE,H24_AUSLASS_LISTE,VSR_FAMILIE_LISTE,DRO_FAMILIE_LISTE,VSR_AUSLASS_LISTE,DRO_AUSLASS_LISTE)
#             einbauteil = Einbauteile(auslass.elem,DICT_MEP_NUMMER_NAME,logger,[auslass])
#             if einbauteil.raumid not in DICT_MEP_EINBAUTEILE.keys():
#                 DICT_MEP_EINBAUTEILE[auslass.raumid] = []
#             DICT_MEP_EINBAUTEILE[auslass.raumid].append(einbauteil)
            

#             if not auslass.raumluftrelevant:
#                 if auslass.raumid not in DICT_MEP_UN_AUNLASS.keys():
#                     DICT_MEP_UN_AUNLASS[auslass.raumid] = []
#                 DICT_MEP_UN_AUNLASS[auslass.raumid].append(auslass)

#                 continue
            
#             if not auslass.raumid in DICT_MEP_AUSLASS.keys():
#                  DICT_MEP_AUSLASS[auslass.raumid] = {}
#             if not auslass.art in DICT_MEP_AUSLASS[auslass.raumid].keys():
#                 DICT_MEP_AUSLASS[auslass.raumid][auslass.art] = {}
#             if not auslass.familyandtyp in DICT_MEP_AUSLASS[auslass.raumid][auslass.art].keys():
#                 DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp] = []
#             DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp].append(auslass)

#             for elemid in auslass.Einbauteile_Liste:
               
#                 if elemid not in DICT_EINBAUTEILE_AUSLASS.keys():
#                     DICT_EINBAUTEILE_AUSLASS[elemid] = []
#                 DICT_EINBAUTEILE_AUSLASS[elemid].append(auslass)


#             if auslass.familyandtyp in VSR_AUSLASS_LISTE:
#                 continue
#             if not auslass.vsr:
#                 if auslass.Muster_Pruefen() != True:
#                     logger.error('Kein VSR mit Auslass {} angeschlossen. Raum {}, Familie: {}'.format(auslass.elemid,auslass.raumnummer,auslass.familyandtyp))
#                 continue
            
#             if not auslass.vsr in DICT_VSR_MEP.keys():
#                 DICT_VSR_MEP[auslass.vsr] = [auslass.raum]
#                 DICT_VSR_MEP_NUR_NUMMER[auslass.vsr] = [auslass.raumnummer]
#             else:
#                 if auslass.raum not in DICT_VSR_MEP[auslass.vsr]:
#                     DICT_VSR_MEP[auslass.vsr].append(auslass.raum)
#                     DICT_VSR_MEP_NUR_NUMMER[auslass.vsr].append(auslass.raumnummer)
#             if auslass.iris not in [-1,'']:

#                 if not auslass.iris in DICT_VSR_MEP.keys():
#                     DICT_VSR_MEP[auslass.iris] = [auslass.raum]
#                     DICT_VSR_MEP_NUR_NUMMER[auslass.iris] = [auslass.raumnummer]
#                 else:
#                     if auslass.raum not in DICT_VSR_MEP[auslass.iris]:
#                         DICT_VSR_MEP[auslass.iris].append(auslass.raum)
#                         DICT_VSR_MEP_NUR_NUMMER[auslass.iris].append(auslass.raumnummer)
            
#             if len(auslass.VSR_Liste) != 0 and auslass.vsr not in auslass.VSR_Liste:
#                 if auslass.vsr not in DICT_VSR_VSRLISTE.keys():
#                     DICT_VSR_VSRLISTE[auslass.vsr] = auslass.VSR_Liste
#                 else:
#                     DICT_VSR_VSRLISTE[auslass.vsr].extend(auslass.VSR_Liste)
#             else:
#                 LISTE_VSR.append(auslass.vsr)
            
#             if auslass.iris not in [-1,'']:
#                 LISTE_IRIS.append(auslass.iris)

#             if not auslass.raumid in DICT_MEP_VSR.keys():
#                 DICT_MEP_VSR[auslass.raumid] = []
#                 DICT_MEP_VSR[auslass.raumid].append(auslass.vsr) 
#                 if auslass.iris not in [-1,'']:
#                     DICT_MEP_VSR[auslass.raumid].append(auslass.iris)
#             else:
#                 if auslass.vsr not in DICT_MEP_VSR[auslass.raumid]:
#                     DICT_MEP_VSR[auslass.raumid].append(auslass.vsr)
#                 if auslass.iris not in [-1,'']:
#                     DICT_MEP_VSR[auslass.raumid].append(auslass.iris)

#             if not auslass.vsr in DICT_VSR_AUSLASS.keys():
#                 DICT_VSR_AUSLASS[auslass.vsr] = {}
#             if not auslass.art in DICT_VSR_AUSLASS[auslass.vsr].keys():
#                 DICT_VSR_AUSLASS[auslass.vsr][auslass.art] = {}
#             if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.vsr][auslass.art].keys():
#                 DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp] = []
#             DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp].append(auslass)
#             if auslass.iris not in [-1,'']:
#                 if not auslass.iris in DICT_VSR_AUSLASS.keys():
#                     DICT_VSR_AUSLASS[auslass.iris] = {}
#                 if not auslass.art in DICT_VSR_AUSLASS[auslass.iris].keys():
#                     DICT_VSR_AUSLASS[auslass.iris][auslass.art] = {}
#                 if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.iris][auslass.art].keys():
#                     DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp] = []
#                 DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp].append(auslass)  


#     liste_temp = DICT_VSR_AUSLASS.keys()[:]
#     liste_temp1 = LISTE_VSR[:]
#     liste_temp2 = LISTE_IRIS[:]
#     liste_temp.extend(liste_temp1)
#     liste_temp.extend(liste_temp2)
#     liste_temp = set(liste_temp)
#     liste_temp = list(liste_temp)
   

#     with forms.ProgressBar(title = "{value}/{max_value} Volumenstromregler", cancellable=True, step=10) as pb:
#         for n, vsrid in enumerate(liste_temp):
#             if pb.cancelled:
#                 script.exit()
#             pb.update_progress(n + 1, len(liste_temp))
#             if vsrid in DICT_VSR_VSRLISTE.keys():
#                 vsrliste = set(DICT_VSR_VSRLISTE[vsrid])
#                 vsrliste = list(vsrliste)                
#                 for vsrid_neu in vsrliste:
#                     if vsrid != vsrid_neu:
#                         if vsrid_neu not in DICT_VSR_AUSLASS.keys():
#                             logger.error('Kein Auslass mit VSR {} verbunden.'.format(vsrid_neu))
#                             continue
#                         for key0 in DICT_VSR_AUSLASS[vsrid_neu].keys():
#                             for key1 in DICT_VSR_AUSLASS[vsrid_neu][key0].keys():
#                                 for auslass in DICT_VSR_AUSLASS[vsrid_neu][key0][key1]:
#                                     if key0 not in DICT_VSR_AUSLASS[vsrid].keys():
#                                         DICT_VSR_AUSLASS[vsrid][key0] = {}
#                                     if key1 not in DICT_VSR_AUSLASS[vsrid][key0].keys():
#                                         DICT_VSR_AUSLASS[vsrid][key0][key1] = []
#                                     if auslass not in DICT_VSR_AUSLASS[vsrid][key0][key1]:
#                                         DICT_VSR_AUSLASS[vsrid][key0][key1].append(auslass)
#                                     if auslass.raumid not in DICT_MEP_VSR.keys():
#                                         DICT_MEP_VSR[auslass.raumid] = []
#                                     if vsrid not in DICT_MEP_VSR[auslass.raumid]:DICT_MEP_VSR[auslass.raumid].append(vsrid) 
#                         for mep in DICT_VSR_MEP[vsrid_neu]:
#                             if mep not in DICT_VSR_MEP[vsrid]:
#                                 DICT_VSR_MEP[vsrid].append(mep)
#                         for mep in DICT_VSR_MEP_NUR_NUMMER[vsrid_neu]:
#                             if mep not in DICT_VSR_MEP_NUR_NUMMER[vsrid]:
#                                 DICT_VSR_MEP_NUR_NUMMER[vsrid].append(mep)
                        
#                         if not auslass.raumid in DICT_MEP_VSR.keys():
#                             DICT_MEP_VSR[auslass.raumid] = []
#                             DICT_MEP_VSR[auslass.raumid].append(auslass.vsr) 
#                             if auslass.iris not in [-1,'']:
#                                 DICT_MEP_VSR[auslass.raumid].append(auslass.iris)
                  
#             rvt_vsr = doc.GetElement(DB.ElementId(int(vsrid)))
#             vsr = VSR(rvt_vsr,DICT_MEP_NUMMER_NAME,logger,DICT_DatenBank,Liste_Fabrikat,Liste_Datenbank,Liste_Datenbank1,Dict_Herstellertyp)
#             if vsr.elemid.ToString() in LISTE_IRIS:
#                 vsr.IsIris = True
#             elif vsrid in DICT_VSR_VSRLISTE.keys():
#                 vsr.IsHaupt = True

#             vsr.Auslass = ObservableCollection[Luftauslass]()
#             for art in sorted(DICT_VSR_AUSLASS[vsrid].keys()):
#                 for fam in sorted(DICT_VSR_AUSLASS[vsrid][art].keys()):
#                     for terminal in DICT_VSR_AUSLASS[vsrid][art][fam]:
#                         vsr.Auslass.Add(terminal)
#             vsr.get_Art()
#             vsr.Luftmengenermitteln()
#             vsr.vsrauswaelten()
#             vsr.vsrueberpruefen()
#             vsr.liste_Raum = DICT_VSR_MEP[vsrid]
#             vsr.liste_Raum_nurNummer = DICT_VSR_MEP_NUR_NUMMER[vsrid]
            
#             DICT_VSR[vsrid] = vsr
#             for auslass in vsr.Auslass:
#                 if vsr.IsIris:
#                     auslass.IRIS_Class = vsr
#                 elif vsr.IsHaupt:
#                     auslass.Haupt_Class = vsr
#                 else:
#                     auslass.VSR_Class = vsr
            
#             if vsr.raumid not in DICT_MEP_VSRKLASSE.keys():
#                 DICT_MEP_VSRKLASSE[vsr.raumid] = []
#             DICT_MEP_VSRKLASSE[vsr.raumid].append(vsr)

#     with forms.ProgressBar(title = "{value}/{max_value} Einbauteile", cancellable=True, step=10) as pb:
#         for n, einbauteilid in enumerate(DICT_EINBAUTEILE_AUSLASS.keys()):
#             if pb.cancelled:
#                 script.exit()
#             pb.update_progress(n + 1, len(DICT_EINBAUTEILE_AUSLASS.keys()))
                  
#             rvt_einbauteil = doc.GetElement(DB.ElementId(int(einbauteilid)))
#             einbauteil = Einbauteile(rvt_einbauteil,DICT_MEP_NUMMER_NAME,logger,DICT_EINBAUTEILE_AUSLASS[einbauteilid])
#             if einbauteil.raumid not in DICT_MEP_EINBAUTEILE.keys():
#                 DICT_MEP_EINBAUTEILE[einbauteil.raumid] = []
#             DICT_MEP_EINBAUTEILE[einbauteil.raumid].append(einbauteil)

# Get_Auslass_Info()

# spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
# Dict_Ueber_Manuell = {}
# for ele in spaces_collector:
#     raum = get_value(ele.LookupParameter("TGA_RLT_RaumÜberströmungAus"))
#     if raum:
#         summe2 = get_value(ele.LookupParameter('TGA_RLT_RaumÜberströmungMenge'))
#         if raum not in Dict_Ueber_Manuell.keys():
#             Dict_Ueber_Manuell[raum] = summe2
#         else:Dict_Ueber_Manuell[raum] += summe2

# DICT_MEP_ITEMSSOIRCE = {}
# LISTE_MEP = ObservableCollection[MEPFOREXPORTITEM]()
# Schachte = []
# def Get_MEP_Info():
#     spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
#     with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb:
#         for n,space_id in enumerate(spaces):
#             if pb.cancelled:
#                 script.exit()
#             pb.update_progress(n+1, len(spaces))
            
#             if space_id.Id.ToString() in DICT_MEP_VSR.keys():
#                 list_vsr = ObservableCollection[VSR]()
#                 dict_vsr = {}
#                 for e in DICT_MEP_VSR[space_id.Id.ToString()]:
#                     temp = DICT_VSR[e]
#                     if temp.art not in dict_vsr.keys():
#                         dict_vsr[temp.art] = {}
#                     if temp.familyandtyp not in dict_vsr[temp.art].keys():
#                         dict_vsr[temp.art][temp.familyandtyp] = []
#                     dict_vsr[temp.art][temp.familyandtyp].append(temp)
#                 for art in sorted(dict_vsr.keys()):
#                     for fam in sorted(dict_vsr[art].keys()):
#                         for kla in dict_vsr[art][fam]:
#                             try:list_vsr.Add(kla)
#                             except:pass
#             else:list_vsr = ObservableCollection[VSR]()
#             einbauteilliste = []
#             if space_id.Id.ToString() in DICT_MEP_EINBAUTEILE.keys():
#                 einbauteilliste = DICT_MEP_EINBAUTEILE[space_id.Id.ToString()]
            
#             vsrliste = []
#             if space_id.Id.ToString() in DICT_MEP_VSRKLASSE.keys():
#                 vsrliste = DICT_MEP_VSRKLASSE[space_id.Id.ToString()]
            

#             mepraum = MEPRaum(space_id,list_vsr,LISTE_SCHACHT,logger,DICT_MEP_AUSLASS,DICT_MEP_UEBERSTROM,Dict_Ueber_Manuell,DICT_MEP_UN_AUNLASS,einbauteilliste,vsrliste)
            
#             if mepraum.flaeche == 0 and mepraum.volumen == 0:
#                 continue
#             if mepraum.IsSchacht:Schachte.append(mepraum)
#             # if not mepraum.IsSchacht:
#             DICT_MEP_ITEMSSOIRCE[mepraum.Raumnr] = mepraum
#     for Raumnr in sorted(DICT_MEP_ITEMSSOIRCE.keys()):
#         LISTE_MEP.Add(MEPFOREXPORTITEM(Raumnr,DICT_MEP_ITEMSSOIRCE[Raumnr].bezugsname))
    

# Get_MEP_Info()

# uebersicht = MEPRaum_Uebersicht(RED,BLACK,LISTE_MEP,VISIBLE,HIDDEN,DICT_MEP_ITEMSSOIRCE,Schachte)
# uebersicht.externaleventliste.GUI = uebersicht
# uebersicht.Show()
