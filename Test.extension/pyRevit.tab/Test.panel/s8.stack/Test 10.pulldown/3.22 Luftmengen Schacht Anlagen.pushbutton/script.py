# coding: utf8
from rpw import revit,DB
from pyrevit import script, forms
from IGF_log import getlog
from IGF_lib import get_value

__title__ = "3.22 Raumluft Schächte & Anlagen (neu)"
__doc__ = """Anlagen- und Schachtberechnung

!!!Achtung: neue Parameter IGF_AnlagenNR... werden benutzt.

!!!Achtung: Bitte nutzen Sie zuerst Funktion 3.20 Ebenen sortieren. 

imput Parameter:
-------------------------
*************************
MEP Räume:
IGF_RLT_Verteilung_EbenenName: Name der Ebene für Verteilung
IGF_RLT_Verteilung_EbenenSortieren: Nummer zum EBenesortierung für Verteilung
TGA_RLT_AnlagenNrZuluft: RLT-Anlagenummer für Zuluft
TGA_RLT_AnlagenNrAbluft: RLT-Anlagenummer für Abluft
TGA_RLT_AnlagenNr24hAbluft: RLT-Anlagenummer für 24h Abluft
TGA_RLT_InstallationsSchacht: Ist InstallationsSchacht Ja/Nein
TGA_RLT_InstallationsSchachtName: InstallationsSchachtname
IGF_RLT_AnlagenRaumZuluft: Zuluft über RLT-Anlage
IGF_RLT_AnlagenRaumAbluft: Abluft über RLT-Anlage
IGF_RLT_AnlagenRaum24hAbluft: 24h Abluft über RLT-Anlage
TGA_RLT_SchachtZuluft: Schacht für Zuluft
TGA_RLT_SchachtAbluft: Schacht für Abluft
TGA_RLT_Schacht24hAbluft: Schacht für 24h Abluft
*************************
Luftkanal System (Type):
IGF_X_SystemName: SystemName
IGF_X_AnlagenGeräteNr: RLT-Anlagen Gerätenummer
IGF_X_AnlagenGeräteAnzahl: RLT-Anlagen Geräteanzahl
IGF_X_AnlagenNr: RLT-Anlagen Nummer
IGF_X_AnlagenProzentualAnzahl: RLT-Anlagen Anzahl der prozentualen Geräten
*************************
-------------------------

output Parameter:
-------------------------
*************************
MEP-Räume:
TGA_RLT_SchachtZuluftMenge:
TGA_RLT_SchachtAbluftMenge
TGA_RLT_Schacht24hAbluftMenge
IGF_RLT_VerteilungZuluft
IGF_RLT_VerteilungAbluft
IGF_RLT_Verteilung24hAbluft
*************************
Luftkanal System (Type):
IGF_RLT_AnlagenZuMenge: RLT-Anlagen Zuluftmengen
IGF_RLT_AnlagenProzentualZuMenge: RLT-Anlagen prozentuale Zuluftmengen
IGF_RLT_AnlagenAbMenge: RLT-Anlagen Abluftmengen
IGF_RLT_AnlagenProzentualAbMenge: RLT-Anlagen prozentuale Abluftmengen
IGF_RLT_Anlagen24hAbMenge: RLT-Anlagen 24h-Abluftmengen
IGF_RLT_AnlagenProzentual24hAbMenge: RLT-Anlagen prozentuale 24h-Abluftmengen
*************************
-------------------------

Update: Sub-Schächte

[2023.02.07]
Version: 1.3
"""
__author__ = "Menghui Zhang"


try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

# MEP Räume aus aktueller Projekt
spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()

# Systemen aus Projekt
System_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType().ToElements()
list_id = []
systemen = []
for el in System_collector:
    type0 = el.GetTypeId()
    if type0.ToString() not in list_id:
        list_id.append(type0.ToString())
        systemen.append(type0)

if not (spaces and systemen):
    logger.error("Keine MEP Räume/Luftkanalsystemen in aktueller Projekt gefunden")
    script.exit()

class MEPRaum:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.Raum_RZU = []
        self.Raum_RAB = []
        self.Raum_LAB = []
        self.Raum_H24 = []
        self.Raum_TAB = []
        self.Raum_TZU = []
        self.Dict_Ebene = {}

        self.SUB_RZU_Schacht = []
        self.SUB_RAB_Schacht = []
        self.SUB_TZU_Schacht = []
        self.SUB_TAB_Schacht = []
        self.SUB_LAB_Schacht = []
        self.SUB_H24_Schacht = []
        
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['ebene', 'IGF_RLT_Verteilung_EbenenName'],
            ['ebene_nr', 'IGF_RLT_Verteilung_EbenenSortieren'],

            #Schacht
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],

            # Luftmengen
            ['RAB', 'IGF_RLT_Luftmenge_RAB'],
            ['RZU', 'IGF_RLT_Luftmenge_RZU'],
            ['H24', 'IGF_RLT_Luftmenge_24h'],
            ['LAB', 'IGF_RLT_Luftmenge_LAB'],
            ['TAB', 'IGF_RLT_Luftmenge_max_TAB'],
            ['TZU', 'IGF_RLT_Luftmenge_max_TZU'],

             # Anlagen
            ['ARAB', 'IGF_RLT_AnlagenNr_RAB'],
            ['ARZU', 'IGF_RLT_AnlagenNr_RZU'],
            ['AH24', 'IGF_RLT_AnlagenNr_24h'],
            ['ATAB', 'IGF_RLT_AnlagenNr_TAB'],
            ['ATZU', 'IGF_RLT_AnlagenNr_TZU'],
            ['ALAB', 'IGF_RLT_AnlagenNr_LAB'],

            # Schacht
            ['SRZU', 'TGA_RLT_SchachtZuluft'],
            ['SRAB', 'TGA_RLT_SchachtAbluft'],
            ['SH24', 'TGA_RLT_Schacht24hAbluft'],
            ['STAB', 'IGF_RLT_Schacht_TAB'],
            ['STZU', 'IGF_RLT_Schacht_TZU'],
            ['SLAB', 'IGF_RLT_Schacht_LAB'],

            ['fk_Reduzieren','IGF_RLT_Schacht-ReduziertFaktor']
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(revit_name))
        
        if not self.fk_Reduzieren or self.fk_Reduzieren == 0:
            self.fk_Reduzieren = 1
        
        # Prüfung
        self.Ueberpruefen()
        
        # Schachtverteilung
        self.SRZU_Menge = ''
        self.SRZU_Verteilen = ''
        self.SRAB_Menge = ''
        self.SRAB_Verteilen = ''
        self.SH24_Menge = ''
        self.SH24_Verteilen = ''
        self.SLAB_Menge = ''
        self.SLAB_Verteilen = ''
        self.STZU_Menge = ''
        self.STZU_Verteilen = ''
        self.STAB_Menge = ''
        self.STAB_Verteilen = ''
        
        # reduzierte Schacht
        # self.STZUR_Menge = ''
        # self.STABR_Menge = ''
        self.SRZUR_Menge = ''
        self.SRABR_Menge = ''
        self.SLABR_Menge = ''
        self.SH24R_Menge = ''
        
    
    def Ueberpruefen(self):
        def Ueberprufen_Einzeln(raumluft,schachtname,anlagename,art):
            if raumluft == None:
                return 0
            if raumluft > 0:
                if not anlagename:
                    logger.error("{}-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(art,self.nummer,self.name,self.ebene,self.element_id.ToString()))
                if not schachtname:
                    logger.error("{}-Schacht in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(art,self.nummer,self.name,self.ebene,self.element_id.ToString()))
                return raumluft
            else:
                return 0
        self.RAB = Ueberprufen_Einzeln(self.RAB,self.SRAB,self.ARAB,'Raumabluft')
        self.RZU = Ueberprufen_Einzeln(self.RZU,self.SRZU,self.ARZU,'Raumzuluft')
        self.H24 = Ueberprufen_Einzeln(self.H24,self.SH24,self.AH24,'24h-Abluft')
        self.LAB = Ueberprufen_Einzeln(self.LAB,self.SLAB,self.ALAB,'Laborabluft')
        self.TZU = Ueberprufen_Einzeln(self.TZU,self.STZU,self.ATZU,'TierhaltungZuluft')
        self.TAB = Ueberprufen_Einzeln(self.TAB,self.STAB,self.ATAB,'TierhaltungAbluft')
            
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.element.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)

        wert_schreiben("IGF_RLT_Schacht_RZU_Menge", self.SRZU_Menge)
        wert_schreiben("IGF_RLT_Schacht_RAB_Menge", self.SRAB_Menge)
        wert_schreiben("IGF_RLT_Schacht_24h_Menge", self.SH24_Menge)
        wert_schreiben("IGF_RLT_Schacht_LAB_Menge", self.SLAB_Menge)
        wert_schreiben("IGF_RLT_Schacht_TZU_Menge", self.STZU_Menge)
        wert_schreiben("IGF_RLT_Schacht_TAB_Menge", self.STAB_Menge)

        wert_schreiben("IGF_RLT_Verteilung_LAB", self.SLAB_Verteilen)
        wert_schreiben("IGF_RLT_VerteilungZuluft", self.SRZU_Verteilen)
        wert_schreiben("IGF_RLT_VerteilungAbluft", self.SRAB_Verteilen)
        wert_schreiben("IGF_RLT_Verteilung24hAbluft", self.SH24_Verteilen)

        wert_schreiben("IGF_RLT_Schacht-Reduziert_RZU", self.SRZUR_Menge)
        wert_schreiben("IGF_RLT_Schacht-Reduziert_RAB", self.SRABR_Menge)
        wert_schreiben("IGF_RLT_Schacht-Reduziert_24h", self.SH24R_Menge)
        wert_schreiben("IGF_RLT_Schacht-Reduziert_LAB", self.SLABR_Menge)
        wert_schreiben("IGF_RLT_Schacht-ReduziertFaktor", self.fk_Reduzieren)

    def Datenanalyse(self):
        def Textbearbeiten(_dict):
            Text = '[m3/h] - '
            if len(_dict) == 0:return ''
            for ebene_nr in sorted(_dict.keys()):
                ebenename = self.Dict_Ebene[ebene_nr]
                try:Text += ebenename + ': '
                except:Text +=  'None: '
                for anl in sorted(_dict[ebene_nr].keys()):
                    try:Text += 'Anl ' + str(anl) + ': ' + str(_dict[ebene_nr][anl]) + ', '
                    except:Text +=  'Anl None' + ': ' + str(_dict[ebene_nr][anl]) + ', '
            return Text[:-2]

        Dict_RZU = {}
        RZU = 0
        Dict_RAB = {}
        RAB = 0
        Dict_H24 = {}
        H24 = 0
        Dict_LAB = {}
        LAB = 0
        Dict_TZU = {}
        TZU = 0
        Dict_TAB = {}
        TAB = 0

        for el in self.Raum_RZU:
            RZU += el.RZU
            if el.ebene_nr not in Dict_RZU.keys():
                Dict_RZU[el.ebene_nr] = {}
            if el.ARZU not in Dict_RZU[el.ebene_nr].keys():
                Dict_RZU[el.ebene_nr][el.ARZU] = 0 
            Dict_RZU[el.ebene_nr][el.ARZU] += el.RZU * self.fk_Reduzieren

        for el in self.Raum_RAB:
            RAB += el.RAB
            if el.ebene_nr not in Dict_RAB.keys():
                Dict_RAB[el.ebene_nr] = {}
            if el.ARAB not in Dict_RAB[el.ebene_nr].keys():
                Dict_RAB[el.ebene_nr][el.ARAB] = 0
            Dict_RAB[el.ebene_nr][el.ARAB] += el.RAB * self.fk_Reduzieren

        for el in self.Raum_LAB:
            LAB += el.LAB
            if el.ebene_nr not in Dict_LAB.keys():
                Dict_LAB[el.ebene_nr] = {}
            if el.ALAB not in Dict_LAB[el.ebene_nr].keys():
                Dict_LAB[el.ebene_nr][el.ALAB] = 0
            Dict_LAB[el.ebene_nr][el.ALAB] += el.LAB * self.fk_Reduzieren

            RAB += el.LAB
            if el.ebene_nr not in Dict_RAB.keys():
                Dict_RAB[el.ebene_nr] = {}
            if el.ARAB not in Dict_RAB[el.ebene_nr].keys():
                Dict_RAB[el.ebene_nr][el.ARAB] = 0
            Dict_RAB[el.ebene_nr][el.ARAB] += el.LAB * self.fk_Reduzieren

        for el in self.Raum_H24:
            H24 += el.H24
            if el.ebene_nr not in Dict_H24.keys():
                Dict_H24[el.ebene_nr] = {}
            if el.AH24 not in Dict_H24[el.ebene_nr].keys():
                Dict_H24[el.ebene_nr][el.AH24] = 0
            Dict_H24[el.ebene_nr][el.AH24] += el.H24 * self.fk_Reduzieren
        
        for el in self.Raum_TZU:
            TZU += el.TZU
            if el.ebene_nr not in Dict_TZU.keys():
                Dict_TZU[el.ebene_nr] = {}
            if el.ATZU not in Dict_TZU[el.ebene_nr].keys():
                Dict_TZU[el.ebene_nr][el.ATZU] = 0
            Dict_TZU[el.ebene_nr][el.ATZU] += el.TZU * self.fk_Reduzieren
        
        for el in self.Raum_TAB:
            TAB += el.TAB
            if el.ebene_nr not in Dict_TAB.keys():
                Dict_TAB[el.ebene_nr] = {}
            if el.ATAB not in Dict_TAB[el.ebene_nr].keys():
                Dict_TAB[el.ebene_nr][el.ATAB] = 0
            Dict_TAB[el.ebene_nr][el.ATAB] += el.TAB * self.fk_Reduzieren
                
        self.SRZU_Menge = RZU
        self.SRAB_Menge = RAB
        self.SLAB_Menge = LAB
        self.SH24_Menge = H24
        self.STAB_Menge = TAB
        self.STZU_Menge = TZU

        self.SRZUR_Menge = RZU * self.fk_Reduzieren
        self.SRABR_Menge = RAB * self.fk_Reduzieren
        self.SLABR_Menge = LAB * self.fk_Reduzieren
        self.SH24R_Menge = H24 * self.fk_Reduzieren

        self.SRZU_Verteilen = Textbearbeiten(Dict_RZU)
        self.SRAB_Verteilen = Textbearbeiten(Dict_RAB)
        self.SH24_Verteilen = Textbearbeiten(Dict_H24)
        self.SLAB_Verteilen = Textbearbeiten(Dict_LAB)

    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)
        if not param:
            logger.error("Parameter ({}) konnte nicht gefunden werden".format(param_name))
            return None

        return get_value(param)

    def weiter_analyse(self):
        if len(self.SUB_RZU_Schacht) > 0:
            text = ''
            for el in self.SUB_RZU_Schacht:
                self.SRZU_Menge += el.SRZU_Menge
                self.SRZUR_Menge += el.SRZUR_Menge
                if el.SRZU_Menge > 0:
                    text += el.Install_Schacht_Name +': ' + str(el.SRZU_Menge) + ', '
            if text:
                text = text[:-2]
            
            if self.SRZU_Verteilen:
                if text:self.SRZU_Verteilen = self.SRZU_Verteilen + ', ' + text
            else:
                if text:self.SRZU_Verteilen = '[m3/h] - ' + text
        if len(self.SUB_RAB_Schacht) > 0:
            text = ''
            for el in self.SUB_RAB_Schacht:
                self.SRAB_Menge += el.SRAB_Menge
                self.SRABR_Menge += el.SRABR_Menge
                if el.SRAB_Menge > 0:
                    text += el.Install_Schacht_Name +': ' + str(el.SRAB_Menge) + ', '
            if text:
                text = text[:-2]
            
            if self.SRAB_Verteilen:
                if text:self.SRAB_Verteilen = self.SRAB_Verteilen + ', ' + text
            else:
                if text:self.SRAB_Verteilen = '[m3/h] - ' + text
        if len(self.SUB_LAB_Schacht) > 0:
            text = ''
            for el in self.SUB_LAB_Schacht:
                self.SLAB_Menge += el.SLAB_Menge
                self.SLABR_Menge += el.SLABR_Menge
                if el.SLAB_Menge > 0:
                    text += el.Install_Schacht_Name +': ' + str(el.SLAB_Menge) + ', '
            if text:
                text = text[:-2]
            
            if self.SLAB_Verteilen:
                if text:self.SLAB_Verteilen = self.SLAB_Verteilen + ', ' + text
            else:
                if text:self.SLAB_Verteilen = '[m3/h] - ' + text
        if len(self.SUB_H24_Schacht) > 0:
            text = ''
            for el in self.SUB_H24_Schacht:
                self.SH24_Menge += el.SH24_Menge
                self.SH24R_Menge += el.SH24R_Menge
                if el.SH24_Menge > 0:
                    text += el.Install_Schacht_Name +': ' + str(el.SH24_Menge) + ', '
            if text:
                text = text[:-2]
            
            if self.SH24_Verteilen:
                if text:self.SH24_Verteilen = self.SH24_Verteilen + ', ' + text
            else:
                if text:self.SH24_Verteilen = '[m3/h] - ' + text

class DuctSystem:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.ARZU = []
        self.ARAB = []
        self.ALAB = []
        self.AH24 = []
        self.ATAB = []
        self.ATZU = []

        attr = [
            ['name', 'IGF_X_SystemName'],
            ['Ger_Anzahl', 'IGF_X_AnlagenGeräteAnzahl'],
            ['Anl_Nr', 'IGF_X_AnlagenNr'],
            ['Anl_Pro_Anz', 'IGF_X_AnlagenProzentualAnzahl'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(revit_name))

        if self.Ger_Anzahl == 0:
            self.Ger_Anzahl = 1
        if self.Anl_Pro_Anz == 0:
            self.Anl_Pro_Anz = 1

        self.Anl_Nr = str(int(self.Anl_Nr))

        self.zuluft = 0
        self.abluft = 0
        self.ab24h = 0
        self.lab = 0
        self.tzu = 0
        self.tab = 0
        self.zu_geraete = 0
        self.zu_geraete_p = 0
        self.ab_geraete = 0
        self.ab_geraete_p = 0
        self.ab_geraete24h = 0
        self.ab_geraete24h_p = 0
        self.lab_geraete = 0
        self.lab_geraete_p = 0
        self.tab_geraete = 0
        self.tab_geraete_p = 0
        self.tzu_geraete = 0
        self.tab_geraete_p = 0

    def Datenanalyse(self):
        RZU = 0
        RAB = 0
        TZU = 0
        TAB = 0
        LAB = 0
        H24 = 0
        for el in self.ARZU:
            RZU += el.RZU
        
        for el in self.ARAB:
            RAB += el.RAB
        
        for el in self.ATZU:
            TZU += el.TZU
        
        for el in self.ATAB:
            TAB += el.TAB
        
        for el in self.ALAB:
            LAB += el.LAB
        
        for el in self.AH24:
            H24 += el.H24

        self.zuluft = RZU
        self.abluft = RAB
        self.ab24h = H24
        self.tab = TAB
        self.lab = LAB
        self.tzu = TZU

    def Berechnen(self):
        self.Datenanalyse()
        self.zu_geraete = round(self.zuluft / int(self.Ger_Anzahl),1)
        self.zu_geraete_p = round(self.zuluft / int(self.Anl_Pro_Anz),1)
        self.ab_geraete = round(self.abluft / int(self.Ger_Anzahl),1)
        self.ab_geraete_p = round(self.abluft / int(self.Anl_Pro_Anz),1)
        self.ab_geraete24h = round(self.ab24h / int(self.Ger_Anzahl),1)
        self.ab_geraete24h_p = round(self.ab24h / int(self.Anl_Pro_Anz),1)
        self.tzu_geraete = round(self.tzu / int(self.Ger_Anzahl),1)
        self.tzu_geraete_p = round(self.tzu / int(self.Anl_Pro_Anz),1)
        self.tab_geraete = round(self.tab / int(self.Ger_Anzahl),1)
        self.tab_geraete_p = round(self.tab / int(self.Anl_Pro_Anz),1)
        self.lab_geraete = round(self.lab / int(self.Ger_Anzahl),1)
        self.lab_geraete_p = round(self.lab / int(self.Anl_Pro_Anz),1)

    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter ({}) konnte nicht gefunden werden".format(param_name))
            return None

        return get_value(param)

    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.element.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)
        
        wert_schreiben("IGF_RLT_AnlagenZuMenge", self.zu_geraete)
        wert_schreiben("IGF_RLT_AnlagenProzentualZuMenge", self.zu_geraete_p)
        wert_schreiben("IGF_RLT_AnlagenAbMenge", self.ab_geraete)
        wert_schreiben("IGF_RLT_AnlagenProzentualAbMenge", self.ab_geraete_p)
        wert_schreiben("IGF_RLT_Anlagen24hAbMenge", self.ab_geraete24h)
        wert_schreiben("IGF_RLT_AnlagenProzentual24hAbMenge", self.ab_geraete24h_p)

        wert_schreiben("IGF_RLT_AnlagenLABMenge", self.lab_geraete)
        wert_schreiben("IGF_RLT_AnlagenProzentualLABMenge", self.lab_geraete_p)
        wert_schreiben("IGF_RLT_AnlagenTABMenge", self.tab_geraete)
        wert_schreiben("IGF_RLT_AnlagenProzentualTABMenge", self.tab_geraete_p)
        wert_schreiben("IGF_RLT_AnlagenTZUMenge", self.tzu_geraete)
        wert_schreiben("IGF_RLT_AnlagenProzentualTZUMenge", self.tzu_geraete_p)

        wert_schreiben("IGF_X_AnlagenGeräteAnzahl", self.Ger_Anzahl)
        wert_schreiben("IGF_X_AnlagenProzentualAnzahl", self.Anl_Pro_Anz)

SRZU = {}
SRAB = {}
SLAB = {}
S24H = {}

Schacht_liste = []
Schacht_Daten = {}
Ebene_nr_name = {}
Anlagen_Liste = {}

SRZU = {}
SRAB = {}
SLAB = {}
SH24 = {}
STZU = {}
STAB = {}
ARZU = {}
ARAB = {}
ALAB = {}
AH24 = {}
ATAB = {}
ATZU = {}

SUB_RZU = {}
SUB_RAB = {}
SUB_TZU = {}
SUB_TAB = {}
SUB_H24 = {}
SUB_LAB = {}

with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=min(10,int(len(spaces)/300.0))) as pb:
    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        # Schacht berechnen
        if mepraum.Install_Schacht:
            Schacht_liste.append(mepraum)
            if mepraum.SRZU and mepraum.SRZU != mepraum.Install_Schacht_Name:
                if mepraum.SRZU not in SUB_RZU.keys():
                    SUB_RZU[mepraum.SRZU] = []
                SUB_RZU[mepraum.SRZU].append(mepraum)
            if mepraum.SRAB and mepraum.SRAB != mepraum.Install_Schacht_Name:
                if mepraum.SRAB not in SUB_RAB.keys():
                    SUB_RAB[mepraum.SRAB] = []
                SUB_RAB[mepraum.SRAB].append(mepraum)
            if mepraum.STZU and mepraum.STZU != mepraum.Install_Schacht_Name:
                if mepraum.STZU not in SUB_TZU.keys():
                    SUB_TZU[mepraum.STZU] = []
                SUB_TZU[mepraum.STZU].append(mepraum)
            if mepraum.STAB and mepraum.STAB != mepraum.Install_Schacht_Name:
                if mepraum.STAB not in SUB_TAB.keys():
                    SUB_TAB[mepraum.STAB] = []
                SUB_TAB[mepraum.STAB].append(mepraum)
            if mepraum.SLAB and mepraum.SLAB != mepraum.Install_Schacht_Name:
                if mepraum.SLAB not in SUB_LAB.keys():
                    SUB_LAB[mepraum.SLAB] = []
                SUB_LAB[mepraum.SLAB].append(mepraum)
            if mepraum.SH24 and mepraum.SH24 != mepraum.Install_Schacht_Name:
                if mepraum.SH24 not in SUB_H24.keys():
                    SUB_H24[mepraum.SH24] = []
                SUB_H24[mepraum.SH24].append(mepraum)

        else:
            if not mepraum.ebene_nr in Ebene_nr_name.keys():
                Ebene_nr_name[mepraum.ebene_nr] = mepraum.ebene

            if float(mepraum.RZU) > 1:
                if not mepraum.SRZU in SRZU.keys():
                    SRZU[mepraum.SRZU] = []
                SRZU[mepraum.SRZU].append(mepraum)

            if float(mepraum.RAB) > 1:
                if not mepraum.SRAB in SRAB.keys():
                    SRAB[mepraum.SRAB] = []
                SRAB[mepraum.SRAB].append(mepraum)

            if float(mepraum.LAB) > 1:
                if not mepraum.SLAB in SLAB.keys():
                    SLAB[mepraum.SLAB] = []
                SLAB[mepraum.SLAB].append(mepraum)

            if float(mepraum.H24) > 1:
                if not mepraum.SH24 in SH24.keys():
                    SH24[mepraum.SH24] = []
                SH24[mepraum.SH24].append(mepraum)
            
            if float(mepraum.TZU) > 1:
                if not mepraum.STZU in STZU.keys():
                    STZU[mepraum.STZU] = []
                STZU[mepraum.STZU].append(mepraum)
            
            if float(mepraum.TAB) > 1:
                if not mepraum.STAB in STAB.keys():
                    STAB[mepraum.STAB] = []
                STAB[mepraum.STAB].append(mepraum)

            

            # Anlagen berechnen
            if not str(mepraum.ARZU) in ARZU.keys():
                ARZU[str(mepraum.ARZU)] = []
            ARZU[str(mepraum.ARZU)].append(mepraum)

            if not str(mepraum.ARAB) in ARAB.keys():
                ARAB[str(mepraum.ARAB)] = []
            ARAB[str(mepraum.ARAB)].append(mepraum)

            if not str(mepraum.AH24) in AH24.keys():
                AH24[str(mepraum.AH24)] = []
            AH24[str(mepraum.AH24)].append(mepraum)

            if not str(mepraum.ALAB) in ALAB.keys():
                ALAB[str(mepraum.ALAB)] = []
            ALAB[str(mepraum.ALAB)].append(mepraum)

            if not str(mepraum.ATAB) in ATAB.keys():
                ATAB[str(mepraum.ATAB)] = []
            ATAB[str(mepraum.ATAB)].append(mepraum)

            if not str(mepraum.ATZU) in ATZU.keys():
                ATZU[str(mepraum.ATZU)] = []
            ATZU[str(mepraum.ATZU)].append(mepraum)
        
# Schacht berechnen
with forms.ProgressBar(title='{value}/{max_value} Schächte',cancellable=True, step=1) as pb1:
    for n, schacht in enumerate(Schacht_liste):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(Schacht_liste))

        Liste_RAB = []
        Liste_LAB = []
        Liste_TAB = []
        Liste_H24 = []
        Liste_RZU = []
        Liste_TZU = []
        if schacht.Install_Schacht_Name in SRZU.keys():
            Liste_RZU = SRZU[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SRAB.keys():
            Liste_RAB = SRAB[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in STZU.keys():
            Liste_TZU = STZU[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in STAB.keys():
            Liste_TAB = STAB[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SLAB.keys():
            Liste_LAB = SLAB[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SH24.keys():
            Liste_H24 = SH24[schacht.Install_Schacht_Name]    
        
        if schacht.Install_Schacht_Name in SUB_RZU.keys():schacht.SUB_RZU_Schacht = SUB_RZU[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SUB_RAB.keys():schacht.SUB_RAB_Schacht = SUB_RAB[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SUB_TZU.keys():schacht.SUB_TZU_Schacht = SUB_TZU[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SUB_TAB.keys():schacht.SUB_TAB_Schacht = SUB_TAB[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SUB_LAB.keys():schacht.SUB_LAB_Schacht = SUB_LAB[schacht.Install_Schacht_Name]
        if schacht.Install_Schacht_Name in SUB_H24.keys():schacht.SUB_H24_Schacht = SUB_H24[schacht.Install_Schacht_Name]
        
        schacht.Raum_RZU = Liste_RZU
        schacht.Raum_RAB = Liste_RAB
        schacht.Raum_LAB = Liste_LAB
        schacht.Raum_H24 = Liste_H24
        schacht.Raum_TAB = Liste_TAB
        schacht.Raum_TZU = Liste_TZU
        schacht.Dict_Ebene = Ebene_nr_name
        schacht.Datenanalyse() 

for schacht in Schacht_liste:
    schacht.weiter_analyse() 
    
if forms.alert('Berechnete Werte in Schächte schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} Schächte',cancellable=True, step=1) as pb2:

        t = DB.Transaction(doc, "Luftmengen Schächte")
        t.Start()

        for n, schacht in enumerate(Schacht_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n + 1, len(Schacht_liste))
            schacht.werte_schreiben()
        t.Commit()

table_data_System = []
Systemen_liste = []

with forms.ProgressBar(title='{value}/{max_value} Luftkanal Systeme',cancellable=True, step=10) as pb1:
    for n, System_id in enumerate(systemen):
        if pb1.cancelled:
            script.exit()

        pb1.update_progress(n + 1, len(systemen))
        ductsystem = DuctSystem(System_id)
        Liste_RAB = []
        Liste_LAB = []
        Liste_TAB = []
        Liste_H24 = []
        Liste_RZU = []
        Liste_TZU = []
        if ductsystem.Anl_Nr in ARZU.keys():
            Liste_RZU = ARZU[ductsystem.Anl_Nr]
        if ductsystem.Anl_Nr in ARAB.keys():
            Liste_RAB = ARAB[ductsystem.Anl_Nr]
        if ductsystem.Anl_Nr in ATZU.keys():
            Liste_TZU = ATZU[ductsystem.Anl_Nr]
        if ductsystem.Anl_Nr in ATAB.keys():
            Liste_TAB = ATAB[ductsystem.Anl_Nr]
        if ductsystem.Anl_Nr in ALAB.keys():
            Liste_LAB = ALAB[ductsystem.Anl_Nr]
        if ductsystem.Anl_Nr in AH24.keys():
            Liste_H24 = AH24[ductsystem.Anl_Nr]  
        ductsystem.ARZU = Liste_RZU
        ductsystem.ARAB = Liste_RAB
        ductsystem.ALAB = Liste_LAB
        ductsystem.AH24 = Liste_H24
        ductsystem.ATAB = Liste_TAB
        ductsystem.ATZU = Liste_TZU
        ductsystem.Berechnen()
        Systemen_liste.append(ductsystem)

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Anlagen(Systeme) schreiben?', ok=False, yes=True, no=True):

    t1 = DB.Transaction(doc, "Luftmengen Anlagen")
    t1.Start()

    with forms.ProgressBar(title='{value}/{max_value} Luftkanal Systeme', cancellable=True, step=1) as pb2:
        for n, system in enumerate(Systemen_liste):
            if pb2.cancelled:
                t1.RollBack()
                script.exit()

            pb2.update_progress(n+1, len(Systemen_liste))
            system.werte_schreiben()

    t1.Commit()
