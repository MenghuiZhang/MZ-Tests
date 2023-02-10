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

[2022.11.21]
Version: 1.2
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
        self.Raum_HAB = []
        self.Raum_TAB = []
        self.Raum_TZU = []
        self.Dict_Ebene = {}
        
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
            ['HAB', 'IGF_RLT_Luftmenge_24h'],
            ['LAB', 'IGF_RLT_Luftmenge_LAB'],
            ['TAB', 'IGF_RLT_Luftmenge_max_TAB'],
            ['TZU', 'IGF_RLT_Luftmenge_max_TZU'],

             # Anlagen
            ['ARAB', 'IGF_RLT_AnlagenNr_RAB'],
            ['ARZU', 'IGF_RLT_AnlagenNr_RZU'],
            ['AHAB', 'IGF_RLT_AnlagenNr_24h'],
            ['ATAB', 'IGF_RLT_AnlagenNr_TAB'],
            ['ATZU', 'IGF_RLT_AnlagenNr_TZU'],
            ['ALAB', 'IGF_RLT_AnlagenNr_LAB'],

            # Schacht
            ['SRZU', 'TGA_RLT_SchachtZuluft'],
            ['SRAB', 'TGA_RLT_SchachtAbluft'],
            ['SHAB', 'TGA_RLT_Schacht24hAbluft'],
            ['STAB', 'IGF_RLT_Schacht_TAB'],
            ['STZU', 'IGF_RLT_Schacht_TZU'],
            ['SLAB', 'IGF_RLT_Schacht_LAB'],

            ['fk_Reduzieren','IGF_RLT_Schacht-ReduziertFaktor']
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(revit_name))
        
        # self.fk_Reduzieren = 1
        if not self.fk_Reduzieren or self.fk_Reduzieren == 0:self.fk_Reduzieren = 1
        
        # Prüfung
        if self.RAB > 0:
            if not self.ARAB:
                logger.error("Raumabluft-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
            if not self.SRAB:
                logger.error("Raumabluft-Schacht in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
        else:
            self.RAB = 0

        if self.RZU > 0:
            if not self.ARZU:
                logger.error("Raumzuluft-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
            if not self.SRZU:
                logger.error("Raumzuluft-Schacht in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
        else:
            self.RZU = 0

        if self.HAB > 0:
            if not self.AHAB:
                logger.error("24h-Abluft-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
            if not self.SHAB:
                logger.error("24h-Abluft-Schacht in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
        else:
            self.HAB = 0
        
        if self.LAB > 0:
            if not self.ALAB:
                logger.error("Laborabluft-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
            if not self.SLAB:
                logger.error("Laborabluft-Schacht {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
        else:
            self.LAB = 0

        if self.TZU > 0:
            if not self.ATZU:
                logger.error("TierhaltungZuluft-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
            if not self.STZU:
                logger.error("TierhaltungZuluft-Schacht in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
        else:
            self.TZU = 0

        if self.TAB > 0:
            if not self.ATAB:
                logger.error("TierhaltungAbluft-Anlage-Nummer in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
            if not self.STAB:
                logger.error("TierhaltungAbluft-Schacht in Raum {}-{} nicht gefunden, Ebene: {}, ElementId: {}".format(self.nummer,self.name,self.ebene,self.element_id.ToString()))
        else:
            self.TAB = 0
        
        self.SRZU_Menge = ''
        self.SRZU_Verteilen = ''
        self.SRAB_Menge = ''
        self.SRAB_Verteilen = ''
        self.SHAB_Menge = ''
        self.SHAB_Verteilen = ''
        self.SLAB_Menge = ''
        self.SLAB_Verteilen = ''

        self.STZU_Menge = ''
        # self.STZU_Verteilen = ''
        self.STAB_Menge = ''
        # self.STAB_Verteilen = ''
        

        self.SRZUR_Menge = ''
        self.SRABR_Menge = ''
        self.SLABR_Menge = ''
        self.SHABR_Menge = ''
    
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
        wert_schreiben("IGF_RLT_Schacht_24h_Menge", self.SHAB_Menge)
        wert_schreiben("IGF_RLT_Schacht_LAB_Menge", self.SLAB_Menge)
        wert_schreiben("IGF_RLT_Schacht_TZU_Menge", self.STZU_Menge)
        wert_schreiben("IGF_RLT_Schacht_TAB_Menge", self.STAB_Menge)

        wert_schreiben("IGF_RLT_Verteilung_LAB", self.SLAB_Verteilen)
        wert_schreiben("IGF_RLT_VerteilungZuluft", self.SRZU_Verteilen)
        wert_schreiben("IGF_RLT_VerteilungAbluft", self.SRAB_Verteilen)
        wert_schreiben("IGF_RLT_Verteilung24hAbluft", self.SHAB_Verteilen)

        wert_schreiben("IGF_RLT_Schacht-Reduziert_RZU", self.SRZUR_Menge)
        wert_schreiben("IGF_RLT_Schacht-Reduziert_RAB", self.SRABR_Menge)
        wert_schreiben("IGF_RLT_Schacht-Reduziert_24h", self.SHABR_Menge)
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
                    # Text += 'Anl ' + str(anl) + ', '
            return Text[:-2]

        Dict_RZU = {}
        RZU = 0
        Dict_RAB = {}
        RAB = 0
        Dict_HAB = {}
        HAB = 0
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

        for el in self.Raum_HAB:
            HAB += el.HAB
            if el.ebene_nr not in Dict_HAB.keys():
                Dict_HAB[el.ebene_nr] = {}
            if el.AHAB not in Dict_HAB[el.ebene_nr].keys():
                Dict_HAB[el.ebene_nr][el.AHAB] = 0
            Dict_HAB[el.ebene_nr][el.AHAB] += el.HAB * self.fk_Reduzieren
        
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
        self.SHAB_Menge = HAB
        self.STAB_Menge = TAB
        self.STZU_Menge = TZU

        self.SRZUR_Menge = RZU * self.fk_Reduzieren
        self.SRABR_Menge = RAB * self.fk_Reduzieren
        self.SLABR_Menge = LAB * self.fk_Reduzieren
        self.SHABR_Menge = HAB * self.fk_Reduzieren

        self.SRZU_Verteilen = Textbearbeiten(Dict_RZU)
        self.SRAB_Verteilen = Textbearbeiten(Dict_RAB)
        self.SHAB_Verteilen = Textbearbeiten(Dict_HAB)
        self.SLAB_Verteilen = Textbearbeiten(Dict_LAB)

    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error("Parameter ({}) konnte nicht gefunden werden".format(param_name))
            return

        return get_value(param)

    # def table_row_Scahcht(self):
    #     return [self.Install_Schacht_Name,self.nummer,self.name, self.ZuSchacht, self.ZuSchachtV, 
    #             self.AbSchacht, self.AbSchachtV, self.Ab24hSchacht,self.Ab24hSchachtV]


SRZU = {}
SRAB = {}
SLAB = {}
S24H = {}



Schacht_liste = []
# table_data_Schacht = []
Schacht_Daten = {}
Ebene_nr_name = {}

Anlagen_Liste = {}

# Zuluft_Liste = {}
# Abluft_Liste = {}
# Abluft_24h_Liste = {}
SRZU = {}
SRAB = {}
SLAB = {}
SHAB = {}
STZU = {}
STAB = {}
ARZU = {}
ARAB = {}
ALAB = {}
AHAB = {}
ATAB = {}
ATZU = {}

with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=5) as pb:
    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        # Schacht berechnen
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

        if float(mepraum.HAB) > 1:
            if not mepraum.SHAB in SHAB.keys():
                SHAB[mepraum.SHAB] = []
            SHAB[mepraum.SHAB].append(mepraum)
        
        if float(mepraum.TZU) > 1:
            if not mepraum.STZU in STZU.keys():
                STZU[mepraum.STZU] = []
            STZU[mepraum.STZU].append(mepraum)
        
        if float(mepraum.TAB) > 1:
            if not mepraum.STAB in STAB.keys():
                STAB[mepraum.STAB] = []
            STAB[mepraum.STAB].append(mepraum)

        if mepraum.Install_Schacht:
            Schacht_liste.append(mepraum)

        # Anlagen berechnen
        if not str(mepraum.ARZU) in ARZU.keys():
            ARZU[str(mepraum.ARZU)] = []
        ARZU[str(mepraum.ARZU)].append(mepraum)

        if not str(mepraum.ARAB) in ARAB.keys():
            ARAB[str(mepraum.ARAB)] = []
        ARAB[str(mepraum.ARAB)].append(mepraum)

        if not str(mepraum.AHAB) in AHAB.keys():
            AHAB[str(mepraum.AHAB)] = []
        AHAB[str(mepraum.AHAB)].append(mepraum)

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
        Liste_HAB = []
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
        if schacht.Install_Schacht_Name in SHAB.keys():
            Liste_HAB = SHAB[schacht.Install_Schacht_Name]    
        schacht.Raum_RZU = Liste_RZU
        schacht.Raum_RAB = Liste_RAB
        schacht.Raum_LAB = Liste_LAB
        schacht.Raum_HAB = Liste_HAB
        schacht.Raum_TAB = Liste_TAB
        schacht.Raum_TZU = Liste_TZU
        schacht.Dict_Ebene = Ebene_nr_name
        schacht.Datenanalyse() 
            
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

class DuctSystem:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.ARZU = []
        self.ARAB = []
        self.ALAB = []
        self.AHAB = []
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
        self.zu_geraete = 0
        self.zu_geraete_p = 0
        self.ab_geraete = 0
        self.ab_geraete_p = 0
        self.ab_geraete24h = 0
        self.ab_geraete24h_p = 0

        
        # if self.Anl_Nr in Anlagen_Liste.keys():
        #     self.zuluft = Anlagen_Liste[self.Anl_Nr][0]
        #     self.abluft = Anlagen_Liste[self.Anl_Nr][1]
        #     self.ab24h = Anlagen_Liste[self.Anl_Nr][2]
        #     self.zu_geraete = round(self.zuluft / int(self.Ger_Anzahl),1)
        #     self.zu_geraete_p = round(self.zuluft / int(self.Anl_Pro_Anz),1)
        #     self.ab_geraete = round(self.abluft / int(self.Ger_Anzahl),1)
        #     self.ab_geraete_p = round(self.abluft / int(self.Anl_Pro_Anz),1)
        #     self.ab_geraete24h = round(self.ab24h / int(self.Ger_Anzahl),1)
        #     self.ab_geraete24h_p = round(self.ab24h / int(self.Anl_Pro_Anz),1)

        # if self.Anl_Nr in Abluft_Liste.keys():
        #     self.abluft = Abluft_Liste[self.Anl_Nr]
        #     self.ab_geraete = round(self.abluft / int(self.Ger_Anzahl),1)
        #     self.ab_geraete_p = round(self.abluft / int(self.Anl_Pro_Anz),1)
            

        # if self.Anl_Nr in Abluft_24h_Liste.keys():
        #     self.ab24h = Abluft_24h_Liste[self.Anl_Nr]
        #     self.ab_geraete24h = round(self.ab24h / int(self.Ger_Anzahl),1)
        #     self.ab_geraete24h_p = round(self.ab24h / int(self.Anl_Pro_Anz),1)

    def Datenanalyse(self):
        RZU = 0
        RAB = 0
        TZU = 0
        TAB = 0
        LAB = 0
        HAB = 0
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
        
        for el in self.AHAB:
            HAB += el.HAB

        self.zuluft = RZU
        self.abluft = RAB+LAB
        self.ab24h = HAB

    def Berechnen(self):
        self.Datenanalyse()
        self.zu_geraete = round(self.zuluft / int(self.Ger_Anzahl),1)
        self.zu_geraete_p = round(self.zuluft / int(self.Anl_Pro_Anz),1)
        self.ab_geraete = round(self.abluft / int(self.Ger_Anzahl),1)
        self.ab_geraete_p = round(self.abluft / int(self.Anl_Pro_Anz),1)
        self.ab_geraete24h = round(self.ab24h / int(self.Ger_Anzahl),1)
        self.ab_geraete24h_p = round(self.ab24h / int(self.Anl_Pro_Anz),1)

    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter ({}) konnte nicht gefunden werden".format(param_name))
            return

        return get_value(param)

    # def table_row(self):
    #     """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
    #     return [
    #         self.name,
    #         self.Anl_Nr,
    #         self.Ger_Anzahl,
    #         self.Anl_Pro_Anz,
    #         self.zu_geraete,
    #         self.zu_geraete_p,
    #         self.ab_geraete,
    #         self.ab_geraete_p,
    #         self.ab_geraete24h,
    #         self.ab_geraete24h_p
    #     ]

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
        Liste_HAB = []
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
        if ductsystem.Anl_Nr in AHAB.keys():
            Liste_HAB = AHAB[ductsystem.Anl_Nr]  
        ductsystem.ARZU = Liste_RZU
        ductsystem.ARAB = Liste_RAB
        ductsystem.ALAB = Liste_LAB
        ductsystem.AHAB = Liste_HAB
        ductsystem.ATAB = Liste_TAB
        ductsystem.ATZU = Liste_TZU
        ductsystem.Berechnen()
        Systemen_liste.append(ductsystem)
        # table_data_System.append(ductsystem.table_row())

# Sorteiren nach Anlagennummer und dann Gerätenummer
# table_data_System.sort()


# output.print_table(
#     table_data=table_data_System,
#     title="Luftkanal Systeme",
#     columns=[ 'Systemname', 'Anl. Nr', 'Ger. Anzahl', 'Anl_Pro_Anz', 'Zuluft', 
#     'Zuluft Prozentual', 'Abluft', 'Abluft Prozentual', '24h-Abluft', '24h-Abluft Prozentual' ]
# )


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

# total = time.time() - start
# logger.info("total time: {} {}".format(total, 100 * "_"))