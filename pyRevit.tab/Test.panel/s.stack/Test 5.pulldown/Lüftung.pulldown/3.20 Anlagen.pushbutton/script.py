# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit,DB
from pyrevit import script, forms
from IGF_log import getlog
from IGF_lib import get_value
import time

start = time.time()


__title__ = "3.2 Raumluft Schächte&Anlagen"
__doc__ = """Anlagenberechnung

[01.10.2021]
Version: 1.2"""
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
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

# Systemen aus Projekt
System_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
systemen = System_collector.ToElementIds()
System_collector.Dispose()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
logger.info("{} Luftkanalsystemen ausgewählt".format(len(systemen)))

if not (spaces and systemen):
    logger.error("Keine MEP Räume/Luftkanalsystemen in aktueller Projekt gefunden")
    script.exit()


class MEPRaum:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.ZuSchacht = ''
        self.ZuSchachtV = ''
        self.AbSchacht = ''
        self.AbSchachtV = ''
        self.Ab24hSchacht = ''
        self.Ab24hSchachtV = ''
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['ebene', 'IGF_RLT_Verteilung_EbenenName'],
            ['ebene_nr', 'IGF_RLT_Verteilung_EbenenSortieren'],

            # Anlagen
            ['Zu_Anl_Nr', 'TGA_RLT_AnlagenNrZuluft'],
            ['Ab_Anl_Nr', 'TGA_RLT_AnlagenNrAbluft'],
            ['Ab_Anl_Nr_24h', 'TGA_RLT_AnlagenNr24hAbluft'],

            #Schacht
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],

            # Luftmengen
            ['ZU', 'IGF_RLT_AnlagenRaumZuluft'],
            ['AB', 'IGF_RLT_AnlagenRaumAbluft'],
            ['AB24H', 'IGF_RLT_AnlagenRaum24hAbluft'],
            # Schacht
            ['ZU_S', 'TGA_RLT_SchachtZuluft'],
            ['AB_S', 'TGA_RLT_SchachtAbluft'],
            ['AB24H_S', 'TGA_RLT_Schacht24hAbluft']
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(revit_name))


        # Prüfung
        if self.ZU > 0:
            if not self.Zu_Anl_Nr:
                logger.error("Zuluft-Anlage-Nummer in Raum {} nicht gefunden".format(self.nummer))
            if not self.ZU_S:
                logger.error("Zuluft-Schacht {} nicht gefunden".format(self.nummer))
        if self.AB > 0:
            if not self.Ab_Anl_Nr:
                logger.error("Abluft-Anlage-Nummer in Raum {} nicht gefunden".format(self.nummer))
            if not self.AB_S:
                logger.error("Abluft-Schacht in Raum {} nicht gefunden".format(self.nummer))
        if self.AB24H > 0:
            if not self.Ab_Anl_Nr_24h:
                logger.error("24h-Abluft-Anlage-Nummer in Raum {} nicht gefunden".format(self.nummer))
            if not self.AB24H_S:
                logger.error("24h-Abluft-Schacht in Raum {} nicht gefunden".format(self.nummer))

    @property
    def ZuSchachtV(self):
        return self._ZuSchachtV
    @ZuSchachtV.setter
    def ZuSchachtV(self,value):
        self._ZuSchachtV = value
    @property
    def AbSchachtV(self):
        return self._AbSchachtV
    @AbSchachtV.setter
    def AbSchachtV(self,value):
        self._AbSchachtV = value
    @property
    def Ab24hSchachtV(self):
        return self._Ab24hSchachtV
    @Ab24hSchachtV.setter
    def Ab24hSchachtV(self,value):
        self._Ab24hSchachtV = value
    @property
    def ZuSchacht(self):
        return self._ZuSchacht
    @ZuSchacht.setter
    def ZuSchacht(self,value):
        self._ZuSchacht = value
    @property
    def AbSchacht(self):
        return self._AbSchacht
    @AbSchacht.setter
    def AbSchacht(self,value):
        self._AbSchacht = value
    @property
    def Ab24hSchacht(self):
        return self._Ab24hSchacht
    @Ab24hSchacht.setter
    def Ab24hSchacht(self,value):
        self._Ab24hSchacht = value
    
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

        wert_schreiben("TGA_RLT_SchachtZuluftMenge", self.ZuSchacht)
        wert_schreiben("TGA_RLT_SchachtAbluftMenge", self.AbSchacht)
        wert_schreiben("TGA_RLT_Schacht24hAbluftMenge", self.Ab24hSchacht)
        wert_schreiben("IGF_RLT_VerteilungZuluft", self.ZuSchachtV)
        wert_schreiben("IGF_RLT_VerteilungAbluft", self.AbSchachtV)
        wert_schreiben("IGF_RLT_Verteilung24hAbluft", self.Ab24hSchachtV)


    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error("Parameter ({}) konnte nicht gefunden werden".format(param_name))
            return

        return get_value(param)

    def table_row_Scahcht(self):
        return [self.Install_Schacht_Name,self.nummer,self.name, self.ZuSchacht, self.ZuSchachtV, 
                self.AbSchacht, self.AbSchachtV, self.Ab24hSchacht,self.Ab24hSchachtV]

Schacht_liste = []
table_data_Schacht = []
Schacht_Daten = {}
Ebene_nr_name = {}

Anlagen_Liste = {}

# Zuluft_Liste = {}
# Abluft_Liste = {}
# Abluft_24h_Liste = {}

with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=5) as pb:
    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        # Schacht berechnen
        if not mepraum.ebene_nr in Ebene_nr_name.keys():
            Ebene_nr_name[mepraum.ebene_nr] = mepraum.ebene

        # Zuluft Schacht
        if float(mepraum.ZU) > 1:
            if not mepraum.ZU_S in Schacht_Daten.keys():
                Schacht_Daten[mepraum.ZU_S] = {'0':{},'1':{},'2':{}, '3':float(mepraum.ZU),'4':0,'5':0}
            else:
                Schacht_Daten[mepraum.ZU_S]['3'] += float(mepraum.ZU)

            if not mepraum.ebene_nr in Schacht_Daten[mepraum.ZU_S]['0'].keys():
                Schacht_Daten[mepraum.ZU_S]['0'][mepraum.ebene_nr] = {}
            if not mepraum.Zu_Anl_Nr in Schacht_Daten[mepraum.ZU_S]['0'][mepraum.ebene_nr].keys():
                Schacht_Daten[mepraum.ZU_S]['0'][mepraum.ebene_nr][mepraum.Zu_Anl_Nr] = float(mepraum.ZU)
            else:
                Schacht_Daten[mepraum.ZU_S]['0'][mepraum.ebene_nr][mepraum.Zu_Anl_Nr] += float(mepraum.ZU)
        
        # Abluft Schacht
        if float(mepraum.AB) > 1:
            if not mepraum.AB_S in Schacht_Daten.keys():
                Schacht_Daten[mepraum.AB_S] = {'0':{},'1':{},'2':{},'3':0,'4':float(mepraum.AB),'5':0}
            else:
                Schacht_Daten[mepraum.AB_S]['4'] += float(mepraum.AB)
            if not mepraum.ebene_nr in Schacht_Daten[mepraum.AB_S]['1'].keys():
                Schacht_Daten[mepraum.AB_S]['1'][mepraum.ebene_nr] = {}
            if not mepraum.Ab_Anl_Nr in Schacht_Daten[mepraum.AB_S]['1'][mepraum.ebene_nr].keys():
                Schacht_Daten[mepraum.AB_S]['1'][mepraum.ebene_nr][mepraum.Ab_Anl_Nr] = float(mepraum.AB)
            else:
                Schacht_Daten[mepraum.AB_S]['1'][mepraum.ebene_nr][mepraum.Ab_Anl_Nr] += float(mepraum.AB)

        # 24h Abluft Schacht
        if float(mepraum.AB24H) > 1:
            if not mepraum.AB24H_S in Schacht_Daten.keys():
                Schacht_Daten[mepraum.AB24H_S] = {'0':{},'1':{},'2':{},'3':0,'4':0,'5':float(mepraum.AB24H)}
            else:
                Schacht_Daten[mepraum.AB24H_S]['5'] += float(mepraum.AB24H)
            if not mepraum.ebene_nr in Schacht_Daten[mepraum.AB24H_S]['2'].keys():
                Schacht_Daten[mepraum.AB24H_S]['2'][mepraum.ebene_nr] = {}
            if not mepraum.Ab_Anl_Nr_24h in Schacht_Daten[mepraum.AB24H_S]['2'][mepraum.ebene_nr].keys():
                Schacht_Daten[mepraum.AB24H_S]['2'][mepraum.ebene_nr][mepraum.Ab_Anl_Nr_24h] = float(mepraum.AB24H)
            else:
                Schacht_Daten[mepraum.AB24H_S]['2'][mepraum.ebene_nr][mepraum.Ab_Anl_Nr_24h] += float(mepraum.AB24H)

        if mepraum.Install_Schacht:
            Schacht_liste.append(mepraum)

        # Anlagen berechnen
        if not mepraum.Zu_Anl_Nr in Anlagen_Liste.keys():
            Anlagen_Liste[mepraum.Zu_Anl_Nr] = [float(mepraum.ZU),0,0]
        else:
            Anlagen_Liste[mepraum.Zu_Anl_Nr][0] += float(mepraum.ZU)

        if not mepraum.Ab_Anl_Nr in Anlagen_Liste.keys():
            Anlagen_Liste[mepraum.Ab_Anl_Nr] = [0,float(mepraum.AB),0]
        else:
            Anlagen_Liste[mepraum.Ab_Anl_Nr][1] += float(mepraum.AB)

        if not mepraum.Ab_Anl_Nr_24h in Anlagen_Liste.keys():
            Anlagen_Liste[mepraum.Ab_Anl_Nr_24h] = [0,0,float(mepraum.AB24H)]
        else:
            Anlagen_Liste[mepraum.Ab_Anl_Nr_24h][2] += float(mepraum.AB24H)

# Schacht berechnen
with forms.ProgressBar(title='{value}/{max_value} Schächte',cancellable=True, step=1) as pb1:
    for n, schacht in enumerate(Schacht_liste):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(Schacht_liste))

        Schacht_zu_V = ''
        Schacht_ab_V = ''
        Schacht_ab24h_V = ''

        if schacht.Install_Schacht_Name in Schacht_Daten.keys():
            Zu_Dict = Schacht_Daten[schacht.Install_Schacht_Name]['0']
            Ab_Dict = Schacht_Daten[schacht.Install_Schacht_Name]['1']
            Ab24h_Dict = Schacht_Daten[schacht.Install_Schacht_Name]['2']
            schacht.ZuSchacht = round(Schacht_Daten[schacht.Install_Schacht_Name]['3'],1)
            schacht.AbSchacht = round(Schacht_Daten[schacht.Install_Schacht_Name]['4'],1)
            schacht.Ab24hSchacht = round(Schacht_Daten[schacht.Install_Schacht_Name]['5'],1)
            sorted(Zu_Dict)
            sorted(Ab_Dict)
            sorted(Ab24h_Dict)
            for eb in Zu_Dict.keys():
                Schacht_zu_V += str(Ebene_nr_name[eb]) + ': '
                Zu_Dict_eb = Zu_Dict[eb]
                zu_keys = Zu_Dict_eb.keys()[:]
                zu_keys.sort()
                for anl in zu_keys:
                    Schacht_zu_V += 'Anl ' + str(anl) + '=' + str(int(round(Zu_Dict_eb[anl]))) + ', '
            if Schacht_zu_V:
                schacht.ZuSchachtV = '[m3/h] - ' + Schacht_zu_V[:-2]
            
            
            for eb in Ab_Dict.keys():
                Schacht_ab_V += str(Ebene_nr_name[eb]) + ': '
                Ab_Dict_eb = Ab_Dict[eb]
                ab_keys = Ab_Dict_eb.keys()[:]
                ab_keys.sort()
                for anl in ab_keys:
                    Schacht_ab_V += 'Anl ' + str(anl) + '=' + str(int(round(Ab_Dict_eb[anl]))) + ', '
            if Schacht_ab_V:
                schacht.AbSchachtV = '[m3/h] - ' + Schacht_ab_V[:-2]
            
            for eb in Ab24h_Dict.keys():
                Schacht_ab24h_V += str(Ebene_nr_name[eb]) + ': '
                Ab24h_Dict_eb = Ab24h_Dict[eb]
                ab24h_keys = Ab24h_Dict_eb.keys()[:]
                ab24h_keys.sort()
                for anl in ab24h_keys:
                    Schacht_ab24h_V += 'Anl ' + str(anl) + '=' + str(int(round(Ab24h_Dict_eb[anl]))) + ', '
            if Schacht_ab24h_V:
                schacht.Ab24hSchachtV = '[m3/h] - ' + Schacht_ab24h_V[:-2]
        
        table_data_Schacht.append(schacht.table_row_Scahcht())
table_data_Schacht.sort()

output.print_table(
    table_data=table_data_Schacht,
    title="Luftmengen Verteilung",
    columns=['Schachtname','Nummer', 'Name', 'Zuluft','Zuluft Verteilung','Abluft', 'Abluft Verteilung','24h Abluft', '24h Abluft Verteilung']
)

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

        attr = [
            ['name', 'Systemname'],
            ['Ger_Nr', 'TGA_RLT_AnlagenGeräteNr'],
            ['Ger_Anzahl', 'TGA_RLT_AnlagenGeräteAnzahl'],
            ['Anl_Nr', 'TGA_RLT_AnlagenNr'],
            ['Anl_Pro_Anz', 'TGA_RLT_AnlagenProzentualAnzahl'],
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
        
        if self.Anl_Nr in Anlagen_Liste.keys():
            self.zuluft = Anlagen_Liste[self.Anl_Nr][0]
            self.abluft = Anlagen_Liste[self.Anl_Nr][1]
            self.ab24h = Anlagen_Liste[self.Anl_Nr][2]
            self.zu_geraete = round(self.zuluft / int(self.Ger_Anzahl),1)
            self.zu_geraete_p = round(self.zuluft / int(self.Anl_Pro_Anz),1)
            self.ab_geraete = round(self.abluft / int(self.Ger_Anzahl),1)
            self.ab_geraete_p = round(self.abluft / int(self.Anl_Pro_Anz),1)
            self.ab_geraete24h = round(self.ab24h / int(self.Ger_Anzahl),1)
            self.ab_geraete24h_p = round(self.ab24h / int(self.Anl_Pro_Anz),1)

        # if self.Anl_Nr in Abluft_Liste.keys():
        #     self.abluft = Abluft_Liste[self.Anl_Nr]
        #     self.ab_geraete = round(self.abluft / int(self.Ger_Anzahl),1)
        #     self.ab_geraete_p = round(self.abluft / int(self.Anl_Pro_Anz),1)
            

        # if self.Anl_Nr in Abluft_24h_Liste.keys():
        #     self.ab24h = Abluft_24h_Liste[self.Anl_Nr]
        #     self.ab_geraete24h = round(self.ab24h / int(self.Ger_Anzahl),1)
        #     self.ab_geraete24h_p = round(self.ab24h / int(self.Anl_Pro_Anz),1)


    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter ({}) konnte nicht gefunden werden".format(param_name))
            return

        return get_value(param)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.Anl_Nr,
            self.Ger_Nr,
            self.name,
            self.Ger_Anzahl,
            self.Anl_Pro_Anz,
            self.zu_geraete,
            self.zu_geraete_p,
            self.ab_geraete,
            self.ab_geraete_p,
            self.ab_geraete24h,
            self.ab_geraete24h_p
        ]

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
        
        wert_schreiben("TGA_RLT_AnlagenZuMenge", self.zu_geraete)
        wert_schreiben("TGA_RLT_AnlagenProzentualZuMenge", self.zu_geraete_p)
        wert_schreiben("TGA_RLT_AnlagenAbMenge", self.ab_geraete)
        wert_schreiben("TGA_RLT_AnlagenProzentualAbMenge", self.ab_geraete_p)
        wert_schreiben("TGA_RLT_Anlagen24hAbMenge", self.ab_geraete24h)
        wert_schreiben("TGA_RLT_AnlagenProzentual24hAbMenge", self.ab_geraete24h_p)


table_data_System = []
Systemen_liste = []

with forms.ProgressBar(title='{value}/{max_value} Luftkanal Systeme',cancellable=True, step=10) as pb1:
    for n, System_id in enumerate(systemen):
        if pb1.cancelled:
            script.exit()

        pb1.update_progress(n + 1, len(systemen))
        ductsystem = DuctSystem(System_id)

        Systemen_liste.append(ductsystem)
        table_data_System.append(ductsystem.table_row())

# Sorteiren nach Anlagennummer und dann Gerätenummer
table_data_System.sort()


output.print_table(
    table_data=table_data_System,
    title="Luftkanal Systeme",
    columns=[ 'Anl. Nr','Ger. Nr', 'Systemname', 'Ger. Anzahl', 'Anl_Pro_Anz', 'Zuluft', 
    'Zuluft Prozentual', 'Abluft', 'Abluft Prozentual', '24h-Abluft', '24h-Abluft Prozentual' ]
)


# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Anlagen(Systeme) schreiben?', ok=False, yes=True, no=True):

    t1 = DB.Transaction(doc, "Luftmengen Anlagen")
    t1.Start()

    with forms.ProgressBar(title='{value}/{max_value} Luftkanal Systeme', cancellable=True, step=1) as pb2:
        for n, system in enumerate(Systemen_liste):
            if pb2.cancelled:
                t1.rollBack()
                script.exit()

            pb2.update_progress(n+1, len(Systemen_liste))
            system.werte_schreiben()

    t1.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))