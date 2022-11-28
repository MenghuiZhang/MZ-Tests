# coding: utf8
from cmath import phase
import clr
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit,DB
from Autodesk.Revit.UI import ExternalEvent
from pyrevit import script, forms
from IGF_log import getlog
from IGF_lib import get_value
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility
from eventhandler import ResizeAnsicht,ResetAnsicht


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
active_view = doc.ActiveView

views = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.View3D)).ToElements()
for el in views:
    if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
        uidoc.ActiveView = el
        break

dict_Luftauslass = {}
dict_ueberstrom = {}
dict_vsr = {}
liste_mepraum = []
PHASE = doc.Phases[0]
DICT_UEBERSTROM = {} ## MEP: 'Ein' ..., 'Aus'...
ELEMID_UEBER = [] ## ElememntId Überstrom
DICT_LUFTAUSLASS = {} ## MEP: OB(Auslässe)
DICT_EINBAUTEILE = {} ## MEP: [VSRID]
DICT_MEP_AUSLASS = {} ## VSR: [MEPID,MEPID]
DICT_VSR_AUSLASS = {} ## VSRID: OB(Auslässe)
DICT_VSR = {} ## VSRID: Class VSR

def Muster_Pruefen(elem):
    try:
        return elem.LookupParameter('Bearbeitungsbereich').AsValueString() == 'KG4xx_Musterbereich'
    except:
        return False



# Überstrom
Baugruppen_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType()
Baugruppen = Baugruppen_collector.ToElementIds()
Baugruppen_collector.Dispose()

class UeberStromAuslass(object):
    def __init__(self,elem,vol):
        self.elem = elem
        self.elem_id = elem.Id
        self.menge = vol
        self.conns = self.get_connector()
        self.typ = self.get_typ()
        self.raum = ''
        self.raumid = ''
        self.anderesraum = ""
        self.get_einbauort()
        try:
            self.familieundtyp = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
        except:
            self.familieundtyp = ''
        

    def get_connector(self):
        return list(self.elem.MEPModel.ConnectorManager.Connectors)

    def get_typ(self):
        conn = self.conns[0]
        if conn.Direction.ToString() == 'Out':
            return 'Aus'
        elif conn.Direction.ToString() == 'In':
            return 'Ein'

    def get_einbauort(self):
        try:
            self.raum = self.elem.Space[PHASE].Number + ' - ' + self.elem.Space[PHASE].LookupParameter('Name').AsString()
            self.raumid = self.elem.Space[PHASE].Id.ToString()
        except:
            if not Muster_Pruefen(self.elem):
                logger.error('Überstrom: Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elem.Id.ToString()))
            return

class Baugruppe:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.auslaesse = self.get_Auslass()
        if self.Pruefen():
            self.Volumen = get_value(self.elem.LookupParameter('IGF_RLT_Überströmung'))
            self.Eingang = ''
            self.Ausgang = ''
            try:
                self.Analyse()
            except Exception as e:
                logger.error(e)

            
    def get_Auslass(self):
        auslass_liste = []
        for elemid in self.elem.GetMemberIds():
            elem = doc.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008013' and elem.Symbol.FamilyName.upper().find("ÜBER") != -1:
                auslass_liste.append(elem)
        return auslass_liste
    
    def Pruefen(self):
        return len(self.auslaesse) == 2
   
    
    def Analyse(self):
        for elem in self.auslaesse:
            auslass = UeberStromAuslass(elem,self.Volumen)
            if auslass.typ == 'Aus':
                self.Ausgang = auslass
            elif auslass.typ == 'Ein':
                self.Eingang = auslass
        self.Ausgang.anderesraum = self.Eingang.raum
        self.Eingang.anderesraum = self.Ausgang.raum


with forms.ProgressBar(title = "{value}/{max_value} Überströmungsbaugruppen",cancellable=True, step=10) as pb:
    for n, BGid in enumerate(Baugruppen):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(Baugruppen))
        baugruppe = Baugruppe(BGid)
        if not baugruppe.Pruefen():
            continue
        #print(baugruppe.auslaesse)
        if not baugruppe.Eingang.raumid in DICT_UEBERSTROM.keys():
            DICT_UEBERSTROM[baugruppe.Eingang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        if not baugruppe.Ausgang.raumid in DICT_UEBERSTROM.keys():
            DICT_UEBERSTROM[baugruppe.Ausgang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        DICT_UEBERSTROM[baugruppe.Eingang.raumid][baugruppe.Eingang.typ].Add(baugruppe.Eingang)
        DICT_UEBERSTROM[baugruppe.Ausgang.raumid][baugruppe.Ausgang.typ].Add(baugruppe.Ausgang)
        ELEMID_UEBER.append(baugruppe.Eingang.elem.Id.ToString())
        ELEMID_UEBER.append(baugruppe.Ausgang.elem.Id.ToString())


Ductterminal_coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
Ductterminalids = Ductterminal_coll.ToElementIds()
Ductterminal_coll.Dispose()

class Luftauslass(object):
    def __init__(self,elementid):
        self.elem_id = elementid
        self.elem = doc.GetElement(self.elem_id)
        self.RoutingListe = []
        self.size = self.get_Size()
        self.vsr = ''
        self.Einbauteile = []
        self.Luftmengenmin = 0
        self.Luftmengenmax = 0
        self.Luftmengennacht = 0
        self.Luftmengentiefe = 0
        self.VSR_Class = None
        self.art = ''
        
        try:self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:self.System = ''
        self.Luftmengenermitteln()
        self.get_RountingListe(self.elem)
        self.einbauort = ''
        self.raumnr = ''
        self.get_Einbauort()
        
        try:self.familieundtyp = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()   
        except:self.familieundtyp = ''

        try:self.Typen = self.elem.Symbol.LookupParameter('Typenkommentare').AsString()
        except:self.Typen = ''

        self.get_Art()

    def Luftmengenermitteln(self):
        try:
            self.Luftmengenmin = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin')))
        except:
            pass
        try:
            self.Luftmengenmax = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax')))
        except:
            pass
        try:
            self.Luftmengennacht = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht')))
        except:
            pass
        try:
            self.Luftmengentiefe = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht')))
        except:
            pass
    def get_Einbauort(self):
        try:
            self.einbauort = self.elem.Space[PHASE].Id.ToString()
            self.raumnr = self.elem.Space[PHASE].Number + ' - ' + self.elem.Space[PHASE].LookupParameter('Name').AsString()
        except:
            if not Muster_Pruefen(self.elem):
                logger.error('Luftauslass: Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elem_id.ToString()))
            return

    def get_RountingListe(self,element):
        if self.vsr:
            return
        elemid = element.Id.ToString()
        self.RoutingListe.append(elemid)
        cate = element.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Luftkanaldämmung außen', 'Luftkanaldämmung innen', 'Rohre', 'Rohrformteile', 'Rohrzubehör','Rohrdämmung']:
            conns = None
            try:conns = element.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = element.ConnectorManager.Connectors
                except:pass
            if conns:
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if not owner.Id.ToString() in self.RoutingListe:
                            if owner.Category.Name == 'Luftkanalzubehör':
                                faminame = owner.Symbol.FamilyName
                                if faminame.upper().find('VOLUMENSTROMREGLER') != -1 or faminame.upper().find('VSR') != -1:
                                    self.vsr = owner.Id.ToString()
                                    return
                            self.get_RountingListe(owner)

    def get_Size(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            try:
                return 'DN ' + str(int(round(conn.Radius*609.6)))
            except:
                Height = str(int(round(conn.Height*304.8)))
                Width = str(int(round(conn.Width*304.8)))
                return Width + 'x' + Height
    
    def get_Art(self):
        systyp = self.System
        if systyp:
            if systyp.upper().find('24H') != -1:
                self.art = '24h'
            elif systyp.upper().find('TIERHALTUNG') != -1:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'TZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'TAB' 
            else:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB'
        else:
            self.art = 'XXX'
        try:
            if self.Typen.find('IGF_RLT_Laboranschluss_LAB') != -1:
                self.art = 'LAB'
        except:
            pass

class VSR(object):
    def __init__(self,elementid):
        self.elem_id = DB.ElementId(int(elementid))
        self.elem = doc.GetElement(self.elem_id)
        self.einbauort = ''
        self.checked = False
        self.einbauortnr = ''
        self.size = ''
        self.Auslass = ObservableCollection[Luftauslass]()
        self.liste_Raum = []
        self.Luftmengenmin = 0
        self.Luftmengenmax = 0
        self.Luftmengennacht = 0
        self.Luftmengentiefe = 0
        self.art = ''
        
        self.get_Size()
        self.get_Einbauort()

        try:self.familieundtyp = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
        except:self.familieundtyp = ''
            

    def Luftmengenermitteln(self):
        for auslass in self.Auslass:
            self.Luftmengenmin += auslass.Luftmengenmin
            self.Luftmengenmax += auslass.Luftmengenmax
            self.Luftmengennacht += auslass.Luftmengennacht
            self.Luftmengentiefe += auslass.Luftmengentiefe

    def get_Einbauort(self):
        try:
            self.einbauort = self.elem.Space[PHASE].Id.ToString()
            self.einbauortnr = self.elem.Space[PHASE].Number + ' - ' + self.elem.Space[PHASE].LookupParameter('Name').AsString()
        except:
            if not Muster_Pruefen(self.elem):
                logger.error('VSR: Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elem_id.ToString()))
            return 
    
    def get_Art(self):
        for auslass in self.Auslass:
            if auslass.art:
                self.art = auslass.art

                return

    def get_Size(self):
        try:
            diameter = str(int(list(self.elem.MEPModel.ConnectorManager.Connectors)[0].Radius*609.6))
            return 'DN ' + diameter
        except:
            try:
                conn = list(self.elem.MEPModel.ConnectorManager.Connectors)[0]
                Height = str(int(round(conn.Height*304.8)))
                Width = str(int(round(conn.Width*304.8)))
                return Width + 'x' + Height
            except:
                return ''
       
with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe",cancellable=False, step=10) as pb:
    for n, ductid in enumerate(Ductterminalids):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(Ductterminalids))
        if ductid.ToString() in ELEMID_UEBER:continue
        auslass = Luftauslass(ductid)
        if not auslass.einbauort in DICT_LUFTAUSLASS.keys():
            DICT_LUFTAUSLASS[auslass.einbauort] = ObservableCollection[Luftauslass]()
        DICT_LUFTAUSLASS[auslass.einbauort].Add(auslass)
        if not auslass.vsr:continue
        if not auslass.vsr in DICT_MEP_AUSLASS.keys():
            DICT_MEP_AUSLASS[auslass.vsr] = [auslass.raumnr]
        else:
            if auslass.raumnr not in DICT_MEP_AUSLASS[auslass.vsr]:
                DICT_MEP_AUSLASS[auslass.vsr].append(auslass.raumnr)

        if not auslass.einbauort in DICT_EINBAUTEILE.keys():
            DICT_EINBAUTEILE[auslass.einbauort] = [auslass.vsr]
        else:
            if auslass.vsr not in DICT_EINBAUTEILE[auslass.einbauort]:
                DICT_EINBAUTEILE[auslass.einbauort].append(auslass.vsr)

        if not auslass.vsr in DICT_VSR_AUSLASS.keys():
            DICT_VSR_AUSLASS[auslass.vsr] = ObservableCollection[Luftauslass]()
        DICT_VSR_AUSLASS[auslass.vsr].Add(auslass)

liste_temp = DICT_VSR_AUSLASS.keys()[:]
with forms.ProgressBar(title = "{value}/{max_value} Volumenstromregler",cancellable=False, step=10) as pb:
    for n, vsrid in enumerate(liste_temp):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(liste_temp))
        vsr = VSR(vsrid)
        vsr.Auslass = DICT_VSR_AUSLASS[vsrid]
        vsr.Luftmengenermitteln()
        vsr.liste_Raum = DICT_MEP_AUSLASS[vsrid]
        vsr.get_Art()
        DICT_VSR[vsrid] = vsr
        for auslass in vsr.Auslass:
            auslass.VSR_Class = vsr

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

berechnung_nach_0 = {
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

class MEPGrundInfo(object):
    def __init__(self,name,soll,tooltip):
        self.name = name
        self.soll = soll
        self.ist = ''
        self.tooltip = tooltip

class MEPRaum(object):
    def __init__(self, elem_id,list_auslass,list_vsr,list_ueber):
        self.elem_id = elem_id
        self.elem = doc.GetElement(self.elem_id)
        self.box = self.elem.get_BoundingBox(None)
        self.Raumnr = self.elem.Number + ' - ' + self.elem.LookupParameter('Name').AsString()
        self.list_auslass = list_auslass
        self.list_vsr = list_vsr
        self.list_ueber = list_ueber
        
        # Grundinfo.
        self.bezugsnummer = self.elem.LookupParameter('TGA_RLT_VolumenstromProNummer').AsValueString()
        try:self.bezugsname = berechnung_nach_0[self.bezugsnummer]
        except:self.bezugsname = 'keine'
        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''
        self.flaeche = round(self.get_element('Fläche'),1)
        self.volumen = round(self.get_element('Volumen'),1)
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('TGA_RLT_VolumenstromProFaktor')
        self.einheit = self.get_element('TGA_RLT_VolumenstromProEinheit')
        self.hoehe = int(self.get_element('Lichte Höhe'))
        
        # Übersicht
        self.Uebersicht = ObservableCollection[MEPGrundInfo]()
        self.ab_24h = MEPGrundInfo('24h Abluft',self.get_element('IGF_RLT_AbluftSumme24h'),'IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_24h)
        self.zu_min = MEPGrundInfo('Zuluft min.',self.get_element('IGF_RLT_ZuluftminRaum'),'IGF_RLT_ZuluftminRaum')
        self.Uebersicht.Add(self.zu_min)
        self.zu_max = MEPGrundInfo('Zuluft max.',self.get_element('IGF_RLT_ZuluftmaxRaum'),'IGF_RLT_ZuluftmaxRaum')
        self.Uebersicht.Add(self.zu_max)
        self.ab_min = MEPGrundInfo('Abluft min.',self.get_element('IGF_RLT_AbluftminRaum'),'IGF_RLT_AbluftminRaum')
        self.Uebersicht.Add(self.ab_min)
        self.ab_max = MEPGrundInfo('Abluft max.',self.get_element('IGF_RLT_AbluftmaxRaum'),'IGF_RLT_AbluftmaxRaum')
        self.Uebersicht.Add(self.ab_max)
        self.ab_lab_min = MEPGrundInfo('LAB min.',self.get_element('IGF_RLT_AbluftminSummeLabor'),'IGF_RLT_AbluftminSummeLabor')
        self.Uebersicht.Add(self.ab_lab_min)
        self.ab_lab_max = MEPGrundInfo('LAB max.',self.get_element('IGF_RLT_AbluftmaxSummeLabor'),'IGF_RLT_AbluftmaxSummeLabor')
        self.Uebersicht.Add(self.ab_lab_max)

        self.ueber_sum = MEPGrundInfo('Überstrom',self.get_element('IGF_RLT_ÜberströmungRaum'),'IGF_RLT_ÜberströmungRaum')
        self.Uebersicht.Add(self.ueber_sum)
        self.ueber_in = MEPGrundInfo('Über. Ein.',self.get_element('IGF_RLT_ÜberströmungSummeIn'),'IGF_RLT_ÜberströmungSummeIn')
        self.Uebersicht.Add(self.ueber_in)
        self.ueber_aus = MEPGrundInfo('Über. Aus.',self.get_element('IGF_RLT_ÜberströmungSummeAus'),'IGF_RLT_ÜberströmungSummeAus')
        self.Uebersicht.Add(self.ueber_aus)

        self.nb_von = MEPGrundInfo('NB von',self.get_element('IGF_RLT_NachtbetriebVon'),'IGF_RLT_NachtbetriebVon')
        self.Uebersicht.Add(self.nb_von)
        self.nb_bis = MEPGrundInfo('NB bis',self.get_element('IGF_RLT_NachtbetriebBis'),'IGF_RLT_NachtbetriebBis')
        self.Uebersicht.Add(self.nb_bis)
        self.nb_dauer = MEPGrundInfo('NB Dauer',self.get_element('IGF_RLT_NachtbetriebDauer'),'IGF_RLT_NachtbetriebDauer')
        self.Uebersicht.Add(self.nb_dauer)
        self.nb_zu = MEPGrundInfo('NB Zuluft',self.get_element('IGF_RLT_ZuluftNachtRaum'),'IGF_RLT_ZuluftNachtRaum')
        self.Uebersicht.Add(self.nb_zu)
        self.nb_ab = MEPGrundInfo('NB Abluft',self.get_element('IGF_RLT_AbluftNachtRaum'),'IGF_RLT_AbluftNachtRaum')
        self.Uebersicht.Add(self.nb_ab)

        self.tnb_von = MEPGrundInfo('TNB von',self.get_element('IGF_RLT_TieferNachtbetriebVon'),'IGF_RLT_TieferNachtbetriebVon')
        self.Uebersicht.Add(self.tnb_von)
        self.tnb_bis = MEPGrundInfo('TNB bis',self.get_element('IGF_RLT_TieferNachtbetriebBis'),'IGF_RLT_TieferNachtbetriebBis')
        self.Uebersicht.Add(self.tnb_bis)
        self.tnb_dauer = MEPGrundInfo('TNB Dauer',self.get_element('IGF_RLT_TieferNachtbetriebDauer'),'IGF_RLT_TieferNachtbetriebDauer')
        self.Uebersicht.Add(self.tnb_dauer)
        self.tnb_zu = MEPGrundInfo('TNB Zuluft',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'),'IGF_RLT_ZuluftTieferNachtRaum')
        self.Uebersicht.Add(self.tnb_zu)
        self.tnb_ab = MEPGrundInfo('TNB Abluft',self.get_element('IGF_RLT_AbluftTieferNachtRaum'),'IGF_RLT_AbluftTieferNachtRaum')
        self.Uebersicht.Add(self.tnb_ab)

        # self.anl_mep_rzu = MEPGrundInfo('RLT Anl. RZU',self.get_element(''))
        # self.Uebersicht.Add(self.tnb_von)
        # self.anl_mep_rzu = MEPGrundInfo('L Anl. RZU',self.get_element(''))
        # self.Uebersicht.Add(self.tnb_von)
        # self.anl_mep_rzu = MEPGrundInfo('RZU',self.get_element(''))
        # self.Uebersicht.Add(self.tnb_von)

        # self.anl_mep_rab = MEPGrundInfo('RAB',self.get_element('IGF_RLT_TieferNachtbetriebBis'))
        # self.Uebersicht.Add(self.tnb_bis)
        # self.anl_mep_rab = MEPGrundInfo('RAB',self.get_element('IGF_RLT_TieferNachtbetriebBis'))
        # self.Uebersicht.Add(self.tnb_bis)
        # self.anl_mep_rab = MEPGrundInfo('RAB',self.get_element('IGF_RLT_TieferNachtbetriebBis'))
        # self.Uebersicht.Add(self.tnb_bis)

        # self.anl_mep_tzu = MEPGrundInfo('TZU',self.get_element('IGF_RLT_TieferNachtbetriebDauer'))
        # self.Uebersicht.Add(self.tnb_dauer)
        # self.anl_mep_tzu = MEPGrundInfo('TZU',self.get_element('IGF_RLT_TieferNachtbetriebDauer'))
        # self.Uebersicht.Add(self.tnb_dauer)
        # self.anl_mep_tzu = MEPGrundInfo('TZU',self.get_element('IGF_RLT_TieferNachtbetriebDauer'))
        # self.Uebersicht.Add(self.tnb_dauer)

        # self.anl_mep_tab = MEPGrundInfo('TAB',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_zu)
        # self.anl_mep_tab = MEPGrundInfo('TAB',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_zu)
        # self.anl_mep_tab = MEPGrundInfo('TAB',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_zu)

        # self.anl_mep_24h = MEPGrundInfo('24h',self.get_element('IGF_RLT_AbluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_ab)
        # self.anl_mep_24h = MEPGrundInfo('24h',self.get_element('IGF_RLT_AbluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_ab)
        # self.anl_mep_24h = MEPGrundInfo('24h',self.get_element('IGF_RLT_AbluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_ab)

        # self.anl_mep_tab = MEPGrundInfo('TAB',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_zu)
        # self.anl_mep_tab = MEPGrundInfo('TAB',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_zu)
        # self.anl_mep_24h = MEPGrundInfo('24h',self.get_element('IGF_RLT_AbluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_ab)
        # self.anl_mep_24h = MEPGrundInfo('24h',self.get_element('IGF_RLT_AbluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_ab)
        # self.anl_mep_24h = MEPGrundInfo('24h',self.get_element('IGF_RLT_AbluftTieferNachtRaum'))
        # self.Uebersicht.Add(self.tnb_ab)

        # AnlagenNummer_Sys(L)
        # SchachtNummer(Zu/Ab/24/TZU/TAB)
        # Luftmengen(RLT/L)
        
        
        self.Analyse()
    
    def Analyse(self):
        min_zu = 0
        min_ab = 0
        ab24h = 0
        max_zu = 0
        max_ab = 0
        lab_min = 0
        lab_max = 0
        nb_zu = 0
        nb_ab = 0
        tnb_zu = 0
        tnb_ab = 0
        uber_in = 0
        uber_aus = 0

        for uber in self.list_ueber["Ein"]:
            uber_in += uber.menge
        for uber in self.list_ueber["Aus"]:
            uber_aus += uber.menge


        for auslass in self.list_auslass:
            if auslass.art == '24h':
                ab24h += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentiefe
            elif auslass.art == 'LAB':
                lab_min += auslass.Luftmengenmin
                lab_max += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentiefe
            elif auslass.art == 'RZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentiefe
            elif auslass.art == 'RAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentiefe
            elif auslass.art == 'TZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentiefe
            elif auslass.art == 'TAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentiefe
        
        self.zu_min.ist = min_zu
        self.ab_min.ist = min_ab
        self.ab_24h.ist = ab24h
        self.ab_lab_min.ist = lab_min
        self.ab_lab_max.ist = lab_max
        self.nb_zu.ist = nb_zu
        self.nb_ab.ist = nb_ab
        self.tnb_zu.ist = tnb_zu
        self.tnb_ab.ist = tnb_ab
        self.zu_max.ist = max_zu
        self.ab_max.ist = max_ab
        self.ueber_in.ist = uber_in
        self.ueber_aus.ist = uber_aus
        self.ueber_sum.ist = uber_in-uber_aus

    def get_element(self, param_name):
        param = self.elem.LookupParameter(param_name)
        if not param:
            logger.error("Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return get_value(param)
    

mepraum_liste = {}
with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb:
    for n,space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n+1, len(spaces))
        if space_id.ToString() in DICT_UEBERSTROM.keys():
            list_ueber = DICT_UEBERSTROM[space_id.ToString()]
        else:
            list_ueber = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}

        if space_id.ToString() in DICT_EINBAUTEILE.keys():
            list_vsr = ObservableCollection[VSR]()
            for e in DICT_EINBAUTEILE[space_id.ToString()]:
                try:
                    list_vsr.Add(DICT_VSR[e])
                except Exception as e1:
                    logger.error(e1)
        else:
            list_vsr = ObservableCollection[VSR]()
        
        if space_id.ToString() in DICT_LUFTAUSLASS.keys():
            list_auslass = DICT_LUFTAUSLASS[space_id.ToString()]
        else:
            list_auslass = ObservableCollection[Luftauslass]()

        mepraum = MEPRaum(space_id,list_auslass,list_vsr,list_ueber)
        mepraum_liste[mepraum.Raumnr] = mepraum

auslass_liste = []
einbauteile_dict = {}

class MEPRaum_Uebersicht(forms.WPFWindow):
    def __init__(self, xaml_file_name,mepraum_liste):
        from System.Windows import Visibility
        from rpw import revit,DB
        # self.XYZAnpassen = DB.XYZ(2/304.8,2/304.8,2/304.8)
        self.resizeAnsicht = ResizeAnsicht()
        self.resetAnsicht = ResetAnsicht()
        self.resizeAnsichtEvent = ExternalEvent.Create(self.resizeAnsicht)
        self.resetAnsichtEvent = ExternalEvent.Create(self.resetAnsicht)
        self.uidoc = revit.uidoc
        self.doc = revit.doc
        self.view = self.uidoc.ActiveView
        self.raumanzeigen = False
        self.DB = DB
        self.originalboxmax = DB.XYZ(self.view.GetSectionBox().Max.X,self.view.GetSectionBox().Max.Y,self.view.GetSectionBox().Max.Z)
        self.originalboxmin = DB.XYZ(self.view.GetSectionBox().Min.X,self.view.GetSectionBox().Min.Y,self.view.GetSectionBox().Min.Z)
        self.visible = Visibility.Visible
        self.hidden = Visibility.Hidden
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.mepraum_liste = mepraum_liste
        self.Raumnr.ItemsSource = sorted(self.mepraum_liste.keys())
        self.raumnr = None
        self.mepraum = None
        self.temp_luftauslass = ObservableCollection[Luftauslass]()
        self.warnung.Visibility= self.hidden
        self.auswahl.Visibility= self.hidden
        self.raumwechsel_hinweis.Visibility= self.hidden
        self.Rauminfo.Visibility= self.hidden
        self.rauman.IsEnabled = False
          
    
    def set_up(self):
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
        

        self.lv_vsr.ItemsSource = self.mepraum.list_vsr
        self.lv_auslass.ItemsSource = self.mepraum.list_auslass
        
        self.lv_ueber_aus.ItemsSource = self.mepraum.list_ueber['Aus']
        self.lv_ueber_in.ItemsSource = self.mepraum.list_ueber['Ein']
    
    def zeigraum(self, sender, args):
        self.resizeAnsicht.box = self.mepraum.box
        self.raumanzeigen = True
        self.resizeAnsichtEvent.Raise()
    
    def resetview(self):
        if self.raumanzeigen:
            self.resetAnsicht.boxmax = self.originalboxmax
            self.resetAnsicht.boxmin = self.originalboxmin
            self.raumanzeigen = False
            self.resetAnsichtEvent.Raise()


    def anzeigen_VSR(self, sender, args):
        item = self.lv_vsr.SelectedItem
        if not item:
            return

        try:
            sel = self.uidoc.Selection.GetElementIds()
            sel.Clear()
            sel.Add(item.elem_id)
            self.uidoc.Selection.SetElementIds(sel)
            self.uidoc.ShowElements(self.doc.GetElement(item.elem_id))
        except Exception as e:
            print(e)

    def anzeigen_auslass(self, sender, args):
        from rpw import revit
        uidoc = revit.uidoc
        doc = revit.doc
        item = self.lv_auslass.SelectedItem
        if not item:
            return

        try:
            sel = uidoc.Selection.GetElementIds()
            sel.Clear()
            sel.Add(item.elem_id)
            uidoc.Selection.SetElementIds(sel)
            uidoc.ShowElements(doc.GetElement(item.elem_id))
        except Exception as e:
            print(e)
    
    def anzeigen_ueberaus(self, sender, args):
        from rpw import revit
        uidoc = revit.uidoc
        doc = revit.doc
        item = self.lv_ueber_aus.SelectedItem
        if not item:
            return

        try:
            sel = uidoc.Selection.GetElementIds()
            sel.Clear()
            sel.Add(item.elem_id)
            uidoc.Selection.SetElementIds(sel)
            uidoc.ShowElements(doc.GetElement(item.elem_id))
        except Exception as e:
            print(e)

    def anzeigen_ueberin(self, sender, args):
        from rpw import revit
        uidoc = revit.uidoc
        doc = revit.doc
        item = self.lv_ueber_in.SelectedItem
        if not item:
            return

        try:
            sel = uidoc.Selection.GetElementIds()
            sel.Clear()
            sel.Add(item.elem_id)
            uidoc.Selection.SetElementIds(sel)
            uidoc.ShowElements(doc.GetElement(item.elem_id))
        except Exception as e:
            print(e)
    def nummer_changed(self, sender, args):
        text = str(self.Raumnr.SelectedItem)
        self.resetview()
        self.rauman.IsEnabled = True
     
        self.temp_luftauslass.Clear()
        try:
            self.mepraum = self.mepraum_liste[text]
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
            for el in self.lv_vsr.Items:
                el.checked = False
            self.set_up()
        except Exception as e:
            print(e)
    
    def nummer_changed_vsr(self, sender, args):
        text = str(self.auswahl.SelectedItem)
        text_original = str(self.Raumnr.SelectedItem)
        if text != text_original:
            self.Rauminfo.Visibility = self.visible
            self.Rauminfo.Text = "Die angezeigte Rauminfomationen beziehen sich auf Raum => " + text
        else:
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
                if el.einbauort == self.mepraum.elem_id.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass

        except Exception as e:
            print(e)
    
    def suche_changed(self, sender, args):
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
        except Exception as e:
            logger.error(e)

    def checked_changed(self, sender, args):
        Checked = sender.IsChecked
        item = sender.DataContext
        self.temp_luftauslass.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            item.checked = True

            self.lv_vsr.Items.Refresh()

            for el in item.Auslass:
                if el.einbauort == self.mepraum.elem_id.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass

            if len(item.liste_Raum) > 1:
                self.raumwechsel_hinweis.Visibility= self.visible
                self.warnung.Visibility= self.visible
                self.auswahl.Visibility= self.visible
                self.auswahl.ItemsSource = item.liste_Raum
            else:
                self.raumwechsel_hinweis.Visibility= self.hidden
                self.warnung.Visibility= self.hidden
                self.auswahl.Visibility= self.hidden

        else:
            self.raumwechsel_hinweis.Visibility= self.hidden
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
            self.lv_auslass.ItemsSource = self.mepraum.list_auslass

mepWPF = MEPRaum_Uebersicht("MEP.xaml",mepraum_liste)
try:
    mepWPF.Show()
except Exception as e:
    logger.error(e)
    mepWPF.Close()
    script.exit()