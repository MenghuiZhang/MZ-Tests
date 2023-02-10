# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,UIView,RevitCommandId,UIApplication,PostableCommand 
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility
from rpw import revit,DB,UI
from pyrevit import script,forms
from System.Windows.Input import ModifierKeys,Keyboard,Key
from System.Collections.Generic import List
from System.Windows.Media import Brushes 

def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if not param:return
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

def Muster_Pruefen(el):
    '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
    try:
        bb = el.LookupParameter('Bearbeitungsbereich').AsValueString()
        if bb == 'KG4xx_Musterbereich':return True
        else:return False
    except:return False

logger = script.get_logger()

doc = revit.doc

RED = Brushes.Red
BLACK = Brushes.Black

HIDDEN = Visibility.Hidden
VISIBLE = Visibility.Visible

DICT_MEP_NUMMER_NAME = {} # 100.103 : 100.103 - Büro
DICT_MEP_AUSLASS = {} ## MEPID: OB(Auslässe)
DICT_MEP_VSR = {} ## MEPID: [VSRID]
DICT_MEP_UEBERSTROM = {} ## MEPID: {'Ein': ..., 'Aus':...}

ELEMID_UEBER = [] ## ElememntId Überstrom

DICT_VSR_MEP_NUR_NUMMER = {} ## VSRID: [MEPNr,MEPNr]
DICT_VSR_MEP = {} ## VSRID: [MEPNr-Name,MEPNr-Name]
DICT_VSR_AUSLASS = {} ## VSRID: OB(Auslässe)
DICT_VSR = {} ## VSRID: Class VSR

def get_MEP_NUMMER_NAMR():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        DICT_MEP_NUMMER_NAME[el.Number] = [el.Number + ' - ' + el.LookupParameter('Name').AsString(),el.Id.ToString()]
    
get_MEP_NUMMER_NAMR()


class NoMEPSpace(Exception):
    def __init__(self,elemid,typ,family,art):
        self.elemid = elemid
        self.typ = typ
        self.family = family
        self.art = art
        self.errorinfo = '{}: Einbauort konnte nicht ermittelt werden, FamilieName: {}, TypName: {}, ElementId: {}'.format(self.art,self.family,self.typ,self.elemid.ToString())
    
    def __str__(self):
        return self.errorinfo

class FamilieExemplar(object):
    def __init__(self,elemid,art):
        self.elemid = elemid
        self.art = art
        self.elem = doc.GetElement(self.elemid)
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.familyandtyp = self.familyname + ': ' + self.typname
        self.ismuster =  Muster_Pruefen(self.elem)
        self.phase = doc.GetElement(self.elem.CreatedPhaseId)
        self.raum = ''
        self.raumnummer = ''
        self.raumid = ''
        self.GetRaum()
    
    def GetRaum(self):
        try:
            self.raumid = self.elem.Space[self.phase].Id.ToString()
            self.raumnummer = self.elem.Space[self.phase].Number
            self.raum = self.elem.Space[self.phase].Number + ' - ' + self.elem.Space[self.phase].LookupParameter('Name').AsString()
        except:
            if not self.ismuster:
                try:
                    param_einbauort = get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
                    if param_einbauort not in DICT_MEP_NUMMER_NAME.keys():
                        raise NoMEPSpace(self.elemid,self.typname,self.familyname,self.art)
                    else:
                        self.raum = DICT_MEP_NUMMER_NAME[param_einbauort][0]
                        self.raumid = DICT_MEP_NUMMER_NAME[param_einbauort][1]
                except NoMEPSpace as e:
                    logger.error(str(e))
            return

class Baugruppe:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        if self.Pruefen():
            self.Volumen = get_value(self.elem.LookupParameter('IGF_RLT_Überströmung'))
            self.Eingang = ''
            self.Ausgang = ''
            try:self.Analyse()
            except Exception as e:logger.error(e)

    @property
    def auslaesse(self):
        auslass_liste = []
        for elemid in self.elem.GetMemberIds():
            elem = doc.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008013' and elem.Symbol.FamilyName.upper().find("ÜBER") != -1:
                auslass_liste.append(elem.Id)
        return auslass_liste
    
    def Pruefen(self):
        return len(self.auslaesse) == 2
   
    
    def Analyse(self):
        for elemid in self.auslaesse:
            auslass = UeberStromAuslass(elemid,self.Volumen)
            if auslass.typ == 'Aus':
                self.Ausgang = auslass
            elif auslass.typ == 'Ein':
                self.Eingang = auslass
        self.Ausgang.anderesraum = self.Eingang.raum
        self.Eingang.anderesraum = self.Ausgang.raum

class UeberStromAuslass(FamilieExemplar):
    def __init__(self,elemid,vol):
        FamilieExemplar.__init__(self,elemid,'Überstrom')
        self.menge = vol
        self.anderesraum = ""
        self.typ = self.get_typ()
        
    @property
    def conns(self):
        return list(self.elem.MEPModel.ConnectorManager.Connectors)

    def get_typ(self):
        conn = self.conns[0]
        if conn.Direction.ToString() == 'Out':
            return 'Aus'
        elif conn.Direction.ToString() == 'In':
            return 'Ein'

class Luftauslass(FamilieExemplar):
    def __init__(self,elemid):
        FamilieExemplar.__init__(self,elemid,'Luftauslass')
        self.Luftmengenmin = 0
        self.Luftmengenmax = 0
        self.Luftmengennacht = 0
        self.Luftmengentiefe = 0
        self.Luftmengenermitteln()

        self.vsr = ''
        self.RoutingListe = []
        self.Einbauteile = []
        self.get_RountingListe(self.elem)

        self.size = self.get_Size()

        self.VSR_Class = None

        try:self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:self.System = ''

        try:
            self.AnlNr = doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId()).LookupParameter('IGF_X_AnlagenNr').AsValueString()
        except:self.AnlNr = ''

        try:self.Typen = self.elem.Symbol.LookupParameter('Typenkommentare').AsString()
        except:self.Typen = ''

        self.art = ''

        self.get_Art()

    def Luftmengenermitteln(self):
        try:
            self.Luftmengenmin = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin'))),2)
        except:
            pass
        try:
            self.Luftmengenmax = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax'))),2)
        except:
            pass
        try:
            self.Luftmengennacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht'))),2)
        except:
            pass
        try:
            self.Luftmengentiefe = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht'))),2)
        except:
            pass

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
                                if faminame.upper().find('VOLUMENSTROMREGLER') != -1 or faminame.upper().find('VSR') != -1 or faminame.upper().find('VR') != -1:
                                    self.vsr = owner.Id.ToString()
                                    return
                            self.get_RountingListe(owner)

    def get_Size(self):
        try:
            conns = self.elem.MEPModel.ConnectorManager.Connectors
            for conn in conns:
                try:
                    return 'DN ' + str(int(round(conn.Radius*609.6)))
                except:
                    Height = str(int(round(conn.Height*304.8)))
                    Width = str(int(round(conn.Width*304.8)))
                    return Width + 'x' + Height
        except:return
    
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
            if self.familyandtyp.find('Abzug') != -1 and self.art != '24h':
                self.art = 'LAB'
        except:
            pass

class VSR(FamilieExemplar):
    def __init__(self,elemid):
        FamilieExemplar.__init__(self,elemid,'VSR')
        self.checked = False
        self.size = ''
        self.Auslass = ObservableCollection[Luftauslass]()
        self.liste_Raum = []
        self.liste_Raum_nurNummer = []
        self.Luftmengenmin = 0
        self.Luftmengenmax = 0
        self.Luftmengennacht = 0
        self.Luftmengentiefe = 0
        self.art = ''
        
        self.size = self.get_Size()

    def Luftmengenermitteln(self):
        Luftmengenmin = 0
        Luftmengenmax = 0
        Luftmengennacht = 0
        Luftmengentiefe = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(auslass.Luftmengenmin)
            Luftmengenmax += float(auslass.Luftmengenmax)
            Luftmengennacht += float(auslass.Luftmengennacht)
            Luftmengentiefe += float(auslass.Luftmengentiefe)
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengennacht = int(round(Luftmengennacht))
        self.Luftmengentiefe = int(round(Luftmengentiefe))
    
    def Luftmengenermitteln_new(self):
        Luftmengenmin = 0
        Luftmengenmax = 0
        Luftmengennacht = 0
        Luftmengentiefe = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
            Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
            Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
            Luftmengentiefe += float(str(auslass.Luftmengentiefe).replace(',', '.'))
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengennacht = int(round(Luftmengennacht))
        self.Luftmengentiefe = int(round(Luftmengentiefe))
        

    def get_Art(self):
        for auslass in self.Auslass:
            if auslass.art:
                self.art = auslass.art
                return

    def get_Size(self):
        try:
            conns = self.elem.MEPModel.ConnectorManager.Connectors
            for conn in conns:
                try:
                    return 'DN ' + str(int(round(conn.Radius*609.6)))
                except:
                    Height = str(int(round(conn.Height*304.8)))
                    Width = str(int(round(conn.Width*304.8)))
                    return Width + 'x' + Height
        except:return

def Get_Ueberstrom_Info():
    Baugruppen = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title = "{value}/{max_value} Überströmungsbaugruppen",step=10) as pb:
        for n, BGid in enumerate(Baugruppen):
            pb.update_progress(n + 1, len(Baugruppen))
            baugruppe = Baugruppe(BGid)
            if not baugruppe.Pruefen():
                continue
            if not baugruppe.Eingang.raumid in DICT_MEP_UEBERSTROM.keys():
                DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
            if not baugruppe.Ausgang.raumid in DICT_MEP_UEBERSTROM.keys():
                DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
            DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid][baugruppe.Eingang.typ].Add(baugruppe.Eingang)
            DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid][baugruppe.Ausgang.typ].Add(baugruppe.Ausgang)
            ELEMID_UEBER.append(baugruppe.Eingang.elemid.ToString())
            ELEMID_UEBER.append(baugruppe.Ausgang.elemid.ToString())

Get_Ueberstrom_Info()

def Get_Auslass_Info():
    Ductterminalids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe", step=10) as pb:
        for n, ductid in enumerate(Ductterminalids):
            pb.update_progress(n + 1, len(Ductterminalids))
            if ductid.ToString() in ELEMID_UEBER:continue
            auslass = Luftauslass(ductid)
            if auslass.familyandtyp.upper().find('VORHALTUNG') != -1:
                continue
            

            if not auslass.raumid in DICT_MEP_AUSLASS.keys():
                DICT_MEP_AUSLASS[auslass.raumid] = ObservableCollection[Luftauslass]()
            DICT_MEP_AUSLASS[auslass.raumid].Add(auslass)

            if not auslass.vsr:continue
            
            if not auslass.vsr in DICT_VSR_MEP.keys():
                DICT_VSR_MEP[auslass.vsr] = [auslass.raum]
                DICT_VSR_MEP_NUR_NUMMER[auslass.vsr] = [auslass.raumnummer]
            else:
                if auslass.raum not in DICT_VSR_MEP[auslass.vsr]:
                    DICT_VSR_MEP[auslass.vsr].append(auslass.raum)
                    DICT_VSR_MEP_NUR_NUMMER[auslass.vsr].append(auslass.raumnummer)

            if not auslass.raumid in DICT_MEP_VSR.keys():
                DICT_MEP_VSR[auslass.raumid] = [auslass.vsr]
            else:
                if auslass.vsr not in DICT_MEP_VSR[auslass.raumid]:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.vsr)

            if not auslass.vsr in DICT_VSR_AUSLASS.keys():
                DICT_VSR_AUSLASS[auslass.vsr] = ObservableCollection[Luftauslass]()
            DICT_VSR_AUSLASS[auslass.vsr].Add(auslass)

    liste_temp = DICT_VSR_AUSLASS.keys()[:]

    with forms.ProgressBar(title = "{value}/{max_value} Volumenstromregler", step=10) as pb:
        for n, vsrid in enumerate(liste_temp):
            pb.update_progress(n + 1, len(liste_temp))
            vsr = VSR(DB.ElementId(int(vsrid)))
            vsr.Auslass = DICT_VSR_AUSLASS[vsrid]

            vsr.Luftmengenermitteln()
            vsr.liste_Raum = DICT_VSR_MEP[vsrid]
            vsr.liste_Raum_nurNummer = DICT_VSR_MEP_NUR_NUMMER[vsrid]
            vsr.get_Art()
            DICT_VSR[vsrid] = vsr
            for auslass in vsr.Auslass:
                auslass.VSR_Class = vsr

Get_Auslass_Info()

class MEPGrundInfo(object):
    def __init__(self,name,soll,tooltip):
        self.name = name
        self.soll = soll
        self.ist = ''
        self.tooltip = tooltip

class MEPSchachtInfo(object):
    def __init__(self,name,nr,menge):
        self.name = name
        self.nr = nr
        self.menge = menge

class MEPAnlagenInfo(object):
    def __init__(self,name,mep_nr,mep_mengen,sys_nr,sys_mengen):
        self.name = name
        self.mep_nr = mep_nr
        self.mep_mengen = mep_mengen
        self.sys_nr = sys_nr
        self.sys_mengen = sys_mengen 

spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
Dict_Ueber_Manuell = {}
for ele in spaces_collector:
    raum = get_value(ele.LookupParameter("TGA_RLT_RaumÜberströmungAus"))
    if raum:
        summe2 = get_value(ele.LookupParameter('TGA_RLT_RaumÜberströmungMenge'))
        if raum not in Dict_Ueber_Manuell.keys():
            Dict_Ueber_Manuell[raum] = summe2
        else:Dict_Ueber_Manuell[raum] += summe2

class MEPRaum(object):
    
    def __init__(self, elemid,list_vsr):
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
        self.einheit = {
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
                }
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.box = self.elem.get_BoundingBox(None)
        self.Raumnr = self.elem.Number + ' - ' + self.elem.LookupParameter('Name').AsString()
        self.list_vsr = list_vsr
        self.IsTierRaum = (self.get_element('IGF_RLT_Tierkäfig_raumunabhängig') != 0)
        self.IsSchacht = (self.get_element('TGA_RLT_InstallationsSchacht') != 0)

        try:self.list_auslass = DICT_MEP_AUSLASS[self.elemid.ToString()]
        except:self.list_auslass = ObservableCollection[Luftauslass]()
        
        try:self.list_ueber = DICT_MEP_UEBERSTROM[self.elemid.ToString()]
        except:self.list_ueber = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        
        # Grundinfo.
        self.bezugsnummer = self.elem.LookupParameter('TGA_RLT_VolumenstromProNummer').AsValueString()
        try:self.bezugsname = self.berechnung_nach[self.bezugsnummer]
        except:self.bezugsname = 'keine'
        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''
        self.flaeche = round(self.get_element('Fläche')+0.5, 0)
        self.volumen = round(self.get_element('Volumen'),1)
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('TGA_RLT_VolumenstromProFaktor')
        try:self.einheit = self.einheit[self.bezugsnummer]
        except:self.einheit = ''
        self.hoehe = int(self.get_element('Lichte Höhe'))
        self.nachtbetrieb = self.get_element('IGF_RLT_Nachtbetrieb')
        self.tiefenachtbetrieb = self.get_element('IGF_RLT_TieferNachtbetrieb')
        self.NB_LW = self.get_element('IGF_RLT_NachtbetriebLW')
        self.T_NB_LW = self.get_element('IGF_RLT_TieferNachtbetriebLW')

        
        # Übersicht
        self.Uebersicht = ObservableCollection[MEPGrundInfo]()
        self.ab_24h = MEPGrundInfo('24h Abluft',self.get_element('IGF_RLT_AbluftSumme24h'),'IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_24h)
        self.zu_min = MEPGrundInfo('Zuluft min.',self.get_element('IGF_RLT_ZuluftminRaum'),'IGF_RLT_ZuluftminRaum')
        self.Uebersicht.Add(self.zu_min)
        self.zu_max = MEPGrundInfo('Zuluft max.',self.get_element('IGF_RLT_ZuluftmaxRaum'),'IGF_RLT_ZuluftmaxRaum')
        self.Uebersicht.Add(self.zu_max)
        self.ab_min = MEPGrundInfo('Abluft min.',self.get_element('IGF_RLT_AbluftminRaum'),'IGF_RLT_AbluftminRaum-IGF_RLT_AbluftminSummeLabor-IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_min)
        self.ab_max = MEPGrundInfo('Abluft max.',self.get_element('IGF_RLT_AbluftmaxRaum'),'IGF_RLT_AbluftmaxRaum-IGF_RLT_AbluftmaxSummeLabor-IGF_RLT_AbluftSumme24h')
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
        self.ueber_in_manuell = MEPGrundInfo('Über.Ein.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        self.Uebersicht.Add(self.ueber_in_manuell)
        self.ueber_aus_manuell = MEPGrundInfo('Über.Aus.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        try:self.ueber_aus_manuell.soll = Dict_Ueber_Manuell[self.elem.Number]
        except:self.ueber_aus_manuell.soll = 0
        self.Uebersicht.Add(self.ueber_aus_manuell)

        self.ab_min.soll = self.ab_min.soll - self.ab_24h.soll
        self.ab_max.soll = self.ab_max.soll - self.ab_24h.soll
        self.angezuluft = 0
        self.angeabluft = 0

        # if self.nachtbetrieb != 0:
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

        # if self.tiefenachtbetrieb != 0:
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
        
        if self.IsTierRaum != 0:
            self.tier_zu_min = MEPGrundInfo('TZU min.',self.get_element('IGF_RLT_Luftmenge_min_TZU'),'IGF_RLT_Luftmenge_min_TZU')
            self.Uebersicht.Add(self.tier_zu_min)
            self.tier_zu_max = MEPGrundInfo('TZU max.',self.get_element('IGF_RLT_Luftmenge_max_TZU'),'IGF_RLT_Luftmenge_max_TZU')
            self.Uebersicht.Add(self.tier_zu_max)
            self.tier_ab_min = MEPGrundInfo('TAB min.',self.get_element('IGF_RLT_Luftmenge_min_TAB'),'IGF_RLT_Luftmenge_min_TAB')
            self.Uebersicht.Add(self.tier_ab_min)
            self.tier_ab_max = MEPGrundInfo('TAB max.',self.get_element('IGF_RLT_Luftmenge_max_TAB'),'IGF_RLT_Luftmenge_max_TAB')
            self.Uebersicht.Add(self.tier_ab_max)
            
        self.Druckstufe = MEPGrundInfo('Druckstufe',self.get_element('IGF_RLT_RaumDruckstufeEingabe'),'IGF_RLT_RaumDruckstufeEingabe')
        self.Uebersicht.Add(self.Druckstufe)
        # Anlagen
        self.Anlagen_info = ObservableCollection[MEPAnlagenInfo]()

        # Schacht
        self.Schacht_info = ObservableCollection[MEPSchachtInfo]()

        self.rzu_Schacht = MEPSchachtInfo('ZU',self.get_element('TGA_RLT_SchachtZuluft'),self.get_element('IGF_RLT_AnlagenRaumZuluft'))
        self.Schacht_info.Add(self.rzu_Schacht)
        self.rab_Schacht = MEPSchachtInfo('AB',self.get_element('TGA_RLT_SchachtAbluft'),self.get_element('IGF_RLT_Anlageluft_roundnRaumAbluft'))
        self.Schacht_info.Add(self.rab_Schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('TGA_RLT_Schacht24hAbluft'),self.get_element('IGF_RLT_AnlagenRaum24hAbluft'))
        self.Schacht_info.Add(self._24h_Schacht)
        if self.IsTierRaum:
            self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.get_element('IGF_RLT_Luftmenge_max_TZU'))
            self.Schacht_info.Add(self.tzu_Schacht)
            self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.get_element('IGF_RLT_Luftmenge_max_TAB'))
            self.Schacht_info.Add(self.tab_Schacht)
       
        self.Analyse()
        self.get_Anlagen_info()
    
    def luft_round(self,luft):
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Nachtbetrieb_Berechnen(self):
        if self.nachtbetrieb:
            self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll + 24.00
            self.nb_zu.soll = self.luft_round(self.NB_LW * self.volumen) + self.Druckstufe.soll
            abweichung = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
            if abweichung < 0:
                self.nb_zu.soll -= abweichung
                self.nb_ab.soll = self.ab_lab_min.soll
            else:
                self.nb_ab.soll = self.ab_lab_min.soll + abweichung

        if self.tiefenachtbetrieb:
            self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll + 24.00
            self.nb_dauer.soll -= self.tnb_dauer.soll
            self.tnb_zu.soll = self.luft_round(self.T_NB_LW * self.volumen) + self.Druckstufe.soll
            abweichung = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
            if abweichung < 0:
                self.tnb_zu.soll -= abweichung
                self.tnb_ab.soll = self.ab_lab_min.soll
            else:
                self.tnb_ab.soll = self.ab_lab_min.soll + abweichung

    def Tagesbetrieb_Berechnen(self):
        if self.flaeche == 0:
            return
        if self.bezugsnummer == '5.1':
            logger.info("Berechnung nach Fläche. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.flaeche * self.faktor) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            self.zu_min.soll += abweichung
            self.ab_min.soll = 0
            # self.angezuluft = self.zu_min.soll
            # self.angeabluft = self.angezuluft
        
        elif self.bezugsnummer == '5.2':
            logger.info("Berechnung nach Luftwechsel. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.volumen * self.faktor) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            self.zu_min.soll += abweichung
            # self.angezuluft = self.zu_min.soll
            # self.angeabluft = self.angezuluft
            self.ab_min.soll = 0
        
        elif self.bezugsnummer == '5.3':
            logger.info("Berechnung nach Person. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.personen * self.faktor) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            self.ab_min.soll = 0      
            self.zu_min.soll += abweichung
            # self.angezuluft = self.zu_min.soll
            # self.angeabluft = self.angezuluft
        
        elif self.bezugsnummer == '6.1':
            logger.info("Berechnung nach Fläche. Nur Abluft")
            self.angezuluft = self.luft_round(self.flaeche * self.faktor) + self.Druckstufe.soll
            abweichung = self.angezuluft - self.ueber_in.soll - self.ueber_in_manuell.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft -= abweichung2
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung2
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
        
        elif self.bezugsnummer == '6.2':
            logger.info("Berechnung nach Luftwechsel. Nur Abluft")
            self.angezuluft = self.luft_round(self.volumen * self.faktor) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft -= abweichung2
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung2
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
        
        elif self.bezugsnummer == '6.3':
            logger.info("Berechnung nach Person. Nur Abluft")
            self.angezuluft = self.luft_round(self.personen * self.faktor) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft -= abweichung2
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung2
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
        
        elif self.bezugsnummer == '1':
            logger.info("Berechnung nach Fläche")
            self.zu_min.soll = self.luft_round(self.flaeche * self.faktor) + self.Druckstufe.soll
            self.angezuluft = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                # self.angezuluft -= abweichung
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll

        elif self.bezugsnummer == '2':
            logger.info("Berechnung nach Luftwechsel")
            self.zu_min.soll = self.luft_round(self.volumen * self.faktor) + self.Druckstufe.soll
            self.angezuluft = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                # self.angezuluft -= abweichung
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll

        elif self.bezugsnummer == '3':
            logger.info("Berechnung nach Personen")
            self.zu_min.soll = self.luft_round(self.personen * self.faktor) + self.Druckstufe.soll
            self.angezuluft = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                # self.angezuluft -= abweichung
                self.ab_min.soll = self.ab_lab_min.soll + self.ab_24h.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + self.ab_24h.soll+ abweichung
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll

        elif self.bezugsnummer == '4':
            logger.info("Berechnung nach manuell")
            self.zu_min.soll = self.luft_round(self.faktor) + self.Druckstufe.soll
            self.angezuluft = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                # self.angezuluft -= abweichung
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll

        elif self.bezugsnummer == '5':
            logger.info("Berechnung nach manuell. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.faktor) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {}-{} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + self.ueber_aus_manuell.soll + (self.ab_lab_min.soll + self.ab_24h.soll)
            self.zu_min.soll += abweichung
            # self.angezuluft = self.zu_min.soll
            # self.angeabluft = self.angezuluft

        elif self.bezugsnummer == '6':
            logger.info("Berechnung nach manuell, Nur Abluft")
            self.angezuluft = self.luft_round(self.faktor) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll + self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft -= abweichung2
                self.ab_min.soll = self.ab_lab_min.soll
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
            else:
                self.ab_min.soll = self.ab_lab_min.soll + abweichung2
                # self.angeabluft = self.angezuluft
                # self.abluft_ohne_24h = self.ab_min.soll - self.ab_24h.soll
                # self.abluft_ges = self.ab_min.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll

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
        
        self.zu_min.ist = int(round(min_zu))
        self.ab_min.ist = int(round(min_ab))
        self.ab_24h.ist = int(round(ab24h))
        self.ab_lab_min.ist = int(round(lab_min))
        self.ab_lab_max.ist = int(round(lab_max))
        if self.nachtbetrieb != 0:
            self.nb_zu.ist = int(round(nb_zu))
            self.nb_ab.ist = int(round(nb_ab))
        if self.tiefenachtbetrieb != 0:
            self.tnb_zu.ist = int(round(tnb_zu))
            self.tnb_ab.ist = int(round(tnb_ab))
        self.zu_max.ist = int(round(max_zu))
        self.ab_max.ist = int(round(max_ab))
        self.ueber_in.ist = int(round(uber_in))
        self.ueber_aus.ist = int(round(uber_aus))
        self.ueber_sum.ist = int(round(uber_in-uber_aus))

    def get_Anlagen_info(self):
        self.Anlagen_info.Clear()
        Dict = {}
        for el in self.list_auslass:
            if not el.art in Dict.keys():
                Dict[el.art] = {}
            if not el.AnlNr in Dict[el.art].keys():
                Dict[el.art][el.AnlNr] = float(el.Luftmengenmax)
            else:
                Dict[el.art][el.AnlNr] += float(el.Luftmengenmax)
        
        if 'RZU' in Dict.keys():
            for anl in sorted(Dict['RZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                    self.get_element('IGF_RLT_Luftmenge_RZU'),anl,Dict['RZU'][anl]))
        else:
            if self.get_element('IGF_RLT_AnlagenNr_RZU') or self.get_element('IGF_RLT_Luftmenge_RZU'):
                self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                    self.get_element('IGF_RLT_Luftmenge_RZU'),'',''))
        if 'TZU' in Dict.keys():
            for anl in sorted(Dict['TZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                    self.get_element('IGF_RLT_Luftmenge_TZU'),anl,Dict['TZU'][anl]))
        else:
            if self.get_element('IGF_RLT_AnlagenNr_TZU') or self.get_element('IGF_RLT_Luftmenge_TZU'):
                self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                    self.get_element('IGF_RLT_Luftmenge_TZU'),'',''))
        if 'RAB' in Dict.keys():
            for anl in sorted(Dict['RAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                    self.get_element('IGF_RLT_Luftmenge_RAB'),anl,Dict['RAB'][anl]))
        else:
            if self.get_element('IGF_RLT_AnlagenNr_RAB') or self.get_element('IGF_RLT_Luftmenge_RAB'):
                self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                    self.get_element('IGF_RLT_Luftmenge_RAB'),'',''))
        if 'TAB' in Dict.keys():
            for anl in sorted(Dict['TAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                    self.get_element('IGF_RLT_Luftmenge_TAB'),anl,Dict['TAB'][anl]))
        else:
            if self.get_element('IGF_RLT_AnlagenNr_TAB') or self.get_element('IGF_RLT_Luftmenge_TAB'):
                self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                    self.get_element('IGF_RLT_Luftmenge_TAB'),'',''))
        if '24h' in Dict.keys():
            for anl in sorted(Dict['24h'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                    self.get_element('IGF_RLT_Luftmenge_24h'),anl,Dict['24h'][anl]))
        else:
            if self.get_element('IGF_RLT_AnlagenNr_24h') or self.get_element('IGF_RLT_Luftmenge_24h'):
                self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                    self.get_element('IGF_RLT_Luftmenge_24h'),'',''))
        if 'LAB' in Dict.keys():
            for anl in sorted(Dict['LAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                    self.get_element('IGF_RLT_Luftmenge_LAB'),anl,Dict['LAB'][anl]))  
        else:
            if self.get_element('IGF_RLT_AnlagenNr_LAB') or self.get_element('IGF_RLT_Luftmenge_LAB'):
                self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                    self.get_element('IGF_RLT_Luftmenge_LAB'),'',''))   

        pass
    
    def get_element(self, param_name):
        param = self.elem.LookupParameter(param_name)
        if not param:
            logger.info("Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return get_value(param)

DICT_MEP_ITEMSSOIRCE = {}

def Get_MEP_Info():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()
    
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb:
        for n,space_id in enumerate(spaces):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n+1, len(spaces))
            
            if space_id.ToString() in DICT_MEP_VSR.keys():
                list_vsr = ObservableCollection[VSR]()
                for e in DICT_MEP_VSR[space_id.ToString()]:
                    try:list_vsr.Add(DICT_VSR[e])
                    except:pass
            else:list_vsr = ObservableCollection[VSR]()

            mepraum = MEPRaum(space_id,list_vsr)
            if not mepraum.IsSchacht:DICT_MEP_ITEMSSOIRCE[mepraum.Raumnr] = mepraum

Get_MEP_Info()

class ResizeAnsicht(IExternalEventHandler):
    def __init__(self):
        self.XYZAnpassen = DB.XYZ(1/3.048,1/3.048,1/3.048)
        self.GUI_MEPRaum = None
        
    def Execute(self,uiapp):
        doc = uiapp.ActiveUIDocument.Document
        uidoc = uiapp.ActiveUIDocument
        view = uidoc.ActiveView
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(self.GUI_MEPRaum.mepraum.elemid)
        uidoc.Selection.SetElementIds(sel)
        t = DB.Transaction(doc,'modifiziert Ansichtsgröße')
        self.GUI_MEPRaum.TransMEPRaum.Insert(0,self.GUI_MEPRaum.mepraum)
        self.GUI_MEPRaum.Trans+=1
        t.Start()
        view.IsSectionBoxActive = True
        coll = DB.FilteredElementCollector(doc,view.Id).OfCategory(DB.BuiltInCategory.OST_SectionBox).WhereElementIsNotElementType().ToElements()
        for el in coll:
            if el.LookupParameter('Bearbeitungsbereich').AsValueString().find(view.Name) != -1:box0 = el.get_BoundingBox(view)
        box = view.GetSectionBox()
        Max = self.GUI_MEPRaum.mepraum.box.Max + self.XYZAnpassen + box.Max-box0.Max
        Min = self.GUI_MEPRaum.mepraum.box.Min - self.XYZAnpassen + box.Max-box0.Max
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()
        t.Dispose()
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                return
    def GetName(self):
        return "modifiziert Ansichtsgröße"

class ResetAnsicht(IExternalEventHandler):
    def __init__(self):
        self.boxmax = None
        self.boxmin = None

    def Execute(self,uiapp):
        doc = uiapp.ActiveUIDocument.Document
        uidoc = uiapp.ActiveUIDocument
        view = uidoc.ActiveView
        t = DB.Transaction(doc,'setzt Ansichtsgröße zurück')
        t.Start()
        view.IsSectionBoxActive = True
        box = view.GetSectionBox()
        box.Max = self.boxmax
        box.Min = self.boxmin
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()
        t.Dispose()
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                return
    def GetName(self):
        return "setzt Ansichtsgröße zurück"

class ChangeParameterWert(IExternalEventHandler):
    def __init__(self):
        self.bauteile = None
        self.GUI_MEPRaum = None

    def Execute(self,uiapp):
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        self.GUI_MEPRaum.Trans+=1
        self.GUI_MEPRaum.TransMEPRaum.Insert(0,self.GUI_MEPRaum.mepraum)
        t.Start()
        for bauteil in self.bauteile:
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(bauteil.Luftmengenmin).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(bauteil.Luftmengenmax).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(bauteil.Luftmengennacht).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(bauteil.Luftmengentiefe).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_X_Einbauort').Set(bauteil.raumnummer)
            except:pass
            try:
                numm = ''
                for nummer in sorted(bauteil.liste_Raum_nurNummer):
                    numm += nummer + ', '
                bauteil.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
            except:pass

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "wert schreiben"

class HighLightElement(IExternalEventHandler):
    def __init__(self):
        self.elemId = None        
    def Execute(self,uiapp):
        uiapp.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId]([self.elemId]))
        
    def GetName(self):
        return "Highlight Element"

class ShowElement(IExternalEventHandler):
    def __init__(self):
        self.Raum = ''
        self.Raum_Einbau = ''
        self.elemId = None        
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        if self.Raum != self.Raum_Einbau:
            UI.TaskDialog.Show('Warnung','Der VSR befindet sich in den Raum {}. Bitte diesen Raum zuerst anzeigen!'.format(self.Raum_Einbau))
            return
        else:uidoc.ShowElements(self.elemId)
        
    def GetName(self):
        return "Show Element"

class UNDO(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None
        self.commandid = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)

    def Execute(self,uiapp):
        if self.GUI_MEPRaum.Trans <= 0:return
        
        try:
            uiapp.PostCommand(self.commandid)
            print(self.GUI_MEPRaum.TransMEPRaum[0].Raumnr)
            self.GUI_MEPRaum.mepraum = self.GUI_MEPRaum.TransMEPRaum[0]
            self.GUI_MEPRaum.TransMEPRaum.Remove(self.GUI_MEPRaum.mepraum)
            self.GUI_MEPRaum.Raumnr.Text = self.GUI_MEPRaum.mepraum.Raumnr

            self.GUI_MEPRaum.set_up()
            self.GUI_MEPRaum.Trans -= 1
            if self.GUI_MEPRaum.Trans == 0:
                self.GUI_MEPRaum.rueckspielen.IsEnabled = False            
        except Exception as e:print(e)
        

    def GetName(self):
        return "zurück"

# class UPDATEGUI(IExternalEventHandler):
#     def __init__(self):
#         self.GUI_MEPRaum = None

#     def Execute(self,uiapp):
#         for el in self.GUI_MEPRaum.lv_auslass.Items:
#             try:
#                 el.Luftmengenermitteln()
#             except Exception as e:print(e)
#         for el in self.GUI_MEPRaum.mepraum.list_vsr:
#             try:
#                 el.Luftmengenermitteln_new()
#             except Exception as e:print(e)

#         self.GUI_MEPRaum.lv_auslass.Items.Refresh()
#         self.GUI_MEPRaum.lv_vsr.Items.Refresh()
#         try:
#             self.GUI_MEPRaum.Auswertung_System()
#             self.GUI_MEPRaum.anlageninfo.Items.Refresh()
#             self.GUI_MEPRaum.grundinfo.Items.Refresh()
#             self.schachtinfo.Items.Refresh()
#         except Exception as e:print(e)
#         # TaskDialog.Show('Info','Erledigt!')

#     def GetName(self):
#         return "zurück"


class UPDATEMODEL(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None

    def Execute(self,app):
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        self.GUI_MEPRaum.Trans+=1
        self.GUI_MEPRaum.TransMEPRaum.Insert(0,self.GUI_MEPRaum.mepraum)

        t.Start()
        for bauteil in self.GUI_MEPRaum.mepraum.list_auslass:
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(bauteil.Luftmengenmin).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(bauteil.Luftmengenmax).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(bauteil.Luftmengennacht).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(bauteil.Luftmengentiefe).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_X_Einbauort').Set(bauteil.raumnummer)
            except:pass
            try:
                numm = ''
                for nummer in sorted(bauteil.liste_Raum_nurNummer):
                    numm += nummer + ', '
                bauteil.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
            except:pass
        
        for bauteil in self.GUI_MEPRaum.mepraum.list_vsr:
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(bauteil.Luftmengenmin).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(bauteil.Luftmengenmax).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(bauteil.Luftmengennacht).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(bauteil.Luftmengentiefe).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_X_Einbauort').Set(bauteil.raumnummer)
            except:pass
            try:
                numm = ''
                for nummer in sorted(bauteil.liste_Raum_nurNummer):
                    numm += nummer + ', '
                bauteil.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
            except:pass

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "wert schreiben"