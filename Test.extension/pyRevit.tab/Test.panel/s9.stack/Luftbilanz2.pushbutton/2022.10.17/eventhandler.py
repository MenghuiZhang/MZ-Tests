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
DICT_VSR_VSRLISTE = {} ## VSRID: [VSRID,VSRID]
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
        self.ismuster =  self.Muster_Pruefen()
        self.phase = doc.GetElement(self.elem.CreatedPhaseId)
        self.raum = ''
        self.raumnummer = ''
        self.raumid = ''
        self.GetRaum()
    
    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
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
    def set_up(self):
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.familyandtyp = self.familyname + ': ' + self.typname

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
        self.Luftmengennacht = 0
        self.Luftmengenmax = 0
        self.Luftmengentnacht = 0
        self.System = ''
        self.AnlNr = ''
        self.Typen = ''
        self.vsr = ''
        self.RoutingListe = []
        self.Einbauteile = []
        self.VSR_Liste = []
        self.lufttyp = ''
        self.size = ''
        self.art = ''
        self.iris = ''
        self.VSR_Class = None
        self.Luftmengenermitteln()
        self.size = self.get_Size()
        try:self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:self.System = ''
        try:
            self.AnlNr = doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId()).LookupParameter('IGF_X_AnlagenNr').AsValueString()
        except:self.AnlNr = ''
        try:self.Typen = self.elem.Symbol.LookupParameter('Typenkommentare').AsString()
        except:self.Typen = ''
        if self.familyandtyp.upper().find('VORHALTUNG') == -1 and self.familyandtyp.upper().find('INDUKTION') == -1:
            self.get_RountingListe(self.elem)   
        self.get_Art()
    
    def set_up(self):
        FamilieExemplar.set_up(self)
        self.Luftmengenermitteln()
        self.size = self.get_Size()

    def Luftmengenermitteln(self):
        try:
            self.Luftmengenmin = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin'))),2)
        except:
            pass
        try:
            self.Luftmengennacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht'))),2)
        except:
            pass

        try:
            self.Luftmengenmax = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax'))),2)
        except:
            pass
        try:
            self.Luftmengentnacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht'))),2)
        except:
            pass

    def get_RountingListe(self,element):
        if self.RoutingListe.Count > 500:return
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

                if conns.Size > 2 and self.iris == '':self.iris = -1
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.RoutingListe:
                            if owner.Category.Name == 'Luftkanalzubehör':
                                faminame = owner.Symbol.FamilyName
                                typname = owner.Name
                                if faminame.upper().find('VOLUMENSTROMREGLER') != -1 or faminame.upper().find('VSR') != -1 or faminame.upper().find('VR') != -1:
                                    conns_temp = owner.MEPModel.ConnectorManager.Connectors
                                    for conn_temp in conns_temp:
                                        if self.lufttyp == 'ReturnAir':
                                            if conn_temp.Direction.ToString() == 'In' or conn_temp.Description == 'Haupt':
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                        if self.lufttyp == 'SupplyAir':
                                            if conn_temp.Direction.ToString() == 'Out' or conn_temp.Description == 'Haupt':
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                    if not self.vsr:
                                        self.vsr = owner.Id.ToString()
                                    self.VSR_Liste.append(owner.Id.ToString())
                                    return
                                elif faminame.upper().find('IRIS') != -1 or typname.upper().find('IRIS') != -1:
                                    if not self.iris:self.iris = owner.Id.ToString()
                                    
                            self.get_RountingListe(owner)

    def get_Size(self):
        try:
            conns = self.elem.MEPModel.ConnectorManager.Connectors

            for conn in conns:
                try:
                    self.lufttyp = conn.DuctSystemType.ToString()
                except:pass
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
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB' 
            else:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB'
        else:
            self.art = 'XXX'
        try:
            # if self.familyandtyp.upper().find('ABZUG') != -1 and self.art != '24h':
            #     self.art = 'LAB'
            # if self.familyandtyp.upper().find('VORHALTUNG') != -1 and self.art != '24h':
            #     self.art = 'LAB'
            # if self.familyandtyp.upper().find('LABORGERÄTE') != -1 and self.art != '24h':
            #     self.art = 'LAB'
            # if self.familyandtyp.upper().find('LABOREINRICHTUNG') != -1 and self.art != '24h':
            #     self.art = 'LAB'
            # if self.familyandtyp.find('CAx Abluft Kanalendgitter RE schräg') != -1 and self.art != '24h':
            #     self.art = 'LAB'
            if self.Typen.upper().find('LABORANSCHLUSS') != -1 and self.art != '24h':
                self.art = 'LAB'
            if self.familyandtyp.upper().find('24H') != -1:
                self.art = '24h'
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
        self.Luftmengennacht = 0
        self.Luftmengenmax = 0
        self.Luftmengentnacht = 0
        self.art = ''
        self.size = self.get_Size()
    
    def set_up(self):
        FamilieExemplar.set_up(self)
        self.size = self.get_Size()

    def Luftmengenermitteln(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(auslass.Luftmengenmin)
            Luftmengennacht += float(auslass.Luftmengennacht)
            Luftmengenmax += float(auslass.Luftmengenmax)
            Luftmengentnacht += float(auslass.Luftmengentnacht)
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht))
    
    def Luftmengenermitteln_new(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
            Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
            Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
            Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def Luftmengenermitteln_um(self):
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0

        for auslass in self.Auslass:
            if auslass.art == 'UMZU':
                Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
                Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
                Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
                Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def get_Art(self):
        art_liste = []
        
        for auslass in self.Auslass:
            if auslass.art:
                if auslass.art not in art_liste:
                    art_liste.append(auslass.art)
        
        if art_liste.Count == 0:return
        if art_liste.Count == 1:
            self.art = art_liste[0]
        elif art_liste.Count == 2:
            if 'RZU' in art_liste and 'RAB' in art_liste:
                self.art = 'RUM'
                for auslass in self.Auslass:
                    if auslass.art == 'RZU':
                        auslass.art = 'UMZU'
                    elif auslass.art == 'RAB':
                        auslass.art = 'UMAB'
            elif 'RAB' in art_liste and 'LAB' in art_liste:
                self.art = 'RAB'
            else:
                print(self.elemid.ToString(),art_liste)
        elif art_liste.Count == 3 and '24h' in art_liste and 'LAB' in art_liste and 'RAB' in art_liste:self.art = 'RAB'
        else:print(self.elemid.ToString(),art_liste)


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
    with forms.ProgressBar(title = "{value}/{max_value} Überströmungsbaugruppen",cancellable=True,step=10) as pb:
        for n, BGid in enumerate(Baugruppen):
            if pb.cancelled:
                return
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
    with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe", cancellable=True,step=10) as pb:
        for n, ductid in enumerate(Ductterminalids):
            if pb.cancelled:
                return
            pb.update_progress(n + 1, len(Ductterminalids))
            if ductid.ToString() in ELEMID_UEBER:continue
            auslass = Luftauslass(ductid)
            
            if not auslass.raumid in DICT_MEP_AUSLASS.keys():
                 DICT_MEP_AUSLASS[auslass.raumid] = {}
            if not auslass.art in DICT_MEP_AUSLASS[auslass.raumid].keys():
                DICT_MEP_AUSLASS[auslass.raumid][auslass.art] = {}
            if not auslass.familyandtyp in DICT_MEP_AUSLASS[auslass.raumid][auslass.art].keys():
                DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp] = []
            DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp].append(auslass)

            if not auslass.vsr:continue
            if auslass.familyandtyp.upper().find('VORHALTUNG') != -1:
                continue
            if auslass.familyandtyp.upper().find('INDUKTION') != -1:
                continue
            
            if not auslass.vsr in DICT_VSR_MEP.keys():
                DICT_VSR_MEP[auslass.vsr] = [auslass.raum]
                DICT_VSR_MEP_NUR_NUMMER[auslass.vsr] = [auslass.raumnummer]
            else:
                if auslass.raum not in DICT_VSR_MEP[auslass.vsr]:
                    DICT_VSR_MEP[auslass.vsr].append(auslass.raum)
                    DICT_VSR_MEP_NUR_NUMMER[auslass.vsr].append(auslass.raumnummer)
            if auslass.iris not in [-1,'']:
                if not auslass.iris in DICT_VSR_MEP.keys():
                    DICT_VSR_MEP[auslass.iris] = [auslass.raum]
                    DICT_VSR_MEP_NUR_NUMMER[auslass.iris] = [auslass.raumnummer]
                else:
                    if auslass.raum not in DICT_VSR_MEP[auslass.iris]:
                        DICT_VSR_MEP[auslass.iris].append(auslass.raum)
                        DICT_VSR_MEP_NUR_NUMMER[auslass.iris].append(auslass.raumnummer)
            
            if not auslass.vsr in DICT_VSR_VSRLISTE.keys():
                DICT_VSR_VSRLISTE[auslass.vsr] = auslass.VSR_Liste
            else:
                DICT_VSR_VSRLISTE[auslass.vsr].extend(auslass.VSR_Liste)

            if not auslass.raumid in DICT_MEP_VSR.keys():
                DICT_MEP_VSR[auslass.raumid] = []
                DICT_MEP_VSR[auslass.raumid].append(auslass.vsr) 
                if auslass.iris not in [-1,'']:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.iris)
            else:
                if auslass.vsr not in DICT_MEP_VSR[auslass.raumid]:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.vsr)
                if auslass.iris not in [-1,'']:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.iris)

            if not auslass.vsr in DICT_VSR_AUSLASS.keys():
                DICT_VSR_AUSLASS[auslass.vsr] = {}
            if not auslass.art in DICT_VSR_AUSLASS[auslass.vsr].keys():
                DICT_VSR_AUSLASS[auslass.vsr][auslass.art] = {}
            if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.vsr][auslass.art].keys():
                DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp] = []
            DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp].append(auslass)
            if auslass.iris not in [-1,'']:
                if not auslass.iris in DICT_VSR_AUSLASS.keys():
                    DICT_VSR_AUSLASS[auslass.iris] = {}
                if not auslass.art in DICT_VSR_AUSLASS[auslass.iris].keys():
                    DICT_VSR_AUSLASS[auslass.iris][auslass.art] = {}
                if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.iris][auslass.art].keys():
                    DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp] = []
                DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp].append(auslass)  


    liste_temp = DICT_VSR_AUSLASS.keys()[:]

    with forms.ProgressBar(title = "{value}/{max_value} Volumenstromregler", cancellable=True, step=10) as pb:
        for n, vsrid in enumerate(liste_temp):
            if pb.cancelled:
                return
            pb.update_progress(n + 1, len(liste_temp))
            if vsrid in DICT_VSR_VSRLISTE.keys():
                vsrliste = set(DICT_VSR_VSRLISTE[vsrid])
                vsrliste = list(vsrliste)
                for vsrid_neu in vsrliste:
                    if vsrid != vsrid_neu:
                        if vsrid_neu not in DICT_VSR_AUSLASS.keys():
                            logger.error(vsrid + ',' + vsrid_neu)
                            continue
                        for key0 in DICT_VSR_AUSLASS[vsrid_neu].keys():
                            for key1 in DICT_VSR_AUSLASS[vsrid_neu][key0].keys():
                                for auslass in DICT_VSR_AUSLASS[vsrid_neu][key0][key1]:
                                    if key0 not in DICT_VSR_AUSLASS[vsrid].keys():
                                        DICT_VSR_AUSLASS[vsrid][key0] = {}
                                    if key1 not in DICT_VSR_AUSLASS[vsrid][key0].keys():
                                        DICT_VSR_AUSLASS[vsrid][key0][key1] = []
                                    if auslass not in DICT_VSR_AUSLASS[vsrid][key0][key1]:
                                        DICT_VSR_AUSLASS[vsrid][key0][key1].append(auslass)

            vsr = VSR(DB.ElementId(int(vsrid)))

            vsr.Auslass = ObservableCollection[Luftauslass]()
            for art in sorted(DICT_VSR_AUSLASS[vsrid].keys()):
                for fam in sorted(DICT_VSR_AUSLASS[vsrid][art].keys()):
                    for terminal in DICT_VSR_AUSLASS[vsrid][art][fam]:
                        vsr.Auslass.Add(terminal)
            vsr.get_Art()
            vsr.Luftmengenermitteln()
            vsr.liste_Raum = DICT_VSR_MEP[vsrid]
            vsr.liste_Raum_nurNummer = DICT_VSR_MEP_NUR_NUMMER[vsrid]
            
            DICT_VSR[vsrid] = vsr
            # for auslass in vsr.Auslass:
            #     auslass.VSR_Class = vsr

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
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.box = self.elem.get_BoundingBox(None)
        self.Raumnr = self.elem.Number + ' - ' + self.elem.LookupParameter('Name').AsString()
        self.list_vsr = list_vsr
        # self.angezuluft = 0
        # self.angeabluft = 0

        self.IsTierRaum = (self.get_element('IGF_RLT_Tierkäfig_raumunabhängig') != 0)
        self.IsSchacht = (self.get_element('TGA_RLT_InstallationsSchacht') != 0)
        self.list_auslass = ObservableCollection[Luftauslass]()
        if self.elemid.ToString() in DICT_MEP_AUSLASS.keys():
             for art in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()].keys()):
                for fam in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()][art].keys()):
                    for terminal in DICT_MEP_AUSLASS[self.elemid.ToString()][art][fam]:
                        self.list_auslass.Add(terminal)

        # try:self.list_auslass = DICT_MEP_AUSLASS[self.elemid.ToString()]
        # except:self.list_auslass = ObservableCollection[Luftauslass]()
        
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
        try:self.einheit = self.einheit_liste[self.bezugsnummer]
        except:self.einheit = ''
        self.hoehe = int(self.get_element('Lichte Höhe'))
        self.nachtbetrieb = self.get_element('IGF_RLT_Nachtbetrieb')
        self.tiefenachtbetrieb = self.get_element('IGF_RLT_TieferNachtbetrieb')
        self.NB_LW = self.get_element('IGF_RLT_NachtbetriebLW')
        self.T_NB_LW = self.get_element('IGF_RLT_TieferNachtbetriebLW')

        
        # Übersicht
        self.Uebersicht = ObservableCollection[MEPGrundInfo]()
        self.angezuluft = MEPGrundInfo('Ange.Zuluft',self.get_element('Angegebener Zuluftstrom'),'Angegebener Zuluftstrom')
        self.Uebersicht.Add(self.angezuluft)
        self.angeabluft = MEPGrundInfo('Ange.Abluft',self.get_element('Angegebener Abluftluftstrom'),'Angegebener Abluftluftstrom')
        self.Uebersicht.Add(self.angeabluft)
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
        self.ab_24h = MEPGrundInfo('24h Abluft',self.get_element('IGF_RLT_AbluftSumme24h'),'IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_24h)

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
        self.ueber_aus_manuell.ist =self.ueber_aus_manuell.soll
        self.ueber_in_manuell.ist = self.ueber_in_manuell.soll
        # print(self.Raumnr)
        # print(self.ab_max.soll,self.ab_24h.soll,self.ab_lab_max.soll)
        self.ab_min.soll = self.ab_min.soll - self.ab_24h.soll - self.ab_lab_min.soll
        self.ab_max.soll = self.ab_max.soll - self.ab_24h.soll - self.ab_lab_max.soll
        # self.angezuluft = 0
        # self.angeabluft = 0
        self.IGF_Legende = ''

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
        else:
            self.tier_zu_min = MEPGrundInfo('TZU min.',0,'IGF_RLT_Luftmenge_min_TZU')
            self.Uebersicht.Add(self.tier_zu_min)
            self.tier_zu_max = MEPGrundInfo('TZU max.',0,'IGF_RLT_Luftmenge_max_TZU')
            self.Uebersicht.Add(self.tier_zu_max)
            self.tier_ab_min = MEPGrundInfo('TAB min.',0,'IGF_RLT_Luftmenge_min_TAB')
            self.Uebersicht.Add(self.tier_ab_min)
            self.tier_ab_max = MEPGrundInfo('TAB max.',0,'IGF_RLT_Luftmenge_max_TAB')
            self.Uebersicht.Add(self.tier_ab_max)
            
        self.Druckstufe = MEPGrundInfo('Druckstufe',self.get_element('IGF_RLT_RaumDruckstufeEingabe'),'IGF_RLT_RaumDruckstufeEingabe')
        self.Uebersicht.Add(self.Druckstufe)
        # Anlagen
        self.Anlagen_info = ObservableCollection[MEPAnlagenInfo]()

        # Schacht
        self.Schacht_info = ObservableCollection[MEPSchachtInfo]()

        self.rzu_Schacht = MEPSchachtInfo('ZU',self.get_element('TGA_RLT_SchachtZuluft'),self.get_element('IGF_RLT_AnlagenRaumZuluft'))
        self.Schacht_info.Add(self.rzu_Schacht)
        self.rab_Schacht = MEPSchachtInfo('AB',self.get_element('TGA_RLT_SchachtAbluft'),self.get_element('IGF_RLT_AnlagenRaumAbluft'))
        self.Schacht_info.Add(self.rab_Schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('TGA_RLT_Schacht24hAbluft'),self.get_element('IGF_RLT_AnlagenRaum24hAbluft'))
        self.Schacht_info.Add(self._24h_Schacht)
       
        self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.get_element('IGF_RLT_Luftmenge_max_TZU'))
        self.Schacht_info.Add(self.tzu_Schacht)
        self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.get_element('IGF_RLT_Luftmenge_max_TAB'))
        self.Schacht_info.Add(self.tab_Schacht)
    
        self.Analyse()
        self.get_Anlagen_info()
    
    def update(self):        
        self.ab_24h.soll = self.get_element('IGF_RLT_AbluftSumme24h')
        self.ab_lab_min.soll = self.get_element('IGF_RLT_AbluftminSummeLabor') 
        self.Druckstufe.soll = self.get_element('IGF_RLT_RaumDruckstufeEingabe')
    
    def luft_round(self,luft):
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Druckstufe_Berechnen(self):
        n = abs(int(self.Druckstufe.soll/10)) if abs(int(self.Druckstufe.soll/10)) < 6 else 5
        if self.Druckstufe.soll > 0:self.IGF_Legende = n*'+'
        else:self.IGF_Legende = n*'-'
    
    def Nachtbetrieb_Berechnen(self):
        if self.nachtbetrieb:
            if self.tiefenachtbetrieb == 0:
                if self.nb_bis.soll - self.nb_von.soll > 0:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll
                else:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll + 24.00
            else:
                if self.nb_bis.soll - self.nb_von.soll > 0:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer
                else:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll + 24.00 - self.tnb_dauer
            if self.bezugsname in ["Fläche","Luftwechsel","Person","manuell"]:
                self.nb_zu.soll = self.luft_round(float(self.NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.nb_zu.soll -= abweichung
                    self.nb_ab.soll = 0
                else:
                    self.nb_ab.soll = abweichung
            elif self.bezugsname in ["nurZU_Fläche","nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
                self.nb_zu.soll = self.luft_round(float(self.NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.nb_ab.soll = 0
                else:
                    self.nb_ab.soll = 0
                self.nb_zu.soll -= abweichung
            
            elif self.bezugsname in ["nurAB_Fläche","nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
                self.nb_zu.soll = 0
                abweichung = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll + self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.nb_zu.soll -= abweichung
                    self.nb_ab.soll = 0
                else:
                    self.nb_ab.soll = abweichung 
    
    def TiefeNachtbetrieb_Berechnen(self):
        if self.tiefenachtbetrieb:
            if self.tnb_bis.soll - self.tnb_von.soll > 0:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
            else:
                self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll + 24.00
            if self.bezugsname in ["Fläche","Luftwechsel","Person","manuell"]:
                self.tnb_zu.soll = self.luft_round(float(self.T_NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.tnb_zu.soll -= abweichung
                    self.tnb_ab.soll = 0
                else:
                    self.tnb_ab.soll = abweichung
            elif self.bezugsname in ["nurZU_Fläche","nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
                self.tnb_zu.soll = self.luft_round(float(self.T_NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.tnb_ab.soll = 0
                else:
                    self.tnb_ab.soll = 0
                self.tnb_zu.soll -= abweichung
            
            elif self.bezugsname in ["nurAB_Fläche","nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
                self.tnb_zu.soll = 0
                abweichung = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll + self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.tnb_zu.soll -= abweichung
                    self.tnb_ab.soll = 0
                else:
                    self.tnb_ab.soll = abweichung 

    def Tagesbetrieb_Berechnen(self):
        if self.flaeche == 0:
            return
        if self.bezugsname == 'nurZU_Fläche':
            logger.info("Berechnung nach Fläche. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            self.zu_min.soll += abweichung
            self.ab_min.soll = 0
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            self.zu_max.soll += abweichung
            self.ab_max.soll = 0

        elif self.bezugsname == 'nurZU_Luftwechsel':
            logger.info("Berechnung nach Luftwechsel. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            self.zu_min.soll += abweichung
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll
            self.ab_min.soll = 0

            self.zu_max.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            self.zu_max.soll += abweichung
        
        elif self.bezugsname == 'nurZU_Person':
            logger.info("Berechnung nach Person. Nur Zuluft")
            self.zu_min.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            self.ab_min.soll = 0      
            self.zu_min.soll += abweichung
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            self.ab_max.soll = 0      
            self.zu_max.soll += abweichung
        
        elif self.bezugsname == 'nurAB_Fläche':
            logger.info("Berechnung nach Fläche. Nur Abluft")
            self.angezuluft.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft.soll -= abweichung2
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
            else:
                self.ab_min.soll = abweichung2
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung2
                
        elif self.bezugsname == 'nurAB_Luftwechsel':
            logger.info("Berechnung nach Luftwechsel. Nur Abluft")
            self.angezuluft.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft.soll -= abweichung2
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
            else:
                self.ab_min.soll = abweichung2
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung2
       
        elif self.bezugsname == 'nurAB_Person':
            logger.info("Berechnung nach Person. Nur Abluft")
            self.angezuluft.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.angezuluft.soll -= abweichung2
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung2
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung2
                
        elif self.bezugsname == 'Fläche':
            logger.info("Berechnung nach Fläche")
            self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung
           
        elif self.bezugsname == 'Luftwechsel':
            logger.info("Berechnung nach Luftwechsel")
            self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung

        elif self.bezugsname == 'Person':
            logger.info("Berechnung nach Personen")
            self.zu_min.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung
  
        elif self.bezugsname == 'manuell':
            logger.info("Berechnung nach manuell")
            self.zu_min.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
                
            else:
                self.ab_max.soll = abweichung
        
        elif self.bezugsname == 'nurZUMa':
            logger.info("Berechnung nach manuell, nur Zuluft")
            self.zu_min.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
                
            else:
                self.ab_max.soll = abweichung
        elif self.bezugsname == 'nurABMa':
            logger.info("Berechnung nach manuell, nur Abluft")
            self.ab_min.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.ab_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
                
            else:
                self.ab_max.soll = abweichung

    def Analyse(self):
        min_zu = 0
        min_ab = 0
        max_zu = 0
        max_ab = 0
        ab24h = 0
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
                ab24h += auslass.Luftmengenmin
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'LAB':
                lab_min += auslass.Luftmengenmin
                lab_max += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'RZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentnacht
            elif auslass.art == 'RAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'TZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentnacht
            elif auslass.art == 'TAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
        
        self.zu_min.ist = int(round(min_zu))
        self.ab_min.ist = int(round(min_ab))
        self.zu_max.ist = int(round(max_zu))
        self.ab_max.ist = int(round(max_ab))
        self.ab_24h.ist = int(round(ab24h))
        self.ab_lab_min.ist = int(round(lab_min))
        self.ab_lab_max.ist = int(round(lab_max))
        self.nb_zu.ist = int(round(nb_zu))
        self.nb_ab.ist = int(round(nb_ab))
        self.tnb_zu.ist = int(round(tnb_zu))
        self.tnb_ab.ist = int(round(tnb_ab))
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
                Dict[el.art][el.AnlNr] = float(el.Luftmengenmin)
            else:
                Dict[el.art][el.AnlNr] += float(el.Luftmengenmin)
        
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
    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
              #  logger.info(
              #      "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                if self.elem.LookupParameter(param_name):
                    if self.elem.LookupParameter(param_name).IsReadOnly is True:
                        logger.error(self.elemid)
                        logger.error(param_name)
                        return
                    self.elem.LookupParameter(param_name).SetValueString(str(wert))
        def wert_schreiben2(param_name, wert):
            if self.elem.LookupParameter(param_name):
                if self.elem.LookupParameter(param_name).IsReadOnly is True:
                    logger.error(self.elemid)
                    logger.error(param_name)
                    return
                self.elem.LookupParameter(param_name).Set(wert)


        wert_schreiben("IGF_RLT_AbluftminRaum", self.ab_min.soll+self.ab_lab_min.soll+self.ab_24h.soll+self.tier_ab_min.soll)

        wert_schreiben2("TGA_RLT_VolumenstromProNameRevit", self.bezugsname)
        wert_schreiben("TGA_RLT_VolumenstromProEinheitRevit", self.einheit)
        wert_schreiben("TGA_RLT_VolumenstromProNummer", self.bezugsnummer)
        wert_schreiben("TGA_RLT_VolumenstromProFaktor", float(self.faktor))

        wert_schreiben("IGF_RLT_Nachtbetrieb", self.nachtbetrieb)
        wert_schreiben("IGF_RLT_NachtbetriebLW", self.NB_LW)

        wert_schreiben("IGF_RLT_AbluftSumme24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_AbluftminRaumL24h", self.ab_24h.soll)    
        wert_schreiben("IGF_RLT_AbluftminSummeLabor", self.ab_lab_min.soll)

        wert_schreiben("IGF_RLT_RaumDruckstufeEingabe", self.Druckstufe.soll)
        wert_schreiben2("IGF_RLT_RaumDruckstufeLegende", self.IGF_Legende)
        

        # wert_schreiben("IGF_RLT_AnlagenRaumAbluft", self.ab_min.soll+self.ab_lab_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaumZuluft", self.zu_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaum24hAbluft", self.ab_24h.soll)
        
        wert_schreiben("IGF_RLT_Luftmenge_RAB", self.ab_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_RZU", self.zu_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_Luftmenge_LAB", self.ab_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TAB", self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TZU", self.tier_zu_min.soll)

        wert_schreiben("Angegebener Zuluftstrom", self.angezuluft.soll)
        wert_schreiben("Angegebener Abluftluftstrom", self.angeabluft.soll)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.nb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.nb_ab.soll)

        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer", self.tnb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tnb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tnb_ab.soll)

        # Neue Parameter
        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zu_min.soll)
        wert_schreiben("IGF_RLT_AbluftminSummeLabor24h", self.ab_24h.soll + self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftminRaum24h", self.ab_24h.soll)


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
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(bauteil.Luftmengennacht).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(bauteil.Luftmengenmax).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(bauteil.Luftmengentnacht).replace(',','.'))
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
        pass
        # if self.GUI_MEPRaum.Trans <= 0:return
        
        # try:
        #     uiapp.PostCommand(self.commandid)
        #     print(self.GUI_MEPRaum.TransMEPRaum[0].Raumnr)
        #     self.GUI_MEPRaum.mepraum = self.GUI_MEPRaum.TransMEPRaum[0]
        #     self.GUI_MEPRaum.TransMEPRaum.Remove(self.GUI_MEPRaum.mepraum)
        #     self.GUI_MEPRaum.Raumnr.Text = self.GUI_MEPRaum.mepraum.Raumnr

        #     self.GUI_MEPRaum.set_up()
        #     self.GUI_MEPRaum.Trans -= 1
        #     if self.GUI_MEPRaum.Trans == 0:
        #         self.GUI_MEPRaum.rueckspielen.IsEnabled = False            
        # except Exception as e:print(e)
        

    def GetName(self):
        return "zurück"

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
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(bauteil.Luftmengennacht).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(bauteil.Luftmengenmax).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(bauteil.Luftmengentnacht).replace(',','.'))
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
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(bauteil.Luftmengennacht).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(bauteil.Luftmengenmax).replace(',','.'))
            except:pass
            try:bauteil.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(bauteil.Luftmengentnacht).replace(',','.'))
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

class BERECHNEN(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None

    def Execute(self,app):
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        self.GUI_MEPRaum.Trans+=1
        self.GUI_MEPRaum.TransMEPRaum.Insert(0,self.GUI_MEPRaum.mepraum)

        t.Start()
        # self.GUI_MEPRaum.mepraum.update()
        try:
            self.GUI_MEPRaum.mepraum.TiefeNachtbetrieb_Berechnen()
   
        except Exception as e:pass#print(e,-1)
        
        try:
            self.GUI_MEPRaum.mepraum.Nachtbetrieb_Berechnen()
   
        except Exception as e:pass#print(e,0)
        
        try:
            self.GUI_MEPRaum.mepraum.Tagesbetrieb_Berechnen()
            
        except Exception as e:pass#print(e,1)

        try:
            self.GUI_MEPRaum.mepraum.Druckstufe_Berechnen()
            
        except Exception as e:pass#print(e,2)


        try:
           
            self.GUI_MEPRaum.mepraum.werte_schreiben()
        except Exception as e:pass#print(e,3)


        t.Commit()
        t.Dispose()
        self.GUI_MEPRaum.lv_auslass.Items.Refresh()
        self.GUI_MEPRaum.lv_vsr.Items.Refresh()
        self.GUI_MEPRaum.Auswertung_System()
        self.GUI_MEPRaum.Auswertung_MEP()
        self.GUI_MEPRaum.anlageninfo.Items.Refresh()
        self.GUI_MEPRaum.grundinfo.Items.Refresh()
        self.GUI_MEPRaum.schachtinfo.Items.Refresh()

    def GetName(self):
        return "Luftmenge berechnen"
    
class VERTEILEN(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None

    def Execute(self,app):
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        self.GUI_MEPRaum.Trans+=1
        self.GUI_MEPRaum.TransMEPRaum.Insert(0,self.GUI_MEPRaum.mepraum)

        t.Start()
        for mep in self.GUI_MEPRaum.mepraum_liste.values():
            
            zu = {}
            ab = {}
            h24 = {}
            lab = {}
            for auslass in mep.list_auslass:
                if auslass.art == 'RZU':
                    if auslass.familyandtyp not in zu.keys():
                        zu[auslass.familyandtyp] = [auslass]
                    else:
                        zu[auslass.familyandtyp].append(auslass)
                if auslass.art == 'RAB':
                    if auslass.familyandtyp not in ab.keys():
                        ab[auslass.familyandtyp] = [auslass]
                    else:
                        ab[auslass.familyandtyp].append(auslass)
                if auslass.art == '24h':
                    if auslass.familyandtyp not in h24.keys():
                        h24[auslass.familyandtyp] = [auslass]
                    else:
                        h24[auslass.familyandtyp].append(auslass)
                if auslass.art == 'LAB':
                    if auslass.familyandtyp not in lab.keys():
                        lab[auslass.familyandtyp] = [auslass]
                    else:
                        lab[auslass.familyandtyp].append(auslass)
            
            if int(mep.ab_24h.soll) != int(mep.ab_24h.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist):
                print(30*'-')
                print('24h-Abluft oder Laborabluft in MEP-Raum {} stimmt nicht übereinandern.'.format(mep.Raumnr))
                print('24h-Abluft-soll: {} m³/h, 24h-Abluft-ist: {} m³/h'.format(mep.ab_24h.soll,mep.ab_24h.ist))
                print('Laborabluftmin-soll: {} m³/h, Laborabluftmin-ist: {} m³/h'.format(mep.ab_lab_min.soll,mep.ab_lab_min.ist))
                print('Laborabluftmax-soll: {} m³/h, Laborabluftmax-ist: {} m³/h'.format(mep.ab_lab_max.soll,mep.ab_lab_max.ist))
                print('Bitte manuell anpassen. ')

                continue
            if (int(mep.zu_min.soll) > 0 or int(mep.zu_max.soll) > 0) and len(zu.keys()) == 0:
                print(30*'-')
                print('Es fehlt Zuluftauslass in MEP-Raum {}.'.format(mep.Raumnr))
                print('Zuluft min : {} m³/h, Zuluft max: {} m³/h'.format(mep.zu_min.soll,mep.zu_max.soll))
                print('Bitte manuell anpassen. ')
                continue
            if (int(mep.ab_min.soll) > 0 or int(mep.ab_max.soll) > 0) and len(ab.keys()) == 0:
                print(30*'-')
                print('Es fehlt Abluftauslass in MEP-Raum {}.'.format(mep.Raumnr))
                print('Abluft min : {} m³/h, Abluft max: {} m³/h'.format(mep.ab_min.soll,mep.ab_max.soll))
                print('Bitte manuell anpassen. ')
                continue

            if len(lab.keys()) > 0:
                for key in lab.keys():
                    for auslass in lab[key]:
                        auslass.Luftmengennacht = auslass.Luftmengenmin
                        auslass.Luftmengentnacht = auslass.Luftmengenmin
            if len(h24.keys()) > 0:
                for key in h24.keys():
                    for auslass in h24[key]:
                        auslass.Luftmengennacht = auslass.Luftmengenmin
                        auslass.Luftmengenmax = auslass.Luftmengenmin
                        auslass.Luftmengentnacht = auslass.Luftmengenmin
                        
            if len(zu.keys()) == 1:
                for key in zu.keys():
                    for auslass in zu[key]:
                        auslass.Luftmengenmin = mep.zu_min.soll *1.0 / len(zu[key])
                        auslass.Luftmengennacht = mep.nb_zu.soll *1.0 / len(zu[key])
                        auslass.Luftmengenmax = mep.zu_max.soll *1.0 / len(zu[key])
                        auslass.Luftmengentnacht = mep.tnb_zu.soll *1.0 / len(zu[key])
            else:
                sum_luft = 0
                for key in zu.keys():
                    for auslass in zu[key]:
                        sum_luft += auslass.Luftmengenmin if auslass.Luftmengenmin > 0 else 0.01
                for key in zu.keys():
                    for auslass in zu[key]:
                        auslass.Luftmengenmin = mep.zu_min.soll *1.0 * auslass.Luftmengenmin / sum_luft
                        auslass.Luftmengennacht = mep.nb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft
                        auslass.Luftmengenmax = mep.zu_max.soll *1.0 * auslass.Luftmengenmin / sum_luft
                        auslass.Luftmengentnacht = mep.tnb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft


            if len(ab.keys()) == 1:
                
                for key1 in ab.keys():
                    for auslass in ab[key1]:
                        auslass.Luftmengenmin = mep.ab_min.soll *1.0 / len(ab[key1])
                        auslass.Luftmengennacht = mep.nb_ab.soll *1.0 / len(ab[key1])
            else:
                sum_luft = 0
                for key in ab.keys():
                    for auslass in ab[key]:
                        sum_luft += auslass.Luftmengenmin if auslass.Luftmengenmin > 0 else 0.01
                for key in ab.keys():
                    for auslass in ab[key]:
                        auslass.Luftmengenmin = mep.ab_min.soll *1.0 * auslass.Luftmengenmin / sum_luft
                        auslass.Luftmengennacht = mep.nb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft
                        auslass.Luftmengenmax = mep.ab_max.soll *1.0 * auslass.Luftmengenmin / sum_luft
                        auslass.Luftmengentnacht = mep.tnb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft

            for auslass in mep.list_auslass:
                # if auslass.elemid.ToString() == '22960632' or auslass.elemid.ToString() == '22960632':
                #     print(mep.Raumnr,mep.Raumnr)
                #     continue
                try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(auslass.Luftmengenmin).replace(',','.'))
                except:pass
                try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(auslass.Luftmengennacht).replace(',','.'))
                except:pass
                try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(auslass.Luftmengenmax).replace(',','.'))
                except:pass
                try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(auslass.Luftmengentnacht).replace(',','.'))
                except:pass
                try:auslass.elem.LookupParameter('IGF_X_Einbauort').Set(auslass.raumnummer)
                except:pass

        for mep in self.GUI_MEPRaum.mepraum_liste.values():
            list_vsr = mep.list_vsr
            for vsr in list_vsr:
                if vsr.art in ['RZU','RAB','LAB','24h','RUM']:
                    # if vsr.elemid.ToString() == '30921033' or vsr.elemid.ToString() == '00960622':
                    #     print(mep.Raumnr,mep.Raumnr)
                    #     continue
                    vsr.Luftmengenermitteln_new()
                    try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(vsr.Luftmengenmin).replace(',','.'))
                    except:pass
                    try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(vsr.Luftmengennacht).replace(',','.'))
                    except:pass
                    try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(vsr.Luftmengenmax).replace(',','.'))
                    except:pass
                    try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(vsr.Luftmengentnacht).replace(',','.'))
                    except:pass
                    try:vsr.elem.LookupParameter('IGF_X_Einbauort').Set(vsr.raumnummer)
                    except:pass
                    try:
                        numm = ''
                        for nummer in sorted(vsr.liste_Raum_nurNummer):
                            numm += nummer + ', '
                        vsr.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
                    except:pass
        t.Commit()
        t.Dispose()
        try:
            self.GUI_MEPRaum.lv_auslass.Items.Refresh()
            self.GUI_MEPRaum.lv_vsr.Items.Refresh()
            self.GUI_MEPRaum.Auswertung_System()
            self.GUI_MEPRaum.Auswertung_MEP()
            self.GUI_MEPRaum.anlageninfo.Items.Refresh()
            self.GUI_MEPRaum.grundinfo.Items.Refresh()
            self.GUI_MEPRaum.schachtinfo.Items.Refresh()
        except:pass

    def GetName(self):
        return "wert schreiben"

class VERTEILEN_MEP(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None

    def Execute(self,app):
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        self.GUI_MEPRaum.Trans+=1
        self.GUI_MEPRaum.TransMEPRaum.Insert(0,self.GUI_MEPRaum.mepraum)

        
        mep = self.GUI_MEPRaum.mepraum
            
        zu = {}
        ab = {}
        h24 = {}
        lab = {}
        for auslass in mep.list_auslass:
            if auslass.art == 'RZU':
                if auslass.familyandtyp not in zu.keys():
                    zu[auslass.familyandtyp] = [auslass]
                else:
                    zu[auslass.familyandtyp].append(auslass)
            if auslass.art == 'RAB':
                if auslass.familyandtyp not in ab.keys():
                    ab[auslass.familyandtyp] = [auslass]
                else:
                    ab[auslass.familyandtyp].append(auslass)
            if auslass.art == '24h':
                if auslass.familyandtyp not in h24.keys():
                    h24[auslass.familyandtyp] = [auslass]
                else:
                    h24[auslass.familyandtyp].append(auslass)
            if auslass.art == 'LAB':
                if auslass.familyandtyp not in lab.keys():
                    lab[auslass.familyandtyp] = [auslass]
                else:
                    lab[auslass.familyandtyp].append(auslass)
        
        if int(mep.ab_24h.soll) != int(mep.ab_24h.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist):
            print(30*'-')
            print('24h-Abluft oder Laborabluft in MEP-Raum stimmt nicht übereinandern.')
            print('24h-Abluft-soll: {} m³/h, 24h-Abluft-ist: {} m³/h'.format(mep.ab_24h.soll,mep.ab_24h.ist))
            print('Laborabluftmin-soll: {} m³/h, Laborabluftmin-ist: {} m³/h'.format(mep.ab_lab_min.soll,mep.ab_lab_min.ist))
            print('Laborabluftmax-soll: {} m³/h, Laborabluftmax-ist: {} m³/h'.format(mep.ab_lab_max.soll,mep.ab_lab_max.ist))
            print('Bitte manuell anpassen. ')

            return
        if (int(mep.zu_min.soll) > 0 or int(mep.zu_max.soll) > 0) and len(zu.keys()) == 0:
            print(30*'-')
            print('Es fehlt Zuluftauslass in MEP-Raum.')
            print('Zuluft min : {} m³/h, Zuluft max: {} m³/h'.format(mep.zu_min.soll,mep.zu_max.soll))
            print('Bitte manuell anpassen. ')
            return
        if (int(mep.ab_min.soll) > 0 or int(mep.ab_max.soll) > 0) and len(ab.keys()) == 0:
            print(30*'-')
            print('Es fehlt Abluftauslass in MEP-Raum.')
            print('Abluft min : {} m³/h, Abluft max: {} m³/h'.format(mep.ab_min.soll,mep.ab_max.soll))
            print('Bitte manuell anpassen. ')
            return

        if len(lab.keys()) > 0:
            for key in lab.keys():
                for auslass in lab[key]:
                    auslass.Luftmengennacht = auslass.Luftmengenmin
                    auslass.Luftmengentnacht = auslass.Luftmengenmin
        if len(h24.keys()) > 0:
            for key in h24.keys():
                for auslass in h24[key]:
                    auslass.Luftmengennacht = auslass.Luftmengenmin
                    auslass.Luftmengenmax = auslass.Luftmengenmin
                    auslass.Luftmengentnacht = auslass.Luftmengenmin
                    
        if len(zu.keys()) == 1:
            for key in zu.keys():
                for auslass in zu[key]:
                    auslass.Luftmengenmin = mep.zu_min.soll *1.0 / len(zu[key])
                    auslass.Luftmengennacht = mep.nb_zu.soll *1.0 / len(zu[key])
                    auslass.Luftmengenmax = mep.zu_max.soll *1.0 / len(zu[key])
                    auslass.Luftmengentnacht = mep.tnb_zu.soll *1.0 / len(zu[key])
        else:
            sum_luft = 0
            for key in zu.keys():
                for auslass in zu[key]:
                    sum_luft += auslass.Luftmengenmin if auslass.Luftmengenmin > 0 else 0.01
            for key in zu.keys():
                for auslass in zu[key]:
                    auslass.Luftmengenmin = mep.zu_min.soll *1.0 * auslass.Luftmengenmin / sum_luft
                    auslass.Luftmengennacht = mep.nb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft
                    auslass.Luftmengenmax = mep.zu_max.soll *1.0 * auslass.Luftmengenmin / sum_luft
                    auslass.Luftmengentnacht = mep.tnb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft


        if len(ab.keys()) == 1:
            
            for key1 in ab.keys():
                for auslass in ab[key1]:
                    auslass.Luftmengenmin = mep.ab_min.soll *1.0 / len(ab[key1])
                    auslass.Luftmengennacht = mep.nb_ab.soll *1.0 / len(ab[key1])
        else:
            sum_luft = 0
            for key in ab.keys():
                for auslass in ab[key]:
                    sum_luft += auslass.Luftmengenmin if auslass.Luftmengenmin > 0 else 0.01
            for key in ab.keys():
                for auslass in ab[key]:
                    auslass.Luftmengenmin = mep.ab_min.soll *1.0 * auslass.Luftmengenmin / sum_luft
                    auslass.Luftmengennacht = mep.nb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft
                    auslass.Luftmengenmax = mep.ab_max.soll *1.0 * auslass.Luftmengenmin / sum_luft
                    auslass.Luftmengentnacht = mep.tnb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft

        for auslass in mep.list_auslass:
            # if auslass.elemid.ToString() == '22960632' or auslass.elemid.ToString() == '22960632':
            #     print(mep.Raumnr,mep.Raumnr)
            #     continue
            try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(auslass.Luftmengenmin).replace(',','.'))
            except:pass
            try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(auslass.Luftmengennacht).replace(',','.'))
            except:pass
            try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(auslass.Luftmengenmax).replace(',','.'))
            except:pass
            try:auslass.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(auslass.Luftmengentnacht).replace(',','.'))
            except:pass
            try:auslass.elem.LookupParameter('IGF_X_Einbauort').Set(auslass.raumnummer)
            except:pass
        t.Start()
        list_vsr = mep.list_vsr
        for vsr in list_vsr:
            if vsr.art in ['RZU','RAB','LAB','24h','RUM']:
                # if vsr.elemid.ToString() == '30921033' or vsr.elemid.ToString() == '00960622':
                #     print(mep.Raumnr,mep.Raumnr)
                #     continue
                vsr.Luftmengenermitteln_new()
                try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(vsr.Luftmengenmin).replace(',','.'))
                except:pass
                try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(vsr.Luftmengennacht).replace(',','.'))
                except:pass
                try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(vsr.Luftmengenmax).replace(',','.'))
                except:pass
                try:vsr.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(vsr.Luftmengentnacht).replace(',','.'))
                except:pass
                try:vsr.elem.LookupParameter('IGF_X_Einbauort').Set(vsr.raumnummer)
                except:pass
                try:
                    numm = ''
                    for nummer in sorted(vsr.liste_Raum_nurNummer):
                        numm += nummer + ', '
                    vsr.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
                except:pass
        t.Commit()
        t.Dispose()
        self.GUI_MEPRaum.lv_auslass.Items.Refresh()
        self.GUI_MEPRaum.lv_vsr.Items.Refresh()
        self.GUI_MEPRaum.Auswertung_System()
        self.GUI_MEPRaum.Auswertung_MEP()
        self.GUI_MEPRaum.anlageninfo.Items.Refresh()
        self.GUI_MEPRaum.grundinfo.Items.Refresh()
        self.GUI_MEPRaum.schachtinfo.Items.Refresh()

    def GetName(self):
        return "wert schreiben"