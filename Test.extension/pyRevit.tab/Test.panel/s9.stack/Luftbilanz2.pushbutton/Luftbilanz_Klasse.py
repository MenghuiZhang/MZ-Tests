# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility,GridLength
import System
import xlsxwriter
from rpw import revit,DB,UI
from pyrevit import script,forms
from System.Windows.Input import ModifierKeys,Keyboard,Key
from System.Collections.Generic import List
from System.Windows.Media import Brushes
from clr import GetClrType
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
import os
from System.Windows.Forms import FolderBrowserDialog,DialogResult


logger = script.get_logger()
          
## Grundklass
class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

## Grundklass mit checked
class Itemtemplate(TemplateItemBase):
    def __init__(self,name,checked = False):
        TemplateItemBase.__init__(self)
        self.name = name
        self._checked = checked
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

## MEP für Export
class MEPFOREXPORTITEM(Itemtemplate):
    def __init__(self, name, berechnungnach,checked=False):
        Itemtemplate.__init__(self,name, checked)
        self.berechnng = berechnungnach

## No MEPSpace Exception
class NoMEPSpace(Exception):
    def __init__(self,elemid,typ,family,art):
        self.elemid = elemid
        self.typ = typ
        self.family = family
        self.art = art
        self.errorinfo = '{}: Einbauort konnte nicht ermittelt werden, FamilieName: {}, TypName: {}, ElementId: {}'.format(self.art,self.family,self.typ,self.elemid.ToString())
    
    def __str__(self):
        return self.errorinfo

class FamilieTemplate(TemplateItemBase):
    def __init__(self,elem):
        TemplateItemBase.__init__(self)
        self.elem = elem
        self.elemid = elem.Id
        self.ismuster =  self.Muster_Pruefen()

    def wert_schreiben_base(self,param,wert):
        if wert is not None:
            para = self.elem.LookupParameter(param)
            if para:
                if not para.IsReadOnly:
                    if para.StorageType.ToString() == 'String':
                        try:para.Set(str(wert)) 
                        except:pass
                    else:
                        try:para.SetValueString(str(wert)) 
                        except:pass

    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False

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

## Grundklasse für alle Familienexamplar
class FamilieExemplar(FamilieTemplate):
    def __init__(self,elem,art,DICT_MEP_NUMMER_NAME,logger):
        FamilieTemplate.__init__(self,elem)
        self.DICT_MEP_NUMMER_NAME = DICT_MEP_NUMMER_NAME
        self.logger = logger
        self.art_temp = art
        self._art = ''
        self._Luftmengenmin = 0
        self._Luftmengennacht = 0
        self._Luftmengenmax = 0
        self._Luftmengentnacht = 0
        self._size = ''
        self._Anmerkung = ''
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self._familyandtyp = self.familyname + ': ' + self.typname
        
        self.phase = self.elem.Document.GetElement(self.elem.CreatedPhaseId)
        self.raum = ''
        self.raumnummer = ''
        self.raumid = ''
        self.space = self.elem.Space[self.phase]
        self.GetRaum()
    
    @property
    def art(self):
        return self._art
    @art.setter
    def art(self,value):
        if value != self._art:
            self._art = value
            self.RaisePropertyChanged('art')
    
    @property
    def familyandtyp(self):
        return self._familyandtyp
    @familyandtyp.setter
    def familyandtyp(self,value):
        if value != self._familyandtyp:
            self._familyandtyp = value
            self.RaisePropertyChanged('familyandtyp')
    
    @property
    def size(self):
        return self._size
    @size.setter
    def size(self,value):
        if value != self._size:
            self._size = value
            self.RaisePropertyChanged('size')
    
    @property
    def Luftmengenmin(self):
        return self._Luftmengenmin
    @Luftmengenmin.setter
    def Luftmengenmin(self,value):
        if value != self._Luftmengenmin:
            self._Luftmengenmin = value
            self.RaisePropertyChanged('Luftmengenmin')
    
    @property
    def Luftmengenmax(self):
        return self._Luftmengenmax
    @Luftmengenmax.setter
    def Luftmengenmax(self,value):
        if value != self._Luftmengenmax:
            self._Luftmengenmax = value
            self.RaisePropertyChanged('Luftmengenmax')
    
    @property
    def Luftmengennacht(self):
        return self._Luftmengennacht
    @Luftmengennacht.setter
    def Luftmengennacht(self,value):
        if value != self._Luftmengennacht:
            self._Luftmengennacht = value
            self.RaisePropertyChanged('Luftmengennacht')
    
    @property
    def Luftmengentnacht(self):
        return self._Luftmengentnacht
    @Luftmengentnacht.setter
    def Luftmengentnacht(self,value):
        if value != self._Luftmengentnacht:
            self._Luftmengentnacht = value
            self.RaisePropertyChanged('Luftmengentnacht')
    
    @property
    def Anmerkung(self):
        return self._Anmerkung
    @Anmerkung.setter
    def Anmerkung(self,value):
        if value != self._Anmerkung:
            self._Anmerkung = value
            self.RaisePropertyChanged('Anmerkung')
    
    def GetRaum(self):
        if self.space:
            self.raumid = self.space.Id.ToString()
            self.raumnummer = self.space.Number
            self.raum = self.space.Number + ' - ' + self.space.LookupParameter('Name').AsString()
        else:
            if not self.ismuster:
                try:
                    param_einbauort = self.get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
                    if param_einbauort not in self.DICT_MEP_NUMMER_NAME.keys():
                        raise NoMEPSpace(self.elemid,self.typname,self.familyname,self.art_temp)
                    else:
                        self.raum = self.DICT_MEP_NUMMER_NAME[param_einbauort][0]
                        self.raumid = self.DICT_MEP_NUMMER_NAME[param_einbauort][1]
                except NoMEPSpace as e:
                    self.logger.error(str(e))
    
    def set_up(self):
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.familyandtyp = self.familyname + ': ' + self.typname
    
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
                    Height = int(round(conn.Height*304.8))
                    Width = int(round(conn.Width*304.8))
                    return str(max([Width,Height])) + 'x' + str(min([Width,Height]))
        except:return
    
    def wert_schreiben(self):
        self.wert_schreiben_base('IGF_RLT_AuslassVolumenstromMin',str(self.Luftmengenmin).replace(',','.'))
        self.wert_schreiben_base('IGF_RLT_AuslassVolumenstromMax',str(self.Luftmengenmax).replace(',','.'))
        self.wert_schreiben_base('IGF_RLT_AuslassVolumenstromNacht',str(self.Luftmengennacht).replace(',','.'))
        self.wert_schreiben_base('IGF_RLT_AuslassVolumenstromTiefeNacht',str(self.Luftmengentnacht).replace(',','.'))
        self.wert_schreiben_base('IGF_X_Einbauort',self.raumnummer)

## Klasse für Überstromsgruppe
class Baugruppe(FamilieTemplate):
    def __init__(self,elem,logger,DICT_MEP_NUMMER_NAME):
        FamilieTemplate.__init__(self,elem)
        self.logger = logger
        self.DICT_MEP_NUMMER_NAME = DICT_MEP_NUMMER_NAME
        if self.Pruefen():
            self.Volumen = self.get_value(self.elem.LookupParameter('IGF_RLT_Überströmung'))
            self.Eingang = ''
            self.Ausgang = ''
            try:self.Analyse()
            except Exception as e:self.logger.error(e)

    @property
    def auslaesse(self):
        auslass_liste = []
        for elemid in self.elem.GetMemberIds():
            elem = self.elem.Document.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008013' and elem.Symbol.FamilyName.upper().find("ÜBER") != -1:
                auslass_liste.append(elem)
        return auslass_liste
    
    def Pruefen(self):
        return len(self.auslaesse) == 2
   
    def Analyse(self):
        for elem in self.auslaesse:
            auslass = UeberStromAuslass(elem,self.Volumen,self.DICT_MEP_NUMMER_NAME,self.logger)
            if auslass.typ == 'Aus':
                self.Ausgang = auslass
            elif auslass.typ == 'Ein':
                self.Eingang = auslass
        self.Ausgang.anderesraum = self.Eingang.raum
        self.Eingang.anderesraum = self.Ausgang.raum

## Klasse für Überstrom
class UeberStromAuslass(FamilieExemplar):
    def __init__(self,elem,vol,DICT_MEP_NUMMER_NAME,logger):
        FamilieExemplar.__init__(self,elem,'Überströmung',DICT_MEP_NUMMER_NAME,logger)
        self._menge = vol
        self._anderesraum = ""
        self.typ = self.get_typ()
    
    @property
    def menge(self):
        return self._menge
    @menge.setter
    def menge(self,value):
        if value != self._menge:
            self._menge = value
            self.RaisePropertyChanged('menge')
    @property
    def anderesraum(self):
        return self._anderesraum
    @anderesraum.setter
    def anderesraum(self,value):
        if value != self._anderesraum:
            self._anderesraum = value
            self.RaisePropertyChanged('anderesraum')
        
    @property
    def conns(self):
        return list(self.elem.MEPModel.ConnectorManager.Connectors)

    def get_typ(self):
        conn = self.conns[0]
        if conn.Direction.ToString() == 'Out':
            return 'Aus'
        elif conn.Direction.ToString() == 'In':
            return 'Ein'

# ## Klasse für Auslässe
class Luftauslass(FamilieExemplar):
    """
    iris: '', -1, str(elemid)
    vsr = elemid
    vsrliste: []
    vsr kann nicht in vsrliste sein.
    wenn vsr in vsrliste ist, vsr ist falsch platziert.
    vsr:[]
    oder vsr: [vsr1,vsr2]
    """
    def __init__(self,elem,DICT_MEP_NUMMER_NAME,logger,LAB_AUSLASS_LISTE,H24_AUSLASS_LISTE,VSR_FAMILIE_LISTE,DRO_FAMILIE_LISTE,VSR_AUSLASS_LISTE,DRO_AUSLASS_LISTE):
        FamilieExemplar.__init__(self,elem,'Luftauslass',DICT_MEP_NUMMER_NAME,logger)

        self.Liste_LAB_Auslass = LAB_AUSLASS_LISTE
        self.Liste_24h_Auslass = H24_AUSLASS_LISTE
        self.VSR_FAMILIE_LISTE = VSR_FAMILIE_LISTE
        self.DRO_FAMILIE_LISTE = DRO_FAMILIE_LISTE
        self.VSR_AUSLASS_LISTE = VSR_AUSLASS_LISTE
        self.DRO_AUSLASS_LISTE = DRO_AUSLASS_LISTE

        self.enabledmin = True
        self.enabledmax = True
        self.enablednacht = True
        self.enabledtnacht = True

        self.System = ''
        self.AnlNr = ''
        self.vsr = ''
        self.RoutingListe = []
        self.Einbauteile = []
        self.VSR_Liste = []
        self.lufttyp = ''
        self.iris = ''
        self.VSR_Class = None
        self.Haupt_Class = None
        self.IRIS_Class = None

        self.Luftmengenermitteln()
        self.size = self.get_Size()

        try:self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:self.System = ''

        try:self.AnlNr = self.elem.Document.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId()).LookupParameter('IGF_X_AnlagenNr').AsValueString()
        except:self.AnlNr = ''
      
        if self.familyandtyp not in self.VSR_AUSLASS_LISTE and self.raumluftrelevant:
            self.get_RountingListe(self.elem)   

        self.get_Art()

        if self.familyname in self.DRO_AUSLASS_LISTE:
            if self.iris in [-1,'']:
                self.Anmerkung = 'Der Auslass ist von Drosselklappe geregelt aber keine ist damit angeschlossen.'

        else:
            if self.iris not in [-1,'']:
                self.Anmerkung = 'Der Auslass ist nicht von Drosselklappe geregelt aber eine ist damit angeschlossen.'

        if self.vsr in self.VSR_Liste:
            self.Anmerkung = 'VSR falsch angeschlossen. Grund kann sein, dass das Eingang des VSRs ist mit ein Zuluftgitter verbunden.'
        
        if not self.vsr:self.Anmerkung = 'Kein VSR.'
        
        self.RoutingListe = []
    
    @property
    def slavevon(self):
        if self.VSR_Class:
            return self.VSR_Class.vsrid
        elif self.Haupt_Class:
            return self.Haupt_Class.vsrid
        else:
            return ''

    @property
    def raumluftrelevant(self):
        return self.get_value('IGF_L_Raumluftunabhängig') == 0

    @property
    def nutzung(self):
        param0 = self.elem.Symbol.LookupParameter('IGF_X_Beschreibung')
        param1 = self.elem.Symbol.LookupParameter('IGF_X_Typ')
        param2 = self.elem.Symbol.LookupParameter('IGF_X_Größe')
        text0 = self.get_value(param0) if param0 else ''
        text1 = self.get_value(param1) if param1 else ''
        text2 = self.get_value(param2) if param2 else ''
        return '{}, {}, {}'.format(text0,text1,text2)

    def set_up(self):
        FamilieExemplar.set_up(self)
        self.Luftmengenermitteln()
        self.size = self.get_Size()
        self.get_Art()

    def Luftmengenermitteln(self):
        def _getLuftmengen(param):
            para = self.elem.LookupParameter(param)
            if not para:
                para = self.elem.Symbol.LookupParameter(param)
            if para:
                return round(float(self.get_value(para)),1)
            else:
                return 0
        
        self.Luftmengenmin = _getLuftmengen('IGF_RLT_AuslassVolumenstromMin')
        self.Luftmengenmax = _getLuftmengen('IGF_RLT_AuslassVolumenstromMax')
        self.Luftmengennacht = _getLuftmengen('IGF_RLT_AuslassVolumenstromNacht')
        self.Luftmengentnacht = _getLuftmengen('IGF_RLT_AuslassVolumenstromTiefeNacht')

    def get_RountingListe(self,element):
        if self.RoutingListe.Count > 300:return
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
                                if faminame == 'CAx RU Absperrklappe' and typname == 'mit Stellantrieb':
                                    return
                                if faminame in self.VSR_FAMILIE_LISTE:
                                    conns_temp = owner.MEPModel.ConnectorManager.Connectors
                                    for conn_temp in conns_temp:
                                        if self.lufttyp == 'ReturnAir':
                                            if conn_temp.Direction.ToString() == 'In' or conn_temp.Description == 'In':#???
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                        if self.lufttyp == 'SupplyAir':
                                            if conn_temp.Direction.ToString() == 'Out' or conn_temp.Description == 'Out':#???
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                    if not self.vsr:
                                        self.vsr = owner.Id.ToString()
                                    self.VSR_Liste.append(owner.Id.ToString())
                                    return
                                elif faminame in self.DRO_FAMILIE_LISTE:
                                    if not self.iris:self.iris = owner.Id.ToString()
                                    
                            self.get_RountingListe(owner)
    
    def get_Art(self):
        if self.System:
            if self.System.upper().find('TIERHALTUNG') != -1:
                if self.System.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif self.System.upper().find('ABLUFT') != -1:
                    self.art = 'RAB' 
            else:
                if self.System.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif self.System.upper().find('ABLUFT') != -1:
                    self.art = 'RAB'
                else:
                    systemklassifizierung = self.elem.LookupParameter('Systemklassifizierung').AsString()
                    if systemklassifizierung.upper().find('ZULUFT') != -1:
                        self.art = 'RZU'
                    elif systemklassifizierung.upper().find('ABLUFT') != -1:
                        self.art = 'RAB'
        else:
            self.art = 'XXX'
       
        if self.familyandtyp in self.Liste_LAB_Auslass:
            self.art = 'LAB'
            self.enabledmin = False
            self.enabledmax = False
            self.enablednacht = False
            self.enabledtnacht = False
        elif self.familyandtyp in self.Liste_24h_Auslass:
            self.art = '24h'
            self.enabledmin = False
            self.enabledmax = False
            self.enablednacht = False
            self.enabledtnacht = False
        
## Klasse für Volumenstromregler
class VSR(FamilieExemplar):
    def __init__(self,elemid,DICT_MEP_NUMMER_NAME,logger,DICT_DatenBank,Liste_Fabrikat,Liste_Datenbank,Liste_Datenbank1):
        FamilieExemplar.__init__(self,elemid,'VSR',DICT_MEP_NUMMER_NAME,logger)
        self.checked = False
        self.Auslass = ObservableCollection[Luftauslass]()
        self.liste_Raum = []
        self.slavevon = '-'
        self.liste_Raum_nurNummer = []
        self.size = self.get_Size()
        self.IsHaupt = False
        self.IsIris = False

        self.List_Iris = ObservableCollection[VSR]()
        self.List_VSR = ObservableCollection[VSR]()
        self.List_Haupt = ObservableCollection[VSR]()

        self._Liste_Herstellertyp = []

        self.DICT_DatenBank = DICT_DatenBank
        self._Liste_Fabrikat = Liste_Fabrikat
        self._Herstellerindex = 0
        self._Herstellertypindex = 0

        self.Datenbank_rund = Liste_Datenbank 
        self.Datenbank_eck = Liste_Datenbank1 
        self.VSR_Hersteller = None

        self._anmerkung = 'kein passender Typ gefunden'
        self.material = self.elem.LookupParameter('IGF_X_Material_Text')
        self._dict = {
            'VR1-80':DB.ElementId(40884949),
            'VR1-100':DB.ElementId(30817223),
            'VR1-125':DB.ElementId(30817226),
            'VR1-160':DB.ElementId(30817239),
            'VR1-200':DB.ElementId(30817229),
            'VR1-250':DB.ElementId(30817232),
            'VR1-315':DB.ElementId(30817235),
            'VRE1-100':DB.ElementId(30778826),
            'VRE1-125':DB.ElementId(30781828),
            'VRE1-160':DB.ElementId(30784831),
            'VRE1-200':DB.ElementId(29973863),
            'VRE1-250':DB.ElementId(29970867),
            'VRE1-315':DB.ElementId(29960042),
            'VRL1-80':DB.ElementId(40884949),
            'VRL1-100':DB.ElementId(30817223),
            'VRL1-125':DB.ElementId(30817226),
            'VRL1-160':DB.ElementId(30817239),
            'VRL1-200':DB.ElementId(30817229),
            'VRL1-250':DB.ElementId(30817232),
        }

        self._Vmin = ''
        self._Vmax = ''
        self.nutzung = ''
    
    @property
    def BKSschild(self):
        return self.get_value(self.elem.LookupParameter('IGF_GA_BKS-Schild'))

    @property
    def vsrid(self):
        return self.get_value(self.elem.LookupParameter('IGF_X_Bauteilnummerierung'))
    
    @property
    def Vmin(self):
        return self._Vmin
    @Vmin.setter
    def Vmin(self,value):
        if value != self._Vmin:
            self._Vmin = value
            self.RaisePropertyChanged('Vmin')
    @property
    def Vmax(self):
        return self._Vmax
    @Vmax.setter
    def Vmax(self,value):
        if value != self._Vmax:
            self._Vmax = value
            self.RaisePropertyChanged('Vmax')
    @property
    def Herstellerindex(self):
        return self._Herstellerindex
    @Herstellerindex.setter
    def Herstellerindex(self,value):
        if value != self._Herstellerindex:
            self._Herstellerindex = value
            self.RaisePropertyChanged('Herstellerindex')
    @property
    def Liste_Herstellertyp(self):
        return self._Liste_Herstellertyp
    @Liste_Herstellertyp.setter
    def Liste_Herstellertyp(self,value):
        if value != self._Liste_Herstellertyp:
            self._Liste_Herstellertyp = value
            self.RaisePropertyChanged('Liste_Herstellertyp')
    @property
    def Liste_Fabrikat(self):
        return self._Liste_Fabrikat
    @Liste_Fabrikat.setter
    def Liste_Fabrikat(self,value):
        if value != self._Liste_Fabrikat:
            self._Liste_Fabrikat = value
            self.RaisePropertyChanged('Liste_Fabrikat')
    @property
    def Herstellertypindex(self):
        return self._Herstellertypindex
    @Herstellertypindex.setter
    def Herstellertypindex(self,value):
        if value != self._Herstellertypindex:
            self._Herstellertypindex = value
            self.RaisePropertyChanged('Herstellertypindex')

    @property
    def anmerkung(self):
        return self._anmerkung
    @anmerkung.setter
    def anmerkung(self,value):
        if value != self._anmerkung:
            self._anmerkung = value
            self.RaisePropertyChanged('anmerkung')
    
    @property
    def ispps(self):
        if self.material:
            if self.material.AsString():return self.material.AsString().upper().find('PPS') != -1
            else:return False
        else:return False
    @property
    def vsrart(self):
        if self.familyandtyp.upper().find('KONST') != -1 or self.familyandtyp.upper().find('KVR') != -1:
            return 'KVR'
        else:
            return 'VVR'

    def changetype(self):
        if self.VSR_Hersteller:
            if self.VSR_Hersteller.typ in self._dict.keys():
                try:self.elem.ChangeTypeId(self._dict[self.VSR_Hersteller.typ])
                except Exception as e:logger.error(e)
        
        if self.ispps:self.material.Set('PPs')
        else:self.material.Set('Blech')
        
        try:self.set_up()
        except Exception as e:logger.error(e)
    
    def vsrueberpruefen(self):
        if self.VSR_Hersteller:
            self.Vmin = self.VSR_Hersteller.vmin
            self.Vmax = self.VSR_Hersteller.vmax
            if self.VSR_Hersteller.dimension != self.size:
                self.anmerkung = 'In Projekt {} modelliert'.format(self.size)
            else:
                self.anmerkung = ''
        else:self.anmerkung = 'kein passender Typ gefunden'
            
    def get_vsrListe(self):
        self.VSR_Hersteller = None
        self.Vmax = 0
        self.Vmin = 0
        self.Liste_Herstellertyp = []

        liste_volumen = [self.Luftmengenmax,self.Luftmengenmin,self.Luftmengennacht,self.Luftmengentnacht]
        while( 0 in liste_volumen):
            liste_volumen.remove(0)
        if not liste_volumen:
            self.VSR_Hersteller = None
            self.Herstellertypindex = -1
            return

        minvol = min(liste_volumen)
        maxvol = max(liste_volumen)
        
        if self.Herstellerindex != -1:
            Hersteller = self.Liste_Fabrikat[self.Herstellerindex]

        if self.size.find('DN') != -1:
            if Hersteller in self.Datenbank_rund.keys():
                liste = self.Datenbank_rund[Hersteller]
            else:
                self.VSR_Hersteller = None
                self.Herstellertypindex = -1
                return
        else:
            if Hersteller in self.Datenbank_eck.keys():liste = self.Datenbank_eck[Hersteller]
            else:
                self.VSR_Hersteller = None
                self.Herstellertypindex = -1
                return

        if self.art == 'RZU':
            for art in [1,0]:
                if art in liste.keys():
                    listetemp = liste[art]
                    if self.vsrart in listetemp.keys():
                        listetemp_vsr = listetemp[self.vsrart]
                        if self.size.find('DN') != -1:
                            for dimension in sorted(listetemp_vsr.keys()):
                                for vsr_temp in listetemp_vsr[dimension]:
                                    if vsr_temp.vmin <= minvol and vsr_temp.vmax > maxvol:
                                        
                                        self.Liste_Herstellertyp.append(vsr_temp.typ)
                                        
                        else:
                            for max_a in sorted(listetemp_vsr.keys()):
                                for min_a in sorted(listetemp_vsr[max_a].keys()):
                                    for vsr_temp in listetemp_vsr[max_a][min_a]:
                                        if vsr_temp.vmin <= minvol and vsr_temp.vmax > maxvol:
                                            
                                            self.Liste_Herstellertyp.append(vsr_temp.typ)
        elif self.art in ['RAB','LAB','24h']:
            if self.ispps:
                if 3 in liste.keys():
                    listetemp = liste[3]
                    if self.vsrart in listetemp.keys():
                        listetemp_vsr = listetemp[self.vsrart]
                        if self.size.find('DN') != -1:
                            for dimension in sorted(listetemp_vsr.keys()):
                                for vsr_temp in listetemp_vsr[dimension]:
                                    if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                        self.Liste_Herstellertyp.append(vsr_temp.typ)
                        else:
                            for max_a in sorted(listetemp_vsr.keys()):
                                for min_a in sorted(listetemp_vsr[max_a].keys()):
                                    for vsr_temp in listetemp_vsr[max_a][min_a]:
                                        if vsr_temp.vmin <= minvol and vsr_temp.vmax > maxvol:
                                            self.Liste_Herstellertyp.append(vsr_temp.typ)

            else:
                for art in [2,0]:
                    if art in liste.keys():
                        listetemp = liste[art]
                        if self.vsrart in listetemp.keys():
                            listetemp_vsr = listetemp[self.vsrart]
                            if self.size.find('DN') != -1:
                                for dimension in sorted(listetemp_vsr.keys()):
                                    for vsr_temp in listetemp_vsr[dimension]:
                                        if vsr_temp.vmin <= minvol and vsr_temp.vmax > maxvol:
                                            self.Liste_Herstellertyp.append(vsr_temp.typ)
                            else:
                                for max_a in sorted(listetemp_vsr.keys()):
                                    for min_a in sorted(listetemp_vsr[max_a].keys()):
                                        for vsr_temp in listetemp_vsr[max_a][min_a]:
                                            if vsr_temp.vmin <= minvol and vsr_temp.vmax > maxvol:
                                                self.Liste_Herstellertyp.append(vsr_temp.typ)
        
    def vsrauswaelten(self):
        self.get_vsrListe()
        if len(self.Liste_Herstellertyp) > 0:
            if self.size.find('DN') != -1:
                self.VSR_Hersteller = self.DICT_DatenBank[self.Liste_Herstellertyp[0]]
            else:
                temp_dict = {}
                for typ in self.Liste_Herstellertyp:
                    item_vsr = self.DICT_DatenBank[typ]
                    temp_dict[item_vsr.dimension] = item_vsr
                if self.size in temp_dict.keys():
                    self.VSR_Hersteller = temp_dict[self.size]
                else:
                    self.VSR_Hersteller = self.DICT_DatenBank[self.Liste_Herstellertyp[0]]

        if self.VSR_Hersteller:
            if self.VSR_Hersteller.typ in self.Liste_Herstellertyp:
                self.Herstellertypindex = self.Liste_Herstellertyp.index(self.VSR_Hersteller.typ)

    def set_up(self):
        FamilieExemplar.set_up(self)        
        self.size = self.get_Size()
        self.get_Art()
        self.vsrauswaelten()
        self.vsrueberpruefen()

    def Luftmengenermitteln(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengenmax = 0
        Luftmengennacht = 0
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
        Luftmengenmax = 0
        Luftmengennacht = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
            Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
            Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
            Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def Luftmengenermitteln_um(self):
        Luftmengenmin = 0
        Luftmengenmax = 0
        Luftmengennacht = 0
        Luftmengentnacht = 0

        for auslass in self.Auslass:
            if auslass.art == 'UMZU':
                Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
                Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
                Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
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
            elif 'RAB' in art_liste and ('LAB' in art_liste or '24h' in art_liste):
                self.art = 'RAB'
                print(self.elemid.ToString(),art_liste)
            elif 'LAB' in art_liste and '24h' in art_liste:
                self.art = 'LAB'
                print(self.elemid.ToString(),art_liste)
            else:
                print(self.elemid.ToString(),art_liste)

        elif art_liste.Count == 3 and '24h' in art_liste and 'LAB' in art_liste and 'RAB' in art_liste:self.art = 'RAB'

        else:print(self.elemid.ToString(),art_liste)

        if self.art == 'LAB' or self.art == '24h':
            for auslass in self.Auslass:
                self.nutzung = auslass.nutzung
                break

    def wert_schreiben(self):
        FamilieExemplar.wert_schreiben()
        numm = ''
        for nummer in sorted(self.liste_Raum_nurNummer):
            numm += nummer + ', '
        if numm:self.wert_schreiben_base('IGF_X_Wirkungsort',numm[:-2])
        else:self.wert_schreiben_base('IGF_X_Wirkungsort','')
    
class SchachtGrundinfo(object):
    def __init__(self,name):
        self.name = name

class MEPGrundInfo(TemplateItemBase):
    def __init__(self,zustand,name,soll,tooltip,Liste):
        TemplateItemBase.__init__(self)
        self.zustand = zustand
        self.name = name
        self._soll = soll
        self._ist = ''
        self._reduziert = soll
        self.tooltip = tooltip
        self.Liste = Liste
    @property
    def soll(self):
        return self._soll
    @soll.setter
    def soll(self,value):
        if value != self._soll:
            self._soll = value
            self.RaisePropertyChanged('soll')
    @property
    def reduziert(self):
        return self._reduziert
    @reduziert.setter
    def reduziert(self,value):
        if value != self._reduziert:
            self._reduziert = value
            self.RaisePropertyChanged('reduziert')
    @property
    def ist(self):
        return self._ist
    @ist.setter
    def ist(self,value):
        if value != self._ist:
            self._ist = value
            self.RaisePropertyChanged('ist')
        
    def _get_ist(self):
        summe = 0
        if self.name.find('Ange.') != -1:
            return self.reduziert
        elif self.name.find('ZU_SUM') != -1:
            if self.zuatand == 'min':
                return
            else:
                return
        elif self.name.find('von') != -1 or self.name.find('bis') != -1 or self.name.find('Dauer') != -1:
            return self.reduziert
        elif self.name.find('AB_SUM') != -1:
            if self.zuatand == 'min':
                return
            else:
                return
        
            return self.reduziert
        elif self.name.find('SUM') != -1:
            return self.reduziert
        elif self.name.find('Ange.') != -1:
            return self.reduzier
        elif self.name.find('Ange.') != -1:
            return self.reduziert
        

class MEPSchachtInfo(TemplateItemBase):
    def __init__(self,name,nr,menge,SchachtListe = []):
        TemplateItemBase.__init__(self)
        self.name = name
        self.nr = nr
        self._menge = menge
        self.SchachtListe = SchachtListe
        self._schachtindex = -1
        self.get_index()
    
    @property
    def menge(self):
        return self._menge
    @menge.setter
    def menge(self,value):
        if value != self._menge:
            self._menge = value
            self.RaisePropertyChanged('menge')
    
    @property
    def schachtindex(self):
        return self._schachtindex
    @schachtindex.setter
    def schachtindex(self,value):
        if value != self._schachtindex:
            self._schachtindex = value
            self.RaisePropertyChanged('schachtindex')

    def get_index(self):
        if self.nr in self.SchachtListe:

            try:
                self.schachtindex = self.SchachtListe.index(self.nr)
            except:
                self.schachtindex = -1
        else:
            self.schachtindex = -1

class MEPAnlagenInfo(TemplateItemBase):
    def __init__(self,name,mep_nr,mep_mengen,sys_nr,sys_mengen):
        TemplateItemBase.__init__(self)
        self.name = name
        self._mep_nr = mep_nr
        self._mep_mengen = mep_mengen
        self._sys_nr = sys_nr
        self._sys_mengen = sys_mengen 
    @property
    def mep_nr(self):
        return self._mep_nr
    @mep_nr.setter
    def mep_nr(self,value):
        if value != self._mep_nr:
            self._mep_nr = value
            self.RaisePropertyChanged('mep_nr')
    @property
    def mep_mengen(self):
        return self._mep_mengen
    @mep_mengen.setter
    def mep_mengen(self,value):
        if value != self._mep_mengen:
            self._mep_mengen = value
            self.RaisePropertyChanged('mep_mengen')
    @property
    def sys_nr(self):
        return self._sys_nr
    @sys_nr.setter
    def sys_nr(self,value):
        if value != self._sys_nr:
            self._sys_nr = value
            self.RaisePropertyChanged('sys_nr')
    @property
    def sys_mengen(self):
        return self._sys_mengen
    @sys_mengen.setter
    def sys_mengen(self,value):
        if value != self._sys_mengen:
            self._sys_mengen = value
            self.RaisePropertyChanged('sys_mengen')

class MEPAuswertung(TemplateItemBase):
    def __init__(self,zustand,name,Liste0,Liste1):
        TemplateItemBase.__init__(self)
        self.name = name
        self.Liste0 = Liste0
        self.Liste1 = Liste1
        self._ist = 0
        self._reduziert = 0
        self.zustand = zustand
        self._set_up()

    @property
    def reduziert(self):
        return self._reduziert
    @reduziert.setter
    def reduziert(self,value):
        if value != self._reduziert:
            self._reduziert = value
            self.RaisePropertyChanged('reduziert')

    @property
    def ist(self):
        return self._ist
    @ist.setter
    def ist(self,value):
        if value != self._ist:
            self._ist = value
            self.RaisePropertyChanged('ist')
    
    def _set_up(self):
        self.ist = self._get_ist()
        self.reduziert = self._get_reduziert()
    
    def _get_reduziert(self):
        if self.name.find('Summe') != -1:
            summe = 0
            for el in self.Liste0:
                summe += el.reduziert
            return summe
        elif self.name.find('Luftbilanz') != -1:
            zuluft = 0
            abluft = 0
            for el in self.Liste0:
                if el.name.find('Zuluft') != -1:
                    zuluft = el.reduziert
                elif el.name.find('Abluft') != -1:
                    abluft = el.reduziert
            return zuluft-abluft
        
        elif self.name.find('Auswertung') != -1:
            summe = 0
            Bilanz = 0
            Druck = 0
            for el in self.Liste0:
                if el.name.find('Luftbilanz') != -1:
                    Bilanz = el.reduziert
                elif el.name.find('Druckstufe') != -1:
                    Druck = el.reduziert
            summe = Bilanz - Druck
            if abs(summe) <= 3:
                return 'OK'
            else:return 'Passt nicht'
        elif self.name.find('Tierhaltung') != -1:
            return ''
        else:
            return self.Liste0.reduziert
        
    def _get_ist(self):
        if self.name.find('Summe') != -1:
            summe = 0
            for el in self.Liste0:
                summe += el.ist
            return summe
        elif self.name.find('Luftbilanz') != -1:
            zuluft = 0
            abluft = 0
            for el in self.Liste0:
                if el.name.find('Zuluft') != -1:
                    zuluft = el.ist
                elif el.name.find('Abluft') != -1:
                    abluft = el.ist
            return zuluft-abluft
        elif self.name.find('Manuel') != -1:
            return self.Liste0.reduziert
        
        elif self.name.find('Tierhaltung') != -1:
            return ''
        
        elif self.name.find('Auswertung') != -1:
            summe = 0
            Bilanz = 0
            Druck = 0
            for el in self.Liste0:
                if el.name.find('Luftbilanz') != -1:
                    Bilanz = el.ist
                elif el.name.find('Druckstufe') != -1:
                    Druck = el.ist
            summe = Bilanz - Druck
            if abs(summe) <= 3:
                return 'OK'
            else:return 'Passt nicht'

        elif self.name.find('Über') != -1:
            summe = 0
            
            for el in self.Liste1:
                summe += el.menge
            return summe

        else:
            summe = 0
            for el in self.Liste1:
                if self.zustand == 'min':
                    summe += el.Luftmengenmin
                elif self.zustand == 'max':
                    summe += el.Luftmengenmax
                elif self.zustand == 'nacht':
                    summe += el.Luftmengennacht
                elif self.zustand == 'tnacht':
                    summe += el.Luftmengentnacht
            return summe

class MEPRaum(object):
    def __init__(self, elem,list_vsr,LISTE_SCHACHT,logger,DICT_MEP_AUSLASS,DICT_MEP_UEBERSTROM,Dict_Ueber_Manuell):
        self.logger = logger
        self.Dict_Ueber_Manuell = Dict_Ueber_Manuell
        self.DICT_MEP_AUSLASS = DICT_MEP_AUSLASS
        self.DICT_MEP_UEBERSTROM = DICT_MEP_UEBERSTROM
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
                                logger.error('VSR {} {} hat zwei Haupt VSR {} {}, {} {}. Bitte überprüfen.'.format(subvsr.elemid,subvsr.vsrid,mainvsr.vsrid,subvsr.slavevon))
                        subvsr.slavevon = mainvsr.vsrid
                        if subvsr.vsrid.find('-->') == -1:
                            subvsr.vsrid = '--> ' + subvsr.vsrid
                        
        self.IsTierRaum = (self.get_element('IGF_RLT_Tierkäfig_raumunabhängig') != 0)
        self.IsSchacht = (self.get_element('TGA_RLT_InstallationsSchacht') != 0)
        self.list_auslass = ObservableCollection[Luftauslass]()
        if self.elemid.ToString() in self.DICT_MEP_AUSLASS.keys():
             for art in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()].keys()):
                for fam in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()][art].keys()):
                    for terminal in DICT_MEP_AUSLASS[self.elemid.ToString()][art][fam]:
                        self.list_auslass.Add(terminal)
       
        try:self.list_ueber = self.DICT_MEP_UEBERSTROM[self.elemid.ToString()]
        except:self.list_ueber = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        
        # Grundinfo.
        self.bezugsnummer = self.elem.LookupParameter('TGA_RLT_VolumenstromProNummer').AsValueString()
        try:self.bezugsname = self.berechnung_nach[self.bezugsnummer]
        except:self.bezugsname = 'keine'
        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''
        self.flaeche = round(self.get_element('Fläche'), 2)
        # self.volumen = round(self.get_element('IGF_A_Volumen'),2)
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('TGA_RLT_VolumenstromProFaktor')
        try:self.einheit = self.einheit_liste[self.bezugsnummer]
        except:self.einheit = ''
        self.hoehe = int(self.get_element('IGF_A_Lichte_Höhe'))
        self.volumen = round((self.hoehe*self.flaeche/1000),2)
        # self.hoehe = round((self.volumen/self.flaeche),2)
        self.nachtbetrieb = self.get_element('IGF_RLT_Nachtbetrieb')
        self.tiefenachtbetrieb = self.get_element('IGF_RLT_TieferNachtbetrieb')
        self.NB_LW = self.get_element('IGF_RLT_NachtbetriebLW')
        self.T_NB_LW = self.get_element('IGF_RLT_TieferNachtbetriebLW')

        self.Liste_RZU = []
        self.Liste_RAB = []
        self.Liste_TZU = []
        self.Liste_TAB = []
        self.Liste_H24 = []
        self.Liste_LAB = []
        self.Liste_UIN = []
        self.Liste_UAU = []
        self.Luftauslass_Analyse()
        
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
        self.Detail_Tnahct = ObservableCollection[MEPAuswertung]()

        # Übersicht 
        self.angezuluft = MEPGrundInfo('Ange.Zuluft',self.get_element('Angegebener Zuluftstrom'),'Angegebener Zuluftstrom')
        self.angeabluft = MEPGrundInfo('Ange.Abluft',self.get_element('Angegebener Abluftluftstrom'),'Angegebener Abluftluftstrom')
        self.ab_24h = MEPGrundInfo('24h Ab',self.get_element('IGF_RLT_AbluftSumme24h'),'IGF_RLT_AbluftSumme24h')
        self.ab_minsum = MEPGrundInfo('min.AB_SUM',0,'RAB,LAB,24h,Über,Über_M')
        self.ab_min = MEPGrundInfo('min.RAB',self.get_element('IGF_RLT_AbluftminRaum'),'IGF_RLT_AbluftminRaum-IGF_RLT_AbluftminSummeLabor-IGF_RLT_AbluftSumme24h')
        self.ab_lab_min = MEPGrundInfo('min.LAB',self.get_element('IGF_RLT_AbluftminSummeLabor'),'IGF_RLT_AbluftminSummeLabor')
        self.zu_minsum = MEPGrundInfo('min.ZU_SUM',0,'RZU,Über,Über_M')
        self.zu_min = MEPGrundInfo('min.RZU',self.get_element('IGF_RLT_ZuluftminRaum'),'IGF_RLT_ZuluftminRaum')
        self.ab_maxsum = MEPGrundInfo('max.AB_SUM',0,'RAB,LAB,24h,Über,Über_M')
        self.ab_max = MEPGrundInfo('max.RAB',self.get_element('IGF_RLT_AbluftmaxRaum'),'IGF_RLT_AbluftmaxRaum-IGF_RLT_AbluftmaxSummeLabor-IGF_RLT_AbluftSumme24h')
        self.ab_lab_max = MEPGrundInfo('max.LAB',self.get_element('IGF_RLT_AbluftmaxSummeLabor'),'IGF_RLT_AbluftmaxSummeLabor')
        self.zu_maxsum = MEPGrundInfo('max.ZU_SUM',0,'RZU,Über,Über_M')
        self.zu_max = MEPGrundInfo('max.RZU',self.get_element('IGF_RLT_ZuluftmaxRaum'),'IGF_RLT_ZuluftmaxRaum')
        self.nb_von = MEPGrundInfo('NB von',self.get_element('IGF_RLT_NachtbetriebVon'),'IGF_RLT_NachtbetriebVon')
        self.nb_bis = MEPGrundInfo('NB bis',self.get_element('IGF_RLT_NachtbetriebBis'),'IGF_RLT_NachtbetriebBis')
        self.nb_dauer = MEPGrundInfo('NB Dauer',self.get_element('IGF_RLT_NachtbetriebDauer'),'IGF_RLT_NachtbetriebDauer')
        self.nb_zu = MEPGrundInfo('NB Zu',self.get_element('IGF_RLT_ZuluftNachtRaum'),'IGF_RLT_ZuluftNachtRaum')
        self.nb_ab = MEPGrundInfo('NB Ab',self.get_element('IGF_RLT_AbluftNachtRaum'),'IGF_RLT_AbluftNachtRaum')
        self.tnb_von = MEPGrundInfo('TNB von',self.get_element('IGF_RLT_TieferNachtbetriebVon'),'IGF_RLT_TieferNachtbetriebVon')
        self.tnb_bis = MEPGrundInfo('TNB bis',self.get_element('IGF_RLT_TieferNachtbetriebBis'),'IGF_RLT_TieferNachtbetriebBis')
        self.tnb_dauer = MEPGrundInfo('TNB Dauer',self.get_element('IGF_RLT_TieferNachtbetriebDauer'),'IGF_RLT_TieferNachtbetriebDauer')
        self.tnb_zu = MEPGrundInfo('TNB Zu',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'),'IGF_RLT_ZuluftTieferNachtRaum')
        self.tnb_ab = MEPGrundInfo('TNB Ab',self.get_element('IGF_RLT_AbluftTieferNachtRaum'),'IGF_RLT_AbluftTieferNachtRaum')
        self.tier_zu_min = MEPGrundInfo('min.TZU',self.get_element('IGF_RLT_Luftmenge_min_TZU'),'IGF_RLT_Luftmenge_min_TZU')
        self.tier_ab_min = MEPGrundInfo('min.TAB',self.get_element('IGF_RLT_Luftmenge_min_TAB'),'IGF_RLT_Luftmenge_min_TAB')
        self.tier_zu_max = MEPGrundInfo('max.TZU',self.get_element('IGF_RLT_Luftmenge_max_TZU'),'IGF_RLT_Luftmenge_max_TZU')
        self.tier_ab_max = MEPGrundInfo('max.TAB',self.get_element('IGF_RLT_Luftmenge_max_TAB'),'IGF_RLT_Luftmenge_max_TAB')
        self.ueber_sum = MEPGrundInfo('Überstrom',self.get_element('IGF_RLT_ÜberströmungRaum'),'IGF_RLT_ÜberströmungRaum')
        self.ueber_in = MEPGrundInfo('Über. Ein.',self.get_element('IGF_RLT_ÜberströmungSummeIn'),'IGF_RLT_ÜberströmungSummeIn')
        self.ueber_aus = MEPGrundInfo('Über. Aus.',self.get_element('IGF_RLT_ÜberströmungSummeAus'),'IGF_RLT_ÜberströmungSummeAus')
        self.ueber_in_manuell = MEPGrundInfo('Über.Ein.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        self.ueber_aus_manuell = MEPGrundInfo('Über.Aus.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        self.Druckstufe = MEPGrundInfo('Druckstufe',self.get_element('IGF_RLT_RaumDruckstufeEingabe'),'IGF_RLT_RaumDruckstufeEingabe')

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
        
        try:self.ueber_aus_manuell.soll = Dict_Ueber_Manuell[self.elem.Number]
        except:self.ueber_aus_manuell.soll = 0
        
        self.ueber_aus_manuell.ist =self.ueber_aus_manuell.soll
        self.ueber_in_manuell.ist = self.ueber_in_manuell.soll

        self.ab_min.soll = self.ab_min.soll - self.ab_24h.soll - self.ab_lab_min.soll
        self.ab_max.soll = self.ab_max.soll - self.ab_24h.soll - self.ab_lab_max.soll
        self.labnacht = 0
        self.labtnacht = 0
        self.ab24nacht = 0
        self.ab24tnacht = 0
        self.sum_update()
        
        self.IGF_Legende = ''     
            
        self.Tagesbetrieb_Berechnen()
        self.Nachtbetrieb_Berechnen()
        self.Druckstufe_Berechnen()
    
        
        self.get_Anlagen_info()
        self.get_Schacht_info()

        self.D_T_MIN_Rzu   = MEPAuswertung('min','Zu-Raum',self.zu_min,self.Liste_RZU)
        self.D_T_MIN_ZuUe  = MEPAuswertung('min','Zu-Über',self.ueber_in,self.Liste_UIN)
        self.D_T_MIN_ZuUeM = MEPAuswertung('min','Zu-Über-Manuel',self.ueber_in_manuell,[])
        self.D_T_MIN_Rab   = MEPAuswertung('min','Ab-Raum',self.ab_min,self.Liste_RAB)
        self.D_T_MIN_AbUe  = MEPAuswertung('min','Ab-Über',self.ueber_aus,self.Liste_UAU)
        self.D_T_MIN_AbUeM = MEPAuswertung('min','Ab-Über-Manuel',self.ueber_aus_manuell,[])
        self.D_T_MIN_Lab   = MEPAuswertung('min','Ab-Labor',self.ab_lab_min,self.Liste_LAB)
        self.D_T_MIN_24h   = MEPAuswertung('min','Ab-24h',self.ab_24h,self.Liste_H24)
        self.D_T_MIN_ZuS   = MEPAuswertung('min','Zuluft-Summe',[self.D_T_MIN_Rzu,self.D_T_MIN_ZuUe,self.D_T_MIN_ZuUeM],[])
        self.D_T_MIN_AbS   = MEPAuswertung('min','Abluft-Summe',[self.D_T_MIN_Rab,self.D_T_MIN_AbUe,self.D_T_MIN_AbUeM,self.D_T_MIN_24h,self.D_T_MIN_Lab],[])
        self.D_T_MIN_BLZ   = MEPAuswertung('min','Luftbilanz',[self.D_T_MIN_ZuS,self.D_T_MIN_AbS],[])
        self.D_T_Druckstufe= MEPAuswertung('min','Druckstufe',self.Druckstufe,[])
        self.D_T_MIN_Aus   = MEPAuswertung('min','Auswertung',[self.D_T_MIN_BLZ,self.D_T_Druckstufe],[])
        self.D_T_MIN_Tier  = MEPAuswertung('min','Tierhaltung',[],[])
        self.D_T_MIN_TZu   = MEPAuswertung('min','Zu-TH',self.tier_zu_min,self.Liste_TZU)
        self.D_T_MIN_TAb   = MEPAuswertung('min','Ab-TH',self.tier_ab_min,self.Liste_TAB)
        self.Detail_Min.Add(self.D_T_MIN_Rzu)
        self.Detail_Min.Add(self.D_T_MIN_ZuUe)
        self.Detail_Min.Add(self.D_T_MIN_ZuUeM)
        self.Detail_Min.Add(self.D_T_MIN_ZuS)
        self.Detail_Min.Add(self.D_T_MIN_Rab)
        self.Detail_Min.Add(self.D_T_MIN_AbUe)
        self.Detail_Min.Add(self.D_T_MIN_AbUeM)
        self.Detail_Min.Add(self.D_T_MIN_Lab)
        self.Detail_Min.Add(self.D_T_MIN_24h)
        self.Detail_Min.Add(self.D_T_MIN_AbS)
        self.Detail_Min.Add(self.D_T_MIN_BLZ)
        self.Detail_Min.Add(self.D_T_Druckstufe)
        self.Detail_Min.Add(self.D_T_MIN_Aus)
        self.Detail_Min.Add(self.D_T_MIN_Tier)
        self.Detail_Min.Add(self.D_T_MIN_TZu)
        self.Detail_Min.Add(self.D_T_MIN_TAb)

        self.D_T_MAX_Rzu   = MEPAuswertung('max','Zu-Raum',self.zu_max,self.Liste_RZU)
        self.D_T_MAX_ZuUe  = MEPAuswertung('max','Zu-Über',self.ueber_in,self.Liste_UIN)
        self.D_T_MAX_ZuUeM = MEPAuswertung('max','Zu-Über-Manuel',self.ueber_in_manuell,[])
        self.D_T_MAX_Rab   = MEPAuswertung('max','Ab-Raum',self.ab_max,self.Liste_RAB)
        self.D_T_MAX_AbUe  = MEPAuswertung('max','Ab-Über',self.ueber_aus,self.Liste_UAU)
        self.D_T_MAX_AbUeM = MEPAuswertung('max','Ab-Über-Manuel',self.ueber_aus_manuell,[])
        self.D_T_MAX_Lab   = MEPAuswertung('max','Ab-Labor',self.ab_lab_max,self.Liste_LAB)
        self.D_T_MAX_24h   = MEPAuswertung('max','Ab-24h',self.ab_24h,self.Liste_H24)
        self.D_T_MAX_ZuS   = MEPAuswertung('max','Zuluft-Summe',[self.D_T_MAX_Rzu,self.D_T_MAX_ZuUe,self.D_T_MAX_ZuUeM],[])
        self.D_T_MAX_AbS   = MEPAuswertung('max','Abluft-Summe',[self.D_T_MAX_Rab,self.D_T_MAX_AbUe,self.D_T_MAX_AbUeM,self.D_T_MAX_24h,self.D_T_MAX_Lab],[])
        self.D_T_MAX_BLZ   = MEPAuswertung('max','Luftbilanz',[self.D_T_MAX_ZuS,self.D_T_MAX_AbS],[])
        self.D_T_MAX_Aus   = MEPAuswertung('max','Auswertung',[self.D_T_MAX_BLZ,self.D_T_Druckstufe],[])
        self.D_T_MAX_Tier  = MEPAuswertung('max','Tierhaltung',[],[])
        self.D_T_MAX_TZu   = MEPAuswertung('max','Zu-TH',self.tier_zu_max,self.Liste_TZU)
        self.D_T_MAX_TAb   = MEPAuswertung('max','Ab-TH',self.tier_ab_max,self.Liste_TAB)
        self.Detail_Max.Add(self.D_T_MAX_Rzu)
        self.Detail_Max.Add(self.D_T_MAX_ZuUe)
        self.Detail_Max.Add(self.D_T_MAX_ZuUeM)
        self.Detail_Max.Add(self.D_T_MAX_ZuS)
        self.Detail_Max.Add(self.D_T_MAX_Rab)
        self.Detail_Max.Add(self.D_T_MAX_AbUe)
        self.Detail_Max.Add(self.D_T_MAX_AbUeM)
        self.Detail_Max.Add(self.D_T_MAX_Lab)
        self.Detail_Max.Add(self.D_T_MAX_24h)
        self.Detail_Max.Add(self.D_T_MAX_AbS)
        self.Detail_Max.Add(self.D_T_MAX_BLZ)
        self.Detail_Max.Add(self.D_T_Druckstufe)
        self.Detail_Max.Add(self.D_T_MAX_Aus)
        self.Detail_Max.Add(self.D_T_MAX_Tier)
        self.Detail_Max.Add(self.D_T_MAX_TZu)
        self.Detail_Max.Add(self.D_T_MAX_TAb)

        
    
    def update(self):        
        self.ab_24h.soll = self.get_element('IGF_RLT_AbluftSumme24h')
        self.ab_lab_min.soll = self.get_element('IGF_RLT_AbluftminSummeLabor') 
        self.Druckstufe.soll = self.get_element('IGF_RLT_RaumDruckstufeEingabe')
    
    def sum_update(self):
        self.ab_minsum.soll = str(int(round(float(self.ab_min.soll)+float(self.ab_lab_min.soll)+float(self.ab_24h.soll)+float(self.ueber_aus.soll)+float(self.ueber_aus_manuell.soll)))) + \
            ' (' + str(int(self.ab_min.soll)) + ', ' + str(int(self.ab_lab_min.soll)) + ', ' + str(int(self.ab_24h.soll)) + ', ' + str(int(self.ueber_aus.soll)) + ', ' + str(int(self.ueber_aus_manuell.soll))+ ')'
        
        self.ab_maxsum.soll = str(int(round(float(self.ab_max.soll)+float(self.ab_lab_max.soll)+float(self.ab_24h.soll)+float(self.ueber_aus.soll)+float(self.ueber_aus_manuell.soll)))) + \
            ' (' + str(int(self.ab_max.soll)) + ', ' + str(int(self.ab_lab_max.soll)) + ', ' + str(int(self.ab_24h.soll)) + ', ' + str(int(self.ueber_aus.soll)) + ', ' + str(int(self.ueber_aus_manuell.soll))+ ')'
        
        self.zu_minsum.soll = str(int(round(float(self.zu_min.soll)+float(self.ueber_in.soll)+float(self.ueber_in_manuell.soll)))) + \
            ' (' + str(int(self.zu_min.soll)) + ', ' + str(int(self.ueber_in.soll)) + ', ' + str(int(self.ueber_in_manuell.soll))+ ')'
        
        self.zu_maxsum.soll = str(int(round(float(self.zu_max.soll)+float(self.ueber_in.soll)+float(self.ueber_in_manuell.soll)))) + \
            ' (' + str(int(self.zu_max.soll)) + ', ' + str(int(self.ueber_in.soll)) + ', ' + str(int(self.ueber_in_manuell.soll)) + ')'
        
        # self.ab_minsum.ist = str(int(round(float(self.ab_min.ist)+float(self.ab_lab_min.ist)+float(self.ab_24h.ist)+float(self.ueber_aus.ist)+float(self.ueber_aus_manuell.ist)))) + \
        #     ' (' + str(int(self.ab_min.ist)) + ', ' + str(int(self.ab_lab_min.ist)) + ', ' + str(int(self.ab_24h.ist)) + ', ' + str(int(self.ueber_aus.ist)) + ', ' + str(int(self.ueber_aus_manuell.ist))+ ')'
        
        # self.ab_maxsum.ist = str(int(round(float(self.ab_max.ist)+float(self.ab_lab_max.ist)+float(self.ab_24h.ist)+float(self.ueber_aus.ist)+float(self.ueber_aus_manuell.ist)))) + \
        #     ' (' + str(int(self.ab_max.ist)) + ', ' + str(int(self.ab_lab_max.ist)) + ', ' + str(int(self.ab_24h.ist)) + ', ' + str(int(self.ueber_aus.ist)) + ', ' + str(int(self.ueber_aus_manuell.ist))+ ')'
        
        # self.zu_minsum.ist = str(int(round(float(self.zu_min.ist)+float(self.ueber_in.ist)+float(self.ueber_in_manuell.ist)))) + \
        #     ' (' + str(int(self.zu_min.ist)) + ', ' + str(int(self.ueber_in.ist)) + ', ' + str(int(self.ueber_in_manuell.ist))+ ')'
        
        # self.zu_maxsum.ist = str(int(round(float(self.zu_max.ist)+float(self.ueber_in.ist)+float(self.ueber_in_manuell.ist)))) + \
        #     ' (' + str(int(self.zu_max.ist)) + ', ' + str(int(self.ueber_in.ist)) + ', ' + str(int(self.ueber_in_manuell.ist)) + ')'

    def luft_round(self,luft):
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Druckstufe_Berechnen(self):
        n = abs(int(self.Druckstufe.soll/10)) if abs(int(self.Druckstufe.soll/10)) < 6 else 5
        if self.Druckstufe.soll > 0:self.IGF_Legende = n*'+'
        else:self.IGF_Legende = n*'-'
          
    def Labmin_24h_Druckstufe_Pruefen(self,zuluftmin):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll
        
        
        # if self.Druckstufe.soll < 0:
        #     if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll + self.Druckstufe.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #         zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.Druckstufe.soll + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # else:
        # if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #     zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # if self.Druckstufe.soll > 0:
        #     zuluftmin = zuluftmin + self.Druckstufe.soll 

        return zuluftmin
    
    def Labmin_24h_Druckstufe_Nacht_Pruefen(self,zuluftmin):

        zuluftmin =  zuluftmin - self.ueber_in_manuell.soll - self.ueber_in.soll
        # if self.Druckstufe.soll < 0:
        #     if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll + self.Druckstufe.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #         zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.Druckstufe.soll + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # else:
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        if zuluftmin < 0:
            zuluftmin = 0
        # if self.Druckstufe.soll > 0:
        #     zuluftmin = zuluftmin + self.Druckstufe.soll 

        return zuluftmin

    def Labmax_24h_Druckstufe_Pruefen(self,zuluftmax):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax :
            zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
        
        # if self.Druckstufe.soll < 0:
        #     if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll + self.Druckstufe.soll > zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #         zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.Druckstufe.soll  - self.ueber_in_manuell.soll - self.ueber_in.soll
        # else:
        # if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #     zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # if self.Druckstufe.soll > 0:
        #     zuluftmax = zuluftmax + self.Druckstufe.soll 

        return zuluftmax

    def Nachtbetrieb_Berechnen(self):
        if self.bezugsname in ['Fläche',"Luftwechsel","Person","manuell"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = self.luft_round(self.T_NB_LW * self.volumen)
                
                    self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll + 24.00
                self.nb_zu.soll = self.luft_round(self.NB_LW * self.volumen)
                self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_zu.soll = 0   
        elif self.bezugsname in ['nurZU_Fläche',"nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = 0-(self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll )#- self.Druckstufe.soll)
                    # self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = 0

                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0
                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll + 24.00
                self.nb_zu.soll = 0-(self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll )#- self.Druckstufe.soll)
                # self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = 0

            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_zu.soll = 0   
        elif self.bezugsname in ['nurAB_Fläche',"nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = 0
                    # self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll# - self.Druckstufe.soll
                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll + 24.00
                self.nb_zu.soll = 0
                # self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_zu.soll = 0   
        else:
            self.nb_dauer.soll = 0
            self.nb_zu.soll = 0
            self.nb_ab.soll = 0
            self.tnb_dauer.soll = 0
            self.tnb_ab.soll = 0
            self.tnb_zu.soll = 0   

    def Tagesbetrieb_Berechnen(self):
        if self.flaeche == 0:
            return
        if self.bezugsname in ['nurZU_Fläche',"nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
            if (self.ab_lab_min.soll + self.ab_24h.soll> 0) or (self.ab_lab_max.soll + self.ab_24h.soll> 0):
                self.logger.error("Berechnungsprinzip von Raum {} ist Falsch. Der Raum ist nur über Überströmung ausströmt aber hat Laborabluft min: {}, max: {} m³/h und 24h-Abluft: {} m³/h".format(self.Raumnr,self.ab_lab_min.soll,self.ab_lab_max.soll,self.ab_24h.soll))
                return
            if self.bezugsname == "nurZU_Fläche":
                self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor))
            elif self.bezugsname == "nurZU_Luftwechsel":
                self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor))
            
            elif self.bezugsname == "nurZU_Person":
                self.zu_min.soll = self.luft_round(self.personen * float(self.faktor))
            elif self.bezugsname == "nurZUMa":
                self.zu_min.soll = self.luft_round(float(self.faktor))
            
            self.ab_min.soll = self.ab_max.soll = 0
            self.zu_max.soll = self.zu_min.soll

            self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen(self.zu_min.soll)
            self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen(self.zu_max.soll)
            self.angeabluft.soll = self.angezuluft.soll = self.zu_min.soll
            
            abweichung = self.ueber_aus_manuell.soll + self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll #+ self.Druckstufe.soll
            
            if abweichung >= 0:
                self.zu_min.soll += abweichung
                self.zu_max.soll += abweichung


            else:
                self.hinweis = "Achtung: Bitte Überströmung-Aus um {} m³/h erhöhen".format(int(0 - abweichung))
                self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))

        elif self.bezugsname in ['nurAB_Fläche',"nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
            if self.bezugsname == "nurAB_Fläche":
                self.angezuluft.soll = self.luft_round(self.flaeche * float(self.faktor))
            elif self.bezugsname == "nurAB_Luftwechsel":
                self.angezuluft.soll = self.luft_round(self.volumen * float(self.faktor))
            elif self.bezugsname == "nurAB_Person":
                self.angezuluft.soll = self.luft_round(self.personen * float(self.faktor))
            elif self.bezugsname == "nurABMa":
                self.angezuluft.soll = self.luft_round(float(self.faktor))
                        
            self.angeabluft.soll = self.angezuluft.soll
            self.zu_min.soll = self.zu_max.soll = 0
            
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
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor))
            elif self.bezugsname == "Luftwechsel":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor))
            elif self.bezugsname == "Person":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.personen * float(self.faktor))
            elif self.bezugsname == "manuell":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(float(self.faktor))
        
            
            if self.bezugsname == 'manuell':
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.zu_min.soll)
                self.zu_max.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.zu_max.soll)
            else:
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen(self.zu_min.soll)
                self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen(self.zu_max.soll)

            self.angeabluft.soll = self.angezuluft.soll = self.zu_max.soll
            

            self.ab_max.soll = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll #- self.Druckstufe.soll
            self.ab_min.soll = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            
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
            
        self.get_Schacht_info()
        self.get_Anlagen_info()
        self.sum_update()

    def Luftauslass_Analyse(self):
        Liste_RZU = []
        Liste_RAB = []
        Liste_TZU = []
        Liste_TAB = []
        Liste_H24 = []
        Liste_LAB = []
        Liste_UIN = self.list_ueber["Ein"]
        Liste_UAU = self.list_ueber["Aus"]

        for auslass in self.list_auslass:
            if auslass.art == '24h':Liste_H24.append(auslass)
            elif auslass.art == 'LAB':Liste_LAB.append(auslass)
            elif auslass.art == 'RZU':Liste_RZU.append(auslass)
            elif auslass.art == 'RAB':Liste_RAB.append(auslass)
            elif auslass.art == 'TZU':Liste_TZU.append(auslass)
            elif auslass.art == 'TAB':Liste_TAB.append(auslass)

        self.Liste_UAU = Liste_UAU
        self.Liste_UIN = Liste_UIN
        self.Liste_LAB = Liste_LAB
        self.Liste_H24 = Liste_H24
        self.Liste_TAB = Liste_TAB
        self.Liste_TZU = Liste_TZU
        self.Liste_RAB = Liste_RAB
        self.Liste_RZU = Liste_RZU
                
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
                    self.zu_max.soll,anl,Dict['RZU'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_RZU') or self.get_element('IGF_RLT_Luftmenge_RZU'):
            self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                self.zu_max.soll,'',''))
        if 'RAB' in Dict.keys():
            for anl in sorted(Dict['RAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                    self.ab_max.soll,anl,Dict['RAB'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_RAB') or self.get_element('IGF_RLT_Luftmenge_RAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                self.ab_max.soll,'',''))
        if 'TZU' in Dict.keys():
            for anl in sorted(Dict['TZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                    self.tier_zu_max.soll,anl,Dict['TZU'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_TZU') or self.get_element('IGF_RLT_Luftmenge_TZU'):
            self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                self.tier_zu_max.soll,'',''))
        
        if 'TAB' in Dict.keys():
            for anl in sorted(Dict['TAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                    self.tier_ab_max.soll,anl,Dict['TAB'][anl]))
        else:
            # if self.get_element('IGF_RLT_AnlagenNr_TAB') or self.get_element('IGF_RLT_Luftmenge_TAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                self.tier_ab_max.soll,'',''))
        
        if '24h' in Dict.keys():
            for anl in sorted(Dict['24h'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                    self.ab_24h.soll,anl,Dict['24h'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_24h') or self.get_element('IGF_RLT_Luftmenge_24h'):
            self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                self.ab_24h.soll,'',''))
           
        if 'LAB' in Dict.keys():
            for anl in sorted(Dict['LAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                    self.ab_lab_max.soll,anl,Dict['LAB'][anl]))  
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_LAB') or self.get_element('IGF_RLT_Luftmenge_LAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                self.ab_lab_max.soll,'',''))
        
    def get_Schacht_info(self):
        self.Schacht_info.Clear()
        self.rzu_Schacht = MEPSchachtInfo('RZU',self.get_element('TGA_RLT_SchachtZuluft'),self.zu_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.rzu_Schacht)
        self.rab_Schacht = MEPSchachtInfo('RAB',self.get_element('TGA_RLT_SchachtAbluft'),self.ab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.rab_Schacht)
        
        self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.tier_zu_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.tzu_Schacht)
        self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.tier_ab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.tab_Schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('TGA_RLT_Schacht24hAbluft'),self.ab_24h.soll,self.liste_schacht)
        self.Schacht_info.Add(self._24h_Schacht)
        self.lab_Schacht = MEPSchachtInfo('LAB',self.get_element('IGF_RLT_Schacht_LAB'),self.ab_lab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.lab_Schacht)
        
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

        wert_schreiben("Angegebener Zuluftstrom", self.angezuluft.soll)
        wert_schreiben("Angegebener Abluftluftstrom", self.angeabluft.soll)
        wert_schreiben("IGF_RLT_AbluftminRaum", self.ab_min.soll+self.ab_lab_min.soll+self.ab_24h.soll+self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_AbluftmaxRaum", self.ab_max.soll+self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zu_min.soll)
        wert_schreiben("IGF_RLT_ZuluftmaxRaum", self.zu_max.soll)
        wert_schreiben("IGF_A_Lichte_Höhe", self.hoehe)
        wert_schreiben("IGF_A_Volumen", self.volumen)

        wert_schreiben2("TGA_RLT_VolumenstromProName", self.bezugsname)
        wert_schreiben("TGA_RLT_VolumenstromProEinheit", self.einheit)
        wert_schreiben("TGA_RLT_VolumenstromProNummer", self.bezugsnummer)
        wert_schreiben("TGA_RLT_VolumenstromProFaktor", float(self.faktor))

        wert_schreiben2("IGF_RLT_Nachtbetrieb", self.nachtbetrieb)
        wert_schreiben("IGF_RLT_NachtbetriebLW", self.NB_LW)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.nb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.nb_ab.soll +self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_NachtbetriebVon", self.nb_von.soll)
        wert_schreiben("IGF_RLT_NachtbetriebBis", self.nb_bis.soll)

        wert_schreiben2("IGF_RLT_TieferNachtbetrieb", self.tiefenachtbetrieb)
        wert_schreiben("IGF_RLT_TieferNachtbetriebLW", self.T_NB_LW)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer", self.tnb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tnb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tnb_ab.soll +self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebVon", self.tnb_von.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebBis", self.tnb_bis.soll)

        wert_schreiben("IGF_RLT_AbluftSumme24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_AbluftminRaumL24h", self.ab_24h.soll)    
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
        wert_schreiben("IGF_RLT_Luftmenge_min_TAB", self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TZU", self.tier_zu_min.soll)
        
        wert_schreiben3('TGA_RLT_SchachtZuluft',self.rzu_Schacht)
        wert_schreiben3('TGA_RLT_SchachtAbluft',self.rab_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TZU',self.tzu_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TAB',self.tab_Schacht)
        wert_schreiben3('TGA_RLT_Schacht24hAbluft',self._24h_Schacht)
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


