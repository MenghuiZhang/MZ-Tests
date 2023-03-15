# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from rpw import DB
from pyrevit import script
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs

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
        self.Luftmengenmin = vol
        self.Luftmengenmax = vol
        self.Luftmengennacht = vol
        self.Luftmengentnacht = vol
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

class UeberStromAuslass_Manuell(object):
    def __init__(self,vol=0):
        self.menge = vol
        self.Luftmengenmin = vol
        self.Luftmengennacht = vol
        self.Luftmengenmax = vol
        self.Luftmengentnacht = vol

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

        self.Einbauteile_Liste = []

        self.System = ''
        self.AnlNr = ''
        self.vsr = ''
        self.RoutingListe = []
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

        if not self.raumluftrelevant:
            self.art = 'xxx'

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
        return self.get_value(self.elem.LookupParameter('IGF_L_Raumluftunabhängig')) != 1

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
        if self.RoutingListe.Count > 150:return
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
                            if owner.Category.Name in ['Luftkanalzubehör','HLS-Bauteile']:
                                faminame = owner.Symbol.FamilyName
                                if faminame not in self.DRO_AUSLASS_LISTE and faminame not in self.VSR_FAMILIE_LISTE:
                                    self.Einbauteile_Liste.append(owner.Id.ToString())
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
class Einbauteile(FamilieExemplar):
    def __init__(self,elemid,DICT_MEP_NUMMER_NAME,logger,auslassliste = []):
        FamilieExemplar.__init__(self,elemid,'Einbauteil',DICT_MEP_NUMMER_NAME,logger)
        self.Auslass = ObservableCollection[Luftauslass](auslassliste)
        self.size = self.get_Size()

    @property
    def Text(self):
        def luftmeng(luftmeng):
            if luftmeng is not None:
                luftmeng = str(int(round(float(str(luftmeng).replace(',','.')))))
                while(len(luftmeng) < 5):
                   luftmeng = '0' + luftmeng
                return luftmeng
            else:
                luftmeng = '00000'
        luftmengmin = luftmeng(self.Luftmengenmin)
        luftmengmax = luftmeng(self.Luftmengenmax)
        luftmengnacht = luftmeng(self.Luftmengennacht)
        luftmengtiefe = luftmeng(self.Luftmengentnacht)       
        return self.art + '_' + self.familyname + '_' + self.typname + '_' + self.size + '_' + luftmengmin + '_' + luftmengmax + '_' + luftmengnacht + '_' + luftmengtiefe
        
    def set_up(self):
        FamilieExemplar.set_up(self)        
        self.size = self.get_Size()
        self.get_Art()

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
            elif 'LAB' in art_liste and '24h' in art_liste:
                self.art = 'LAB'

 
        elif art_liste.Count == 3 and '24h' in art_liste and 'LAB' in art_liste and 'RAB' in art_liste:
            self.art = 'RAB'

        else:print(self.elemid.ToString(),art_liste)

        if self.art == 'LAB' or self.art == '24h':
            for auslass in self.Auslass:
                self.nutzung = auslass.nutzung
                break
    
class VSR(FamilieExemplar):
    def __init__(self,elemid,DICT_MEP_NUMMER_NAME,logger,DICT_DatenBank,Liste_Fabrikat,Liste_Datenbank,Liste_Datenbank1,Dict_Herstellertyp):
        FamilieExemplar.__init__(self,elemid,'VSR',DICT_MEP_NUMMER_NAME,logger)
        self.checked = False
        self.Auslass = ObservableCollection[Luftauslass]()
        self.liste_Raum = []
        self.slavevon = '-'
        self.liste_Raum_nurNummer = []
        self.size = self.get_Size()
        self.IsHaupt = False
        self.IsIris = False
        self.vsrid = self.get_value(self.elem.LookupParameter('IGF_X_Bauteilnummerierung'))

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
        self._dict = Dict_Herstellertyp
        self._Vmin = ''
        self._Vmax = ''
        self.nutzung = ''
    
    @property
    def BKSschild(self):
        try: return self.get_value(self.elem.LookupParameter('IGF_GA_BKS-Schild'))
        except:return ''
    
    @property
    def Text(self):
        def luftmeng(luftmeng):
            if luftmeng is not None:
                luftmeng = str(int(round(float(str(luftmeng).replace(',','.')))))
                while(len(luftmeng) < 5):
                   luftmeng = '0' + luftmeng
                return luftmeng
            else:
                luftmeng = '00000'
        luftmengmin = luftmeng(self.Luftmengenmin)
        luftmengmax = luftmeng(self.Luftmengenmax)
        luftmengnacht = luftmeng(self.Luftmengennacht)
        luftmengtiefe = luftmeng(self.Luftmengentnacht)       
        return self.art + '_' + self.familyname + '_' + self.typname + '_' + self.size + '_' + luftmengmin + '_' + luftmengmax + '_' + luftmengnacht + '_' + luftmengtiefe
       
    
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

    @property
    def Text(self):
        def luftmeng(luftmeng):
            if luftmeng is not None:
                luftmeng = str(int(round(float(str(luftmeng).replace(',','.')))))
                while(len(luftmeng) < 5):
                   luftmeng = '0' + luftmeng
                return luftmeng
            else:
                luftmeng = '00000'
        luftmengmin = luftmeng(self.Luftmengenmin)
        luftmengmax = luftmeng(self.Luftmengenmax)
        luftmengnacht = luftmeng(self.Luftmengennacht)
        luftmengtiefe = luftmeng(self.Luftmengentnacht)       
        return self.art + '_' + self.familyname + '_' + self.typname + '_' + self.size + '_' + luftmengmin + '_' + luftmengmax + '_' + luftmengnacht + '_' + luftmengtiefe
        
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
    
    def Luftmengenermitteln_new1(self,zustand):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengen = 0
        if zustand == 'min':
            for auslass in self.Auslass:
                Luftmengen += float(str(auslass.Luftmengenmin).replace(',', '.'))
            self.Luftmengenmin = int(round(Luftmengen))
        elif zustand == 'max':
            for auslass in self.Auslass:
                Luftmengen += float(str(auslass.Luftmengenmax).replace(',', '.'))
            self.Luftmengenmax = int(round(Luftmengen))
        elif zustand == 'nacht':
            for auslass in self.Auslass:
                Luftmengen += float(str(auslass.Luftmengennacht).replace(',', '.'))
            self.Luftmengennacht = int(round(Luftmengen))
        elif zustand == 'tnacht':
            for auslass in self.Auslass:
                Luftmengen += float(str(auslass.Luftmengentnacht).replace(',', '.'))
            self.Luftmengentnacht = int(round(Luftmengen))
    
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
        FamilieExemplar.wert_schreiben(self)
        numm = ''
        for nummer in sorted(self.liste_Raum_nurNummer):
            numm += nummer + ', '
        if numm:self.wert_schreiben_base('IGF_X_Wirkungsort',numm[:-2])
        else:self.wert_schreiben_base('IGF_X_Wirkungsort','')
    
class SchachtGrundinfo(object):
    def __init__(self,name):
        self.name = name

class MEPGrundInfo(TemplateItemBase):
    def __init__(self,name,tooltip,soll = 0,Liste = []):
        TemplateItemBase.__init__(self)
        self.name = name
        self._soll = soll
        self._ist = soll
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
    def ist(self):
        return self._ist
    @ist.setter
    def ist(self,value):
        if value != self._ist:
            self._ist = value
            self.RaisePropertyChanged('ist')
        
    def _get_ist(self):
        summe = 0
        if self.name.find('Ange.') != -1 or self.name.find('.M.') != -1 or self.name.find('Druckstufe') != -1:
            return
        elif self.name.find('ZU_SUM') != -1:
            rzu = 0
            Ueber = 0
            Ueber_M = 0
            for el in self.Liste:
                if el.name.find('RZU') != -1:
                    rzu = el.ist
                elif el.name.find('Über.Ein.M') != -1:
                    Ueber_M = el.ist
                else:
                    Ueber = el.ist
            summe = int(round(rzu + Ueber + Ueber_M))
            self.ist = '{}({},{},{})'.format(summe,int(round(rzu)),int(round(Ueber)),int(round(Ueber_M)))
        elif self.name.find('AB_SUM') != -1:
            rab = 0
            lab = 0
            h24 = 0
            Ueber = 0
            Ueber_M = 0
            for el in self.Liste:
                if el.name.find('RAB') != -1:
                    rab = el.ist
                elif el.name.find('Über.Aus.M') != -1:
                    Ueber_M = el.ist
                elif el.name.find('LAB') != -1:
                    lab = el.ist
                elif el.name.find('24h') != -1:
                    h24 = el.ist
                else:
                    Ueber = el.ist
            summe = int(round(rab + Ueber + Ueber_M + lab + h24))
            self.ist = '{}({},{},{},{},{})'.format(summe,int(round(rab)),int(round(lab)),int(round(h24)),int(round(Ueber)),int(round(Ueber_M)))
        elif self.name.find('von') != -1 or self.name.find('bis') != -1 or self.name.find('Dauer') != -1:
            return
        elif self.name == 'Überstrom':
            Zu = 0
            Ab = 0
            for el in self.Liste:
                if el.name.find('Ein') != -1:
                    Zu += el.ist
                else:
                    Ab += el.ist
            self.ist = int(round(Zu-Ab)) 
        elif self.name.find('tnb.') != -1:
            if self.name.find('RZU') != -1 or self.name.find('RAB') != -1 or self.name.find('LAB') != -1 or self.name.find('24h') != -1 or self.name == 'nb.Über_in' or self.name == 'nb.Über_aus':
                for el in self.Liste:
                    summe += el.Luftmengentnacht
                self.ist = int(round(summe))
            elif self.name.find('Druck') == -1:
                self.ist = self.soll
            else:
                self.ist = self.soll
        elif self.name.find('nb.') != -1:
            if self.name.find('RZU') != -1 or self.name.find('RAB') != -1 or self.name.find('LAB') != -1 or self.name.find('24h') != -1 or self.name == 'nb.Über_in' or self.name == 'nb.Über_aus':
                for el in self.Liste:
                    summe += el.Luftmengennacht
                self.ist = int(round(summe))
            elif self.name.find('Druck') == -1:
                self.ist = self.soll
            else:
                self.ist = self.soll
        else:
            for el in self.Liste:
                if self.name.find('min') != -1:
                    summe += el.Luftmengenmin
                elif self.name.find('max') != -1:
                    summe += el.Luftmengenmax
                elif self.name == 'NB Zu':
                    summe += el.Luftmengennacht
                elif self.name == 'NB Ab':
                    summe += el.Luftmengennacht
                elif self.name == 'TNB Zu':
                    summe += el.Luftmengentnacht
                elif self.name == 'TNB Ab':
                    summe += el.Luftmengentnacht
                else:
                    summe += el.Luftmengenmin
            self.ist = int(round(summe))
    
    def _get_soll(self):
        if self.name.find('.M.') != -1 or self.name.find('Druckstufe') != -1 or self.name.find('Ange.') != -1:
            self.ist = self.soll
        elif self.name.find('TZU') != -1 or self.name.find('TAB') != -1:
            if not self.soll:self.soll = 0
        
        elif self.name.find('ZU_SUM') != -1:
            rzu = 0
            Ueber = 0
            Ueber_M = 0
            for el in self.Liste:
                if el.name.find('RZU') != -1:
                    rzu = el.soll
                elif el.name.find('Über.Ein.M') != -1:
                    Ueber_M = el.soll
                else:
                    Ueber = el.soll
            summe = rzu + Ueber + Ueber_M
            self.soll = '{}({},{},{})'.format(int(round(summe)),int(round(rzu)),int(round(Ueber)),int(round(Ueber_M)))

        elif self.name.find('AB_SUM') != -1:
            rab = 0
            lab = 0
            h24 = 0
            Ueber = 0
            Ueber_M = 0
            for el in self.Liste:
                if el.name.find('RAB') != -1:
                    rab = el.soll
                elif el.name.find('Über.Aus.M') != -1:
                    Ueber_M = el.soll
                elif el.name.find('LAB') != -1:
                    lab = el.soll
                elif el.name.find('24h') != -1:
                    h24 = el.soll
                else:
                    Ueber = el.soll
            summe = rab + Ueber + Ueber_M + lab + h24
            self.soll = '{}({},{},{},{},{})'.format(int(round(summe)),int(round(rab)),int(round(lab)),int(round(h24)),int(round(Ueber)),int(round(Ueber_M)))

        elif self.name.find('von') != -1 or self.name.find('bis') != -1 or self.name.find('Dauer') != -1:
            self.ist = self.soll
        

    def _set_up_soll(self):
        self._get_soll()
    def _set_up_ist(self):
        self._get_ist()
     
class MEPSchachtInfo(TemplateItemBase):
    def __init__(self,name,nr,elem,SchachtListe = []):
        TemplateItemBase.__init__(self)
        self.name = name
        self.nr = nr
        self.elem = elem
        self._menge = 0
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
    
    def _set_up(self):
        self.menge = self.elem.soll

class MEPAnlagenInfo(TemplateItemBase):
    def __init__(self,name,mep_nr,elem,sys_nr,liste):
        TemplateItemBase.__init__(self)
        self.name = name
        self.elem = elem
        self.liste = liste
        self._mep_nr = mep_nr
        self._mep_mengen = 0
        self._sys_nr = sys_nr
        self._sys_mengen = 0 
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
    
    def _get_soll(self):
        self.mep_mengen = self.elem.soll
    
    def _get_ist(self):
        summe = 0
        for el in self.liste:
            summe += el.Luftmengenmax
        self.sys_mengen = int(round(summe))
    
    def _set_up(self):
        self._get_soll()
        self._get_ist()

class MEPAuswertung(TemplateItemBase):
    def __init__(self,name,Liste0):
        TemplateItemBase.__init__(self)
        self.name = name
        self.Liste0 = Liste0
        self._ist = 0
        self._soll = 0

    @property
    def soll(self):
        return self._soll
    @soll.setter
    def soll(self,value):
        if value != self._soll:
            self._soll = value
            self.RaisePropertyChanged('soll')

    @property
    def ist(self):
        return self._ist
    @ist.setter
    def ist(self,value):
        if value != self._ist:
            self._ist = value
            self.RaisePropertyChanged('ist')
    
    def _set_up_soll(self):
        self._get_soll()
    
    def _set_up_ist(self):
        self._get_ist()
    
    def _get_soll(self):
        if self.name == '':
            self.soll = ''
            return
        if self.name.find('Summe') != -1:
            summe = 0
            for el in self.Liste0:
                summe += el.soll
            self.soll =  int(round(summe))

        elif self.name.find('Luftbilanz') != -1:
            zuluft = 0
            abluft = 0
            for el in self.Liste0:
                if el.name.find('Zuluft') != -1:
                    zuluft = el.soll
                elif el.name.find('Abluft') != -1:
                    abluft = el.soll
            self.soll = int(round(zuluft-abluft))
        
        elif self.name.find('Auswertung') != -1:
            summe = 0
            Bilanz = 0
            Druck = 0
            for el in self.Liste0:
                if el.name.find('Luftbilanz') != -1:
                    Bilanz = el.soll
                elif el.name.find('Druckstufe') != -1:
                    Druck = el.soll
            summe = Bilanz - Druck
            if abs(Bilanz) <= 3:self.soll = 'OK'
            else:self.soll = 'Passt nicht'

        elif self.name.find('Tierhaltung') != -1:
            self.soll = ''

        else:
            self.soll = self.Liste0.soll
        
    def _get_ist(self):
        if self.name == '':
            self.ist = ''
            return
        if self.name.find('Summe') != -1:
            summe = 0
            for el in self.Liste0:
                summe += el.ist
            self.ist =  int(round(summe))
        elif self.name.find('Luftbilanz') != -1:
            zuluft = 0
            abluft = 0
            for el in self.Liste0:
                if el.name.find('Zuluft') != -1:
                    zuluft = el.ist
                elif el.name.find('Abluft') != -1:
                    abluft = el.ist
            self.ist = int(round(zuluft-abluft))
        
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
            if abs(Bilanz) <= 3:self.ist = 'OK'
            else:self.ist = 'Passt nicht'
        elif self.name.find('Tierhaltung') != -1:
            self.ist = ''
        else:
            self.ist = self.Liste0.ist

class MEPRaum(object):
    def __init__(self, elem, list_vsr,LISTE_SCHACHT,logger,DICT_MEP_AUSLASS,DICT_MEP_UEBERSTROM,Dict_Ueber_Manuell,DICT_MEP_UN_AUNLASS,_Einbauteile,_VSR):
        self.logger = logger
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
        self.bezugsnummer = self.elem.LookupParameter('TGA_RLT_VolumenstromProNummer').AsValueString()

        try:self.bezugsname = self.berechnung_nach[self.bezugsnummer]
        except:self.bezugsname = 'keine'

        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''

        self.flaeche = round(self.get_element('Fläche'), 2)
        self.volumen = round(self.get_element('Volumen'),2)
        self.hoehe = int(self.get_element('Lichte Höhe'))
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('TGA_RLT_VolumenstromProFaktor')

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
        self.IsSchacht = (self.get_element('TGA_RLT_InstallationsSchacht') != 0)

        self.schachtname = self.elem.LookupParameter('TGA_RLT_InstallationsSchachtName').AsString()
        self.list_auslass = ObservableCollection[Luftauslass]()

        self.ueber_in_m_class = UeberStromAuslass_Manuell(self.get_element('TGA_RLT_RaumÜberströmungMenge'))
        self.ueber_aus_m_class = UeberStromAuslass_Manuell()

        self._vsr_bearbeiten()
        self._auslass_bearbeiten()
        self._auslass_analyse()
    
        # Übersicht 
        self.angezuluft = MEPGrundInfo('Ange.Zuluft','Angegebener Zuluftstrom')
        self.angeabluft = MEPGrundInfo('Ange.Abluft','Angegebener Abluftluftstrom')
        self.ab_24h = MEPGrundInfo('24h Ab','IGF_RLT_AbluftSumme24h',self.get_element('IGF_RLT_AbluftSumme24h'),self.Liste_H24)
        self.ab_min = MEPGrundInfo('min.RAB','IGF_RLT_AbluftminRaum-IGF_RLT_AbluftminSummeLabor-IGF_RLT_AbluftSumme24h,self.Liste_RZU','',self.Liste_RAB)
        self.ab_lab_min = MEPGrundInfo('min.LAB','IGF_RLT_AbluftminSummeLabor',self.get_element('IGF_RLT_AbluftminSummeLabor'),self.Liste_LAB)
        self.zu_min = MEPGrundInfo('min.RZU','IGF_RLT_ZuluftminRaum','',self.Liste_RZU)
        self.ab_max = MEPGrundInfo('max.RAB','IGF_RLT_AbluftmaxRaum-IGF_RLT_AbluftmaxSummeLabor-IGF_RLT_AbluftSumme24h',0,self.Liste_RAB)
        self.ab_lab_max = MEPGrundInfo('max.LAB','IGF_RLT_AbluftmaxSummeLabor',self.get_element('IGF_RLT_AbluftmaxSummeLabor'),self.Liste_LAB)
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
        self.ueber_in = MEPGrundInfo('Über. Ein.','IGF_RLT_ÜberströmungSummeIn',self.get_element('IGF_RLT_ÜberströmungSummeIn'),self.Liste_UIN)
        self.ueber_aus = MEPGrundInfo('Über. Aus.','IGF_RLT_ÜberströmungSummeAus',self.get_element('IGF_RLT_ÜberströmungSummeAus'),self.Liste_UAU)
        self.ueber_in_manuell = MEPGrundInfo('Über.Ein.M.','TGA_RLT_RaumÜberströmungMenge',self.get_element('TGA_RLT_RaumÜberströmungMenge'))
        self.ueber_aus_manuell = MEPGrundInfo('Über.Aus.M.','TGA_RLT_RaumÜberströmungMenge',self.get_element('TGA_RLT_RaumÜberströmungMenge'))
        self.Druckstufe = MEPGrundInfo('Druckstufe','IGF_RLT_RaumDruckstufeEingabe',self.get_element('IGF_RLT_RaumDruckstufeEingabe'))

        self.ab_minsum = MEPGrundInfo('min.AB_SUM','RAB,LAB,24h,Über,Über_M','',[self.ab_min,self.ab_lab_min,self.ab_24h,self.ueber_aus,self.ueber_aus_manuell])
        self.zu_minsum = MEPGrundInfo('min.ZU_SUM','RZU,Über,Über_M','',[self.zu_min,self.ueber_in,self.ueber_in_manuell])
        self.ab_maxsum = MEPGrundInfo('max.AB_SUM','RAB,LAB,24h,Über,Über_M','',[self.ab_max,self.ab_lab_max,self.ab_24h,self.ueber_aus,self.ueber_aus_manuell])
        self.zu_maxsum = MEPGrundInfo('max.ZU_SUM','RZU,Über,Über_M','',[self.zu_max,self.ueber_in,self.ueber_in_manuell])
        self.nb_zu = MEPGrundInfo('NB Zu','IGF_RLT_ZuluftNachtRaum',0,self.Liste_ZU)
        self.nb_ab = MEPGrundInfo('NB Ab','IGF_RLT_AbluftNachtRaum',0,self.Liste_AB)
        self.tnb_zu = MEPGrundInfo('TNB Zu','IGF_RLT_ZuluftTieferNachtRaum',0,self.Liste_ZU)
        self.tnb_ab = MEPGrundInfo('TNB Ab','IGF_RLT_AbluftTieferNachtRaum',0,self.Liste_AB)
        self.ueber_sum = MEPGrundInfo('Überstrom','IGF_RLT_ÜberströmungRaum',self.get_element('IGF_RLT_ÜberströmungRaum'),[self.ueber_in,self.ueber_aus,self.ueber_aus_manuell,self.ueber_in_manuell])

        self.rzu_Schacht = MEPSchachtInfo('RZU',self.get_element('TGA_RLT_SchachtZuluft'),self.zu_max,self.liste_schacht)
        self.rab_Schacht = MEPSchachtInfo('RAB',self.get_element('TGA_RLT_SchachtAbluft'),self.ab_max,self.liste_schacht)
        self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.tier_zu_max,self.liste_schacht)
        self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.tier_ab_max,self.liste_schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('TGA_RLT_Schacht24hAbluft'),self.ab_24h,self.liste_schacht)
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
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Druckstufe_Berechnen(self):
        n = abs(int(self.Druckstufe.soll/10)) if abs(int(self.Druckstufe.soll/10)) < 6 else 5
        if self.Druckstufe.soll > 0:self.IGF_Legende = n*'+'
        else:self.IGF_Legende = n*'-'
          
    def Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self,zuluftmin):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll
        return self.luft_round(zuluftmin)

    def Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self,zuluftmin):
        zuluftmin =  zuluftmin - self.ueber_in_manuell.soll - self.ueber_in.soll
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll > zuluftmin:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        if zuluftmin < 0:
            zuluftmin = 0

        return self.luft_round(zuluftmin)

    def Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self,zuluftmax):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax:
            zuluftmax = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll
        return zuluftmax

    def Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self,zuluftmax):
        zuluftmax =  zuluftmax - self.ueber_in_manuell.soll - self.ueber_in.soll
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll:
            zuluftmax = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_max.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
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
            
            self.ab_min.soll = self.ab_max.soll = 0

            if self.ueber_Zu_Tag:
                self.zu_max.soll = self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.zu_min.soll)
            else:
                self.zu_max.soll = self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.zu_min.soll)

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
                self.angezuluft.soll = self.luft_round(self.flaeche * float(self.faktor) * self.vol_faktorreduzierung_berechnen )
            elif self.bezugsname == "nurAB_Luftwechsel":
                self.angezuluft.soll = self.luft_round(self.volumen * float(self.faktor) * self.vol_faktorreduzierung_berechnen )
            elif self.bezugsname == "nurAB_Person":
                self.angezuluft.soll = self.luft_round(self.personen * float(self.faktor) * self.vol_faktorreduzierung_berechnen )
            elif self.bezugsname == "nurABMa":
                self.angezuluft.soll = self.luft_round(float(self.faktor) * self.vol_faktorreduzierung_berechnen)
                        
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
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "Luftwechsel":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "Person":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.personen * float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            elif self.bezugsname == "manuell":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(float(self.faktor) * self.vol_faktorreduzierung_berechnen)
            

            if self.ueber_Zu_Tag:
                self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.zu_max.soll)
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_ZU(self.zu_min.soll)
            else:
                self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.zu_max.soll)
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen_Ueber_Ist_Nicht_ZU(self.zu_min.soll)
            
            self.ab_max.soll = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll #- self.Druckstufe.soll
            self.ab_min.soll = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll

            self.angeabluft.soll = self.angezuluft.soll = self.zu_max.soll

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

        wert_schreiben2("TGA_RLT_VolumenstromProName", self.bezugsname)
        wert_schreiben("TGA_RLT_VolumenstromProEinheit", self.einheit)
        wert_schreiben("TGA_RLT_VolumenstromProNummer", self.bezugsnummer)
        wert_schreiben("TGA_RLT_VolumenstromProFaktor", float(self.faktor))

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

    