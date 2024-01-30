# coding: utf8
from IGF_Klasse import DB,ItemTemplate,TemplateItemBase,TEMPCLASS
from System.Collections.ObjectModel import ObservableCollection
from IGF_Funktionen._Allgemein import Nummer_Korrigieren

class SharedParameter(object):
    def __init__(self, name = '',guid = '',type = '',disziplin = '',userinfo = '', group = '', externaldefinition = ''):
        self.name = name
        self.guid = guid
        self.type = type
        self.disziplin = disziplin
        self.userinfo = userinfo
        self.externaldefinition = externaldefinition
        self.group = group
        self.set_up_shared()
    
    def get_grundinfo(self):
        if self.externaldefinition:
            self.name = self.externaldefinition.Name
            self.userinfo = self.externaldefinition.Description
            self.guid = self.externaldefinition.GUID.ToString()
            self.group = self.externaldefinition.OwnerGroup.Name
    
    def get_type(self):
        if self.externaldefinition:
            try:
                self.type = DB.LabelUtils.GetLabelFor(self.externaldefinition.ParameterType)
            except:
                self.type = DB.LabelUtils.GetLabelForSpec(self.externaldefinition.GetDataType())
    
    def get_discipline(self):
        if self.externaldefinition:
            try:
                diszipline = DB.UnitUtils.GetUnitGroup(self.externaldefinition.UnitType)
                if diszipline.ToString() == 'Common':self.disziplin = 'Allgemein'
                elif diszipline.ToString() == 'Energy':self.disziplin = 'Energie'
                elif diszipline.ToString() == 'Structural':self.disziplin = 'Tragwerk'
                elif diszipline.ToString() == 'HVAC':self.disziplin = 'Lüftung'
                elif diszipline.ToString() == 'Piping':self.disziplin = 'Rohre'
                elif diszipline.ToString() == 'Electrical':self.disziplin = 'Elektro'

            except:
                self.disziplin = DB.LabelUtils.GetLabelForDiscipline(DB.UnitUtils.GetDiscipline(self.externaldefinition.GetDataType()))
    
    def set_up_shared(self):
        self.get_grundinfo()
        self.get_type()
        self.get_discipline()

class ProjektParameter(SharedParameter):
    def __init__(self, bindingmap = None):
        SharedParameter.__init__(self)
        self.bindingmap = bindingmap
        self.typOrex = ''
        self.externaldefinition = ''
        self.sharedparam = ''
        self.binding = ''
        self.internaldefinition = ''
        self.cates = ''
        self.set_up_projekt()
        self.set_up_shared()
    
    def get_binding(self):
        if self.bindingmap:
            self.binding = self.bindingmap.Current
            self.internaldefinition = self.bindingmap.Key
    
    def get_paramtyp(self):
        if self.binding:
            if self.binding.GetType().ToString() == 'Autodesk.Revit.DB.InstanceBinding':
                self.typOrex = 'Exemplar'
            else:
                self.typOrex = 'Type'

    def get_sharedparameter(self):
        if self.internaldefinition:
            self.sharedparam = self.doc.GetElement(self.internaldefinition.Id)
            self.guid = self.sharedparam.GuidValue.ToString()
    
    def get_externaldefinition(self):
        if self.guid:
            file = __revit__.Application.OpenSharedParameterFile()
            for g in file.Groups:
                if g:
                    for d in g.Definitions:
                        if d.GUID.ToString() == self.guid:
                            self.externaldefinition = d
                            return

    def get_paramgroup(self):
        if self.internaldefinition:
            try:
                self.paramgroup = DB.LabelUtils.GetLabelFor(self.internaldefinition.ParameterGroup)
            except:
                self.paramgroup = DB.LabelUtils.GetLabelForGroup(self.internaldefinition.GetGroupTypeId())

    def get_cates(self):
        if self.binding:
            cates = self.binding.Categories
            cateName = ''
            for cate in cates:
                cateName = cateName + cate.Name + ','
            self.cates = cateName[:-1]
    
    def set_up_projekt(self):
        self.get_binding()
        self.get_paramtyp()
        self.get_sharedparameter()
        self.get_externaldefinition()
        self.get_cates()
        self.get_paramgroup()

class LaboranschlussItemBase(ItemTemplate):
    def __init__(self,paramname):
        ItemTemplate.__init__(self)
        self.parametername = paramname
        try:self.tmp_Param = self.parametername.replace('_unter','')
        except:self.tmp_Param = self.parametername
        if self.tmp_Param == self.parametername:
            self.IstSuB = False
        else:
            self.IstSuB = True
        self.typname = self.tmp_Param[27:]
        self.luftmengemin = 0
        self.luftmengemax = 0
        self.druckverlust = 0
        self.durchmesser = None
        self.breite = None
        self.hoehe = None
        self.set_up()
        self.Liste_type = []
        self.typenameohnedetail = self.typname[:self.typname.find('_'+str(self.druckverlust)+'Pa')]
    
    def Parameter_aktualisieren(self):
        self.tmp_Param = self.parametername.replace('_unter','')
        if self.tmp_Param == self.parametername:
            self.IstSuB = False
        else:
            self.IstSuB = True
        self.typname = self.tmp_Param[27:]
        self.luftmengemin = 0
        self.luftmengemax = 0
        self.druckverlust = 0
        self.durchmesser = None
        self.breite = None
        self.hoehe = None
        self.set_up()
        self.Liste_type = []
        self.typenameohnedetail = self.typname[:self.typname.find('_'+str(self.druckverlust)+'Pa')]

    @property
    def typname1(self):
        if self.typname.find('Pa') != -1:
            stelle = self.typname.find('Pa')
            while (self.typname[stelle] != '_'):
                stelle-=1
                if stelle == 0:
                    return
            return self.typname[:stelle]

    @property
    def art(self):
        if self.parametername.find('IGF_RLT_Laboranschluss_LAB') != -1:
            return 'LAB'
        elif self.parametername.find('IGF_RLT_Laboranschluss_24h') != -1:
            return '24h'
        elif self.parametername.find('IGF_RLT_Laboranschluss_LZU') != -1:
            return 'LZU'
        elif self.parametername.find('IGF_RLT_Laboranschluss_UML') != -1:
            return 'UML'
        return ''
    
    def set_up(self):
        self.tmp_Param = self.parametername.replace('_unter','')
        if self.tmp_Param == self.parametername:
            self.IstSuB = False
        else:
            self.IstSuB = True

        self.typname = self.tmp_Param[27:]
        self.luftmengemin = 0
        self.luftmengemax = 0
        self.druckverlust = 0
        self.durchmesser = None
        self.breite = None
        self.hoehe = None

        if self.typname.find('_kon') != -1:
            if self.typname[-1] == '_':
                self.luftmengemax = self.luftmengemin = int(self.typname[self.typname.find('_kon')+4:-1])
            else:
                self.luftmengemax = self.luftmengemin = int(self.typname[self.typname.find('_kon')+4:])
        
        elif self.typname.find('_max') != -1 and self.typname.find('_min') != -1:
            self.luftmengemax = int(self.typname[self.typname.find('_max')+4:])
            self.luftmengemin = int(self.typname[self.typname.find('_min')+4:self.typname.find('_max')])

        if self.typname.find('_DN') != -1:
            if self.typname.find('_kon') != -1:
                self.durchmesser = int(self.typname[self.typname.find('_DN')+3:self.typname.find('_kon')])
            elif self.typname.find('_min') != -1:
                self.durchmesser = int(self.typname[self.typname.find('_DN')+3:self.typname.find('_min')])
        elif self.typname.find('Pa_') != -1:
            if self.typname.find('_kon') != -1:
                text = self.typname[self.typname.find('Pa_')+3:self.typname.find('_kon')]
                if text.find('x') != -1:
                    self.breite = int(text[:text.find('x')])
                    self.hoehe = int(text[text.find('x')+1:])
            elif self.typname.find('_min') != -1:
                text = self.typname[self.typname.find('Pa_')+3:self.typname.find('_min')]
                if text.find('x') != -1:
                    self.breite = int(text[:text.find('x')])
                    self.hoehe = int(text[text.find('x')+1:])
        
        if self.typname.find('Pa') != -1:
            stelle1 = self.typname.find('Pa')
            stelle = self.typname.find('Pa')
            while(self.typname[stelle] != '_'):
                stelle -=1
                if stelle == 0:
                    self.druckverlust = 0
                    break
            try:self.druckverlust = int(self.typname[stelle+1:stelle1])
            except:self.druckverlust = 0
        self.Liste_type = []
        self.typenameohnedetail = self.typname[:self.typname.find('_'+str(self.druckverlust)+'Pa')]
    
class Laboranschlusss(LaboranschlussItemBase):
    """Labaranschluss Parameter"""
    def __init__(self,paramname):
        LaboranschlussItemBase.__init__(self,paramname)
        self._Anzahl = ''
        self._ganzahl = 0
        self.Paramwert = ''

    @property
    def Anzahl(self):
        return self._Anzahl
    @Anzahl.setter
    def Anzahl(self,value):
        if value != self._Anzahl:
            self._Anzahl = value
            self.RaisePropertyChanged('Anzahl')

    @property
    def ganzahl(self):
        return self._ganzahl
    @ganzahl.setter
    def ganzahl(self,value):
        if value != self._ganzahl:
            self._ganzahl = value
            self.RaisePropertyChanged('ganzahl')

    def get_ganzahl(self):
        summe = 0
        try:
            if self.Anzahl.find(',') == -1:
                Liste = [self.Anzahl]
            else:
                Liste = self.Anzahl.split(',')
            for i in Liste:
                if not i:
                    continue
                while(i[0] == ' '):
                    i = i[1:]
                    if not i:break
                if not i:
                    continue
                anzahl = int(i[:i.find('xZeile')])
                summe += anzahl
            self.ganzahl = summe
        except:pass

    def get_Paramwert(self):
        if self.ganzahl == 0:
            self.Paramwert = ''
            return
        try:
            self.Paramwert = str(self.ganzahl) + '_(' + self.Anzahl + ')'
        except:
            pass

class Laboranschluss_NEU(TemplateItemBase):
    def __init__(self):
        TemplateItemBase.__init__(self)
        self.Parameter = Laboranschlusss('IGF_RLT_Laboranschluss_000_')
        self._Liste_Typename = []
        self._typname = self.Parameter.typname
        self._TypenIndex = 0
        self._Anzahl_Labor = ''
        self._Anzahl = ''
        self._Anzahl_MEP = ''
        self._ganzahl = 0
        self._Labor = True
        self._MEP = False
        self.nothasLabor = True
        self.Parameter_Art = 'IGF_RLT_Laboranschluss_01_Art'
        self.Parameter_Anzahl = 'IGF_RLT_Laboranschluss_01_Anzahl'
        self.Liste_vorhanden = []
        self.Paramwert = ''
        self.LABMIN = 0
        self.LABMAX = 0
        self.LZUMIN = 0
        self.LZUMAX = 0
        self.LAB24H = 0
        self.Liste_LaborMedien_Heinekamp = ObservableCollection[TEMPCLASS]()
    
    def Reset(self):
        self.Parameter.parametername ='IGF_RLT_Laboranschluss_000_'
        self.Parameter.Parameter_aktualisieren()
        self._Liste_Typename = []
        self._typname = self.Parameter.typname
        self.Anzahl_Labor = ''
        self.Anzahl_MEP = ''
        self.ganzahl = 0
        self.Labor = True
        self.nothasLabor = True
        self.Liste_vorhanden = []
        self.RaisePropertyChanged_art_Typ()
        self.RaisePropertyChanged_Anzahl()
    
    def wert_schreiben(self,mepraum):
        param1 = mepraum.LookupParameter(self.Parameter_Art)
        param2 = mepraum.LookupParameter(self.Parameter_Anzahl)
        if param1:
            if self.Parameter.parametername != 'IGF_RLT_Laboranschluss_000_':
                param1.Set(self.Parameter.parametername)
                if param2:
                    param2.Set(self.Paramwert)
            else:
                param1.Set('')
                if param2:
                    param2.Set('')
    
    def Parameternamechanged(self):
        self._typname = self.Parameter.typname
        self.RaisePropertyChanged('art')
    
    def get_parameter_MEP(self,mepraum):
        param1 = mepraum.LookupParameter(self.Parameter_Art)
        param2 = mepraum.LookupParameter(self.Parameter_Anzahl)
        if param1:
            wert = param1.AsString()
            if wert in ['',None]:self.Parameter.parametername = 'IGF_RLT_Laboranschluss_000_'
            else: self.Parameter.parametername = wert

        if param2:
            value = param2.AsString()
            if value:
                try:
                    detail = value[value.find('[')+1:value.find(']')]
                except:
                    detail = ''
                self.Anzahl_MEP = detail
            else:
                self.Anzahl_MEP = ''

        
        self.Parameter.Parameter_aktualisieren()
        self._typname = self.Parameter.typname
        if not self.typname:
            self.Anzahl_MEP = ''
        self.RaisePropertyChanged_art_Typ()
        self.RaisePropertyChanged_Anzahl()
    
    def TypenIndex_Aktualisieren(self):
        if self.typname == '':
            self.TypenIndex = 0
            self.RaisePropertyChanged_art_Typ()
            self.Liste_LaborMedien_Heinekamp.Clear()
            return
        for param in self.Liste_Typename:
            if param.typname == self.typname:
                self.TypenIndex = self.Liste_Typename.IndexOf(param)
                self.RaisePropertyChanged_art_Typ()
                return

    @property
    def ganzahl(self):
        return self._ganzahl
    @ganzahl.setter
    def ganzahl(self,value):
        if value != self._ganzahl:
            self._ganzahl = value
            self.RaisePropertyChanged('ganzahl')
    
    @property
    def MEP(self):
        return self._MEP
    @MEP.setter
    def MEP(self,value):
        if value != self._MEP:
            self._MEP = value
            self.RaisePropertyChanged('MEP')
    
    @property
    def Labor(self):
        return self._Labor
    @Labor.setter
    def Labor(self,value):
        if value != self._Labor:
            self._Labor = value
            self.RaisePropertyChanged('Labor')

    @property
    def Liste_Typename(self):
        return self._Liste_Typename
    @Liste_Typename.setter
    def Liste_Typename(self,value):
        if self._Liste_Typename != value:
            self._Liste_Typename = value
            self.RaisePropertyChanged('Liste_Typename')
    @property
    def art(self):
        return self.Parameter.art
    
    @property
    def TypenIndex(self):
        return self._TypenIndex
    
    @TypenIndex.setter
    def TypenIndex(self,value):
        if self._TypenIndex != value:
            self._TypenIndex = value
            self.RaisePropertyChanged('TypenIndex')
    
    @property
    def typname(self):
        return self.Parameter.typname
    
    @property
    def Anzahl_MEP(self):
        return self._Anzahl_MEP
    @Anzahl_MEP.setter
    def Anzahl_MEP(self,value):
        if value != self._Anzahl_MEP:
            self._Anzahl_MEP = value
            self.RaisePropertyChanged('Anzahl_MEP')
    
    @property
    def Anzahl_Labor(self):
        return self._Anzahl_Labor
    @Anzahl_Labor.setter
    def Anzahl_Labor(self,value):
        if value != self._Anzahl_Labor:
            self._Anzahl_Labor = value
            self.RaisePropertyChanged('Anzahl_Labor')
    
    @property
    def Anzahl(self):
        if self.Labor:
            return self.Anzahl_Labor
        else:
            return self.Anzahl_MEP
    
    def RaisePropertyChanged_Anzahl(self):
        self.RaisePropertyChanged('Anzahl')
        self.Parameter.Anzahl = self.Anzahl
        self.get_ganzahl()
        try:self.get_Luftmenge()
        except:pass
    
    def RaisePropertyChanged_art_Typ(self):
        self.RaisePropertyChanged('art')
        self.RaisePropertyChanged('typname')
    
    def RaisePropertyChanged_Liste(self):
        self.RaisePropertyChanged('art')
        self.RaisePropertyChanged('typname')
        self.RaisePropertyChanged('Anzahl')

    def get_ganzahl(self):
        summe = 0
        try:
            manuell = self.Anzahl[:self.Anzahl.find('(')]
            connect = self.Anzahl[self.Anzahl.find('('):]
            if manuell.find(',') == -1:
                Liste0 = [manuell]
            else:
                Liste0 = manuell.split(',')
            if connect.find(',') == -1:
                Liste1 = [connect]
            else:
                Liste1 = connect.split(',')
            for el in Liste0:
                if not el:continue
                while(el[0] == ' '):
                    el = el[1:]
                    if not el:break
                anzahl = int(el[:el.find('xZeile')])
                summe += anzahl
            for el in Liste1:
                if el.find('Zeile') != -1:
                    summe +=1
            self.ganzahl = summe
        except:
            pass

    def get_Paramwert(self):
        if self.ganzahl == 0:
            self.Paramwert = ''
            return
        try:
            self.Paramwert = str(self.ganzahl) + '_[' + self.Anzahl + ']'
        except:
            pass
    
    def get_Luftmenge(self):
        self.LABMIN = 0
        self.LABMAX = 0
        self.LZUMIN = 0
        self.LZUMAX = 0
        self.LAB24H = 0

        if self.ganzahl == 0:return
        if self.art == '24h':
            self.LAB24H += self.Parameter.luftmengemin * self.ganzahl
        elif self.art == 'UML':
            return
        elif self.art == 'LAB':
            self.LABMIN += self.Parameter.luftmengemin * self.ganzahl
            liste = self.Anzahl.split(',')
            anzahl = 0
            for el in liste:
                if el.find('-r') != -1:
                    if el.find('xZeile') != -1:
                        tmp = el[:el.find('xZeile')]
                        anzahl += int(tmp)
                    elif el.find('Zeile') != -1:
                        anzahl += 1
            self.LABMAX += self.Parameter.luftmengemin * self.ganzahl + anzahl * (self.Parameter.luftmengemax - self.Parameter.luftmengemin)
        
        elif self.art == 'LZU':
            self.LZUMIN += self.Parameter.luftmengemin * self.ganzahl
            liste = self.Anzahl.split(',')
            anzahl = 0
            for el in liste:
                if el.find('-r') != -1:
                    if el.find('xZeile') != -1:
                        tmp = el[:el.find('xZeile')]
                        anzahl += int(tmp)
                    elif el.find('Zeile') != -1:
                        anzahl += 1
            self.LZUMAX += self.Parameter.luftmengemin * self.ganzahl + anzahl * (self.Parameter.luftmengemax - self.Parameter.luftmengemin)            

class LaborMedien_Parameter_Factory:
    uiapp = __revit__

    @staticmethod
    def GetAllMEPParameters():
        try:
            return DB.FilteredElementCollector(LaborMedien_Parameter_Factory.uiapp.ActiveUIDocument.Document).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()[0].Parameters
        except:return []
    
    @staticmethod
    def GetAllParameterElement():
        try:
            return DB.FilteredElementCollector(LaborMedien_Parameter_Factory.uiapp.ActiveUIDocument.Document).OfClass(DB.ParameterElement).WhereElementIsNotElementType().ToElements()
        except:return []
    
    @staticmethod
    def GetAllLaborMedienParameter():
        Liste = [el.Definition.Name for el in LaborMedien_Parameter_Factory.GetAllMEPParameters() if el.Definition.Name.find('IGF_RLT_Laboranschluss') != -1]
        Liste.sort()
        return ObservableCollection[LaboranschlussItemBase]([LaboranschlussItemBase(param) for param in Liste])
    
    @staticmethod
    def GetAllLaborMedienParameter_Liste(Liste):
        return ObservableCollection[Laboranschlusss]([Laboranschlusss(param) for param in Liste])
    
    @staticmethod
    def GetAllLaborMedienParameter_Neu_Vorlage():
        Liste = [el.Definition.Name for el in LaborMedien_Parameter_Factory.GetAllMEPParameters() if (el.Definition.Name.find('IGF_RLT_Laboranschluss_0') != -1 or el.Definition.Name.find('IGF_RLT_Laboranschluss_1') != -1) and (el.Definition.Name.find('Art') != -1)] 
        neu = ObservableCollection[Laboranschluss_NEU]()
        for n in range(len(Liste)):
            temp = Laboranschluss_NEU()
            temp.Parameter_Art = 'IGF_RLT_Laboranschluss_{}_Art'.format(Nummer_Korrigieren(n+1,2))
            temp.Parameter_Anzahl = 'IGF_RLT_Laboranschluss_{}_Anzahl'.format(Nummer_Korrigieren(n+1,2))
            neu.Add(temp)
        return neu
            
    @staticmethod
    def GetAllLABParameter():
        return ObservableCollection[LaboranschlussItemBase]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter() if el.art == 'LAB'])
    
    @staticmethod
    def GetAllLZUParameter():
        return ObservableCollection[LaboranschlussItemBase]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter() if el.art == 'LZU'])
    
    @staticmethod
    def GetAll24hParameter():
        return ObservableCollection[LaboranschlussItemBase]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter() if el.art == '24h'])
    
    @staticmethod
    def GetAllUMLParameter():
        return ObservableCollection[LaboranschlussItemBase]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter() if el.art == 'UML'])
    
    @staticmethod
    def GetAllLaborMedienParameter_Wert():
        Liste = [el.Definition.Name for el in LaborMedien_Parameter_Factory.GetAllMEPParameters() if el.Definition.Name.find('IGF_RLT_Laboranschluss') != -1]
        Liste.sort()
        liste1 = [Laboranschlusss(param) for param in Liste]
        return ObservableCollection[Laboranschlusss](liste1)
    
    @staticmethod
    def GetAllLABParameter_Wert():
        return ObservableCollection[Laboranschlusss]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter_Wert() if el.art == 'LAB'])
    
    @staticmethod
    def GetAllLZUParameter_Wert():
        return ObservableCollection[Laboranschlusss]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter_Wert() if el.art == 'LZU'])
    
    @staticmethod
    def GetAll24hParameter_Wert():
        return ObservableCollection[Laboranschlusss]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter_Wert() if el.art == '24h'])
    
    @staticmethod
    def GetAllUMLParameter_Wert():
        return ObservableCollection[Laboranschlusss]([el for el in LaborMedien_Parameter_Factory.GetAllLaborMedienParameter_Wert() if el.art == 'UML'])

