# coding: utf8
from IGF_Klasse import DB, ItemTemplateMitElem
from IGF_Allgemein import Parameter_Dict
from IGF_Funktionen._Parameter import get_value
from IGF_Klasse._familie import Familie
from IGF_Klasse._familie_instance import FamilienInstance
from pyrevit import script
logger = script.get_logger()

class LuftauslassFamilie(Familie):
    def __init__(self, elem):
        Familie.__init__(self, elem)

class LuftauslassInstanceTemplate(FamilienInstance):
    def __init__(self,elem,art = ''):
        FamilienInstance.__init__(self,elem)
        self.logger = logger
        self._art = art
        self._Luftmengenmin = 0
        self._Luftmengennacht = 0
        self._Luftmengenmax = 0
        self._Luftmengentnacht = 0
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self._size = ''
        self._Anmerkung = ''
        self.lufttyp = ''
        self._familyandtyp = self.get_value(self.elem.get_Parameter(DB.BuiltInParameter.OST_ELEM_FAMILY_AND_TYPE_PARAM))

    
    def get_value(self,param):
        return get_value(param)
    
    def get_instance_parameter(self,paramname):
        if paramname in Parameter_Dict.keys():
            return self.elem.get_Parameter(paramname)
        return None
    
    def get_type_parameter(self,paramname):
        if paramname in Parameter_Dict.keys():
            return self.elem.Symbol.get_Parameter(paramname)
        return None
    
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

    def set_up(self):
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.familyandtyp = self.get_value(self.elem.get_Parameter(DB.BuiltInParameter.OST_ELEM_FAMILY_AND_TYPE_PARAM))
    
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

