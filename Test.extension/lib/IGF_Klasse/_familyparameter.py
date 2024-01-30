# coding: utf8
from IGF_Klasse import DB,ItemTemplate,TemplateItemBase
from System.Collections.ObjectModel import ObservableCollection
from IGF_Allgemein._FamilyParameter import Dict_DIN_Parameter,Dict_LIN_Parameter,Dict_MC_Eklimax_Parameter,Dict_MC_Parameter,Dict_IGF_Parameter,Dict_IGFGA_Parameter
from IGF_Allgemein import Guid
from rpw import revit

class FamilyParameter(ItemTemplate):
    def __init__(self,name,isExemplar = True, Group = DB.BuiltInParameterGroup.INVALID,sharedfile = None):
        ItemTemplate.__init__(self)
        self.name = name
        self._isExemplar = isExemplar
        self.group = Group
        self._externaldefi = None
        self.sharedfile = sharedfile

    @property
    def doc(self):
        try:return revit.doc
        except:return None

    @property
    def isExemplar(self):
        return self._isExemplar
    @isExemplar.setter
    def isExemplar(self,value):
        if value != self._isExemplar:
            self._isExemplar = value
            self.RaisePropertyChanged('isExemplar')
    
    @property
    def hasExDefi(self):
        return self.ExternalDefi is not None
    
    @property
    def ParameterGUID(self):
        if self.hasExDefi:
            if self.doc:
                return self.doc.FamilyManager.get_Parameter(self.ExternalDefi.GUID)
        return False
    
    @property
    def ParameterNAME(self):
        if self.hasExDefi:
            if self.doc:
                return self.doc.FamilyManager.get_Parameter(self.name)
        return False
    
    @property
    def foreground(self):
        if not self.hasExDefi:
            return 'Gray'
        if self.ParameterGUID:
            return 'Red'
        if self.ParameterNAME:
            return 'Blue'
        return 'Black'
    
    @property
    def Tooltip(self):
        if not self.hasExDefi:
            return 'Kein SharedParameter gefunden'
        if self.ParameterGUID:
            return 'Parameter vorhanden'
        if self.ParameterNAME:
            return 'Parameter ist vorhanden aber entspricht nicht der SharedParameter'
        return ''
    
    @property
    def isenabled(self):
        if self.foreground == 'Black':
            return True
        return False
    
    @property
    def ExternalDefi(self):
        if not self._externaldefi:
            if self.sharedfile:
                self.get_external_Definition(self.sharedfile)
        return self._externaldefi

    
    def get_external_Definition(self,sharedfile):
        try:
            for group in sharedfile.Groups:
                for defi in group.Definitions:
                    if self.name == defi.Name:
                        self._externaldefi = defi
                        return
        except:pass
    
    def AddParameter(self,doc):
        def SetFormula(doc,param,Liste):
            for elem in Liste:
                try:
                    doc.FamilyManager.SetFormula(param,elem)
                    return
                except:pass
    
                
        if self.ExternalDefi:
            if not doc.FamilyManager.get_Parameter(self.ExternalDefi.GUID):
                param = doc.FamilyManager.AddParameter(self.ExternalDefi,self.group,self.isExemplar)
                if self.name == 'MC Discharge Unit':SetFormula(doc,param,['WFU'])
                elif self.name == 'MC Water Flow Actual Cold':SetFormula(doc,param,['CWFU'])
                elif self.name == 'MC Water Flow Actual Hot':SetFormula(doc,param,['HWFU'])
                elif self.name == 'MC Height Instance':SetFormula(doc,param,['Höhe','Height','H'])
                elif self.name == 'MC Length Instance':SetFormula(doc,param,['Länge','Length','L'])
                elif self.name == 'MC Width Instance':SetFormula(doc,param,['Breite','Width','B','Tiefe','tiefe'])
                elif self.name == 'MC Diameter Instance':SetFormula(doc,param,['MC_R1*2','Durchmesser Luftkanal','Durchmesser','Luftkanaldurchmesser','D'])

class FamilyParameterFactory:
    @staticmethod
    def GetParameter(DICT,_sharedfile = None):
        Liste = list(DICT.Keys)
        try:liste = [FamilyParameter(name,DICT[name][0],DICT[name][1],_sharedfile) for name in Liste]
        except:liste = []
        Liste = []
        return ObservableCollection[FamilyParameter](liste)
    
    @staticmethod
    def GetDINParameter(_sharedfile = None):
        return FamilyParameterFactory.GetParameter(Dict_DIN_Parameter,_sharedfile)
    
    @staticmethod
    def GetLINParameter(_sharedfile = None):
        return FamilyParameterFactory.GetParameter(Dict_LIN_Parameter,_sharedfile)
    
    @staticmethod
    def GetMCEKLIMAXParameter(_sharedfile = None):
        return FamilyParameterFactory.GetParameter(Dict_MC_Eklimax_Parameter,_sharedfile)

    @staticmethod
    def GetMCParameter(_sharedfile = None):
        return FamilyParameterFactory.GetParameter(Dict_MC_Parameter,_sharedfile)
    
    @staticmethod
    def GetIGFParameter(_sharedfile = None):
        return FamilyParameterFactory.GetParameter(Dict_IGF_Parameter,_sharedfile)
    
    @staticmethod
    def GetIGFGAParameter(_sharedfile = None):
        return FamilyParameterFactory.GetParameter(Dict_IGFGA_Parameter,_sharedfile)
    