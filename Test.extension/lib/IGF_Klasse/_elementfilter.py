# coding: utf8
from IGF_Klasse import ItemTemplateMitElem,DB,ItemTemplateMitElemUndName,ObservableCollection,ItemTemplateMitName,ItemTemplate
import Autodesk.Revit.UI as UI

class Systemtyp(ItemTemplateMitName):
    def __init__(self, name):
        ItemTemplateMitName.__init__(self,name)
        self._checked = True

class BauteilSet(ItemTemplate):
    def __init__(self,category,familyname,typname,elems=[]):
        ItemTemplate.__init__(self)
        self.category = category
        self.familyname = familyname
        self.typname = typname
        self.elems = elems
        self._anzahl = self.get_anzahl()
    
    def get_anzahl(self):
        return len(self.elems)

    def get_elemids(self):
        return [el.Id for el in self.elems]

    @property
    def anzahl(self):
        return self._anzahl
    @anzahl.setter
    def anzahl(self,value):
        if value != self._anzahl:
            self._anzahl = value
            self.RaisePropertyChanged('anzahl')

class Category_Parameter:
    def __init__(self,param,IsExemplar = True):
        self.param = param
        self.isExemplar = IsExemplar
        self.Id = self.param.Definition.Id
        self.name = self.param.Definition.Name
        self.storygetype = self.param.StorageType.ToString()
    
    @property
    def DisplayUnitType(self):
        if self.storygetype == 'Double':
            try:return self.param.DisplayUnitType
            except:
                try:return self.param.GetUnitTypeId()
                except:return
    
    def createEqualFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createBeginsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateBeginsWithRule(self.Id,value,True))
        return None
    
    def createContainsFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(self.Id,value,True))
        return None
    
    def createEndsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEndsWithRule(self.Id,value,True))
        return None
    
    def createGreaterEqualFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterOrEqualRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterOrEqualRule(self.Id,value,0.1**6))
        except:
            return None
        
    def createGreaterFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateGreaterRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createNoValueFilter(self,value):
        return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateHasNoValueParameterRule(self.Id))
        
    def createHasValueFilter(self,value):
        return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateHasValueParameterRule(self.Id))
        
    def createLessEqualFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessOrEqualRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessOrEqualRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createLessFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateLessRule(self.Id,value,0.1**6))
        except:
            return None
    
    def createNotBeginsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotBeginsWithRule(self.Id,value,True))
        return None
    
    def createNotContainsFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotContainsRule(self.Id,value,True))
        return None
        
    def createNotEndsWithFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotEndsWithRule(self.Id,value,True))
        return None
    
    def createNotEqualsFilter(self,value):
        if self.storygetype != 'Double':
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotEqualsRule(self.Id,value,True))
        try:
            value = float(value)
            value = DB.ConvertToInternalUnits(value,self.storygetype)
            return DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotEqualsRule(self.Id,value,0.1**6))
        except:
            return None
    
class Category(ItemTemplateMitName):
    def __init__(self,category,elems):
        ItemTemplateMitName.__init__(self,category)
        self.elems = elems
        self.dict_Ohne_System = {}
        self.dict_Mit_System = {}
        self.dict_Systems = {}
        self.Systems = ObservableCollection[Systemtyp]()
        self.ItemSet = ObservableCollection[BauteilSet]()
        self.dict_parameters = {}
    
    @property
    def elemids(self):
        return [el.Id for el in self.elems]
    
    def analyse_system(self):
        for elem in self.elems:
            typname = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            familyname = elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
            try:System = elem.LookupParameter('Systemtyp').AsValueString()
            except:System = ''

            if System not in self.dict_Mit_System.keys():
                self.dict_Mit_System[System] = {}
            if familyname not in self.dict_Mit_System[System].keys():
                self.dict_Mit_System[System][familyname] = {}
            if typname not in self.dict_Mit_System[System][familyname].keys():
                self.dict_Mit_System[System][familyname][typname] = []
            self.dict_Mit_System[System][familyname][typname].append(elem)

            if familyname not in self.dict_Ohne_System.keys():
                self.dict_Ohne_System[familyname] = {}
            if typname not in self.dict_Ohne_System[familyname].keys():
                self.dict_Ohne_System[familyname][typname] = []
            self.dict_Ohne_System[familyname][typname].append(elem)

        for system in sorted(self.dict_Mit_System.keys()):
            item = Systemtyp(system)
            self.Systems.Add(item)
            self.dict_Systems[system] = item

        
        for family in sorted(self.dict_Ohne_System.keys()):
            for typ in sorted(self.dict_Ohne_System[family].keys()):
                item = BauteilSet(self.name,family,typ)
                self.ItemSet.Add(item)

    def analyse_Parameter(self):
        families = {}
        for el in self.elems:
            familyid = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsElementId().IntegerValue
            if familyid not in families.keys():
                families[familyid] = el
        
        for elem in families.values():
            try:param_exemplar = elem.Parameters
            except:param_exemplar = []
            try:parameter_typs = elem.Document.GetElement(elem.GetTypeId()).Parameters
            except:parameter_typs = []
            for param in param_exemplar:
                paramid = param.Definition.Id.IntegerValue
                if paramid not in self.dict_parameters.keys():
                    self.dict_parameters[paramid] = Category_Parameter(param)
            for param in parameter_typs:
                paramid = param.Definition.Id.IntegerValue
                if paramid not in self.dict_parameters.keys():
                    self.dict_parameters[paramid] = Category_Parameter(param,False)

class Beschriftung(ItemTemplateMitName):
    def __init__(self,familytyp,elems):
        ItemTemplateMitName.__init__(self,familytyp)
        self.elems = elems
        self.dict_elems = {}
        self.analyse()
    
    def analyse(self):
        for elem in self.elems:
            tagelement = elem.GetTaggedLocalElement()
            familienname = tagelement.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
            typenname = tagelement.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            if familienname not in self.dict_elems.keys():
                self.dict_elems[familienname] = {}
            if typenname not in self.dict_elems[familienname].keys():
                self.dict_elems[familienname][typenname] = []
            self.dict_elems[familienname][typenname].append(elem)
            tagelement.Dispose()
