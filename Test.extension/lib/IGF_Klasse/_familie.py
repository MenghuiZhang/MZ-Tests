# coding: utf8
from System import Enum
from IGF_Klasse import DB, ItemTemplateNurElem
from IGF_Funktionen._Parameter import get_value
from System.Collections.ObjectModel import ObservableCollection
import clr

class Familie(ItemTemplateNurElem):
    def __init__(self, elem):
        ItemTemplateNurElem.__init__(self, elem)
        self.name = self.elem.Name
        self.familiename = self.name
        self.doc = self.elem.Document
        self._defaulttype = None
    
    @property
    def Category(self):
        return self.elem.FamilyCategory
    
    @property
    def CategoryName(self):
        return self.Category.Name
    
    @property
    def CategoryType(self):
        return self.Category.CategoryType.ToString()
    
    @property
    def BuiltInCategory(self):
        try:return self.Category.BuiltInCategory.ToString()
        except:
            try:return Enum.Parse(DB.BuiltInCategory,self.Category.Id.ToString())
            except:pass


    @property
    def Dict_Types(self):
        return {self.doc.GetElement(symbol).get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString():self.doc.GetElement(symbol) for symbol in self.elem.GetFamilySymbolIds()}
    
    @property
    def defaulttype(self):
        if self._defaulttype:return self._defaulttype
        self._defaulttype = self.Dict_Types.values()[0]
        return self._defaulttype

class FamilieFactory:
    uiapp = __revit__

    @staticmethod
    def GetAllFamilie():
        coll = DB.FilteredElementCollector(FamilieFactory.uiapp.ActiveUIDocument.Document).OfClass(clr.GetClrType(DB.Family)).ToElements()
        _dict = {el.Name:el for el in coll}
        return ObservableCollection[Familie]([Familie(_dict[name]) for name in sorted(_dict.keys())])
    
    @staticmethod
    def GetAllModellFamilien():
        return ObservableCollection[Familie]([el for el in FamilieFactory.GetAllFamilie() if el.CategoryType == 'Model'])

    @staticmethod
    def GetAllLuftauslassFamilien():
        return ObservableCollection[Familie]([el for el in FamilieFactory.GetAllFamilie() if el.BuiltInCategory == 'OST_DuctTerminal'])
