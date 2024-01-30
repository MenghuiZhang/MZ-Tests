# coding: utf8
from IGF_Klasse import DB, ItemTemplateNurElem
from System import Enum
from IGF_Funktionen._Parameter import get_value
from System.Collections.ObjectModel import ObservableCollection
import clr

class FamilieTyp(ItemTemplateNurElem):
    def __init__(self, elem):
        ItemTemplateNurElem.__init__(self, elem)
        self.name = self.elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        self.doc = self.elem.Document
    
    @property
    def familyname(self):
        try:return self.elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString()
        except:return ''
    
    @property
    def familytypname(self):
        if self.familyname:
            return self.familyname + ': ' + self.name
        else:
            return self.name

    
    @property
    def Category(self):
        return self.elem.Category
    
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
            try:return Enum.Parse(DB.BuiltInCategory,self.Category.Id.ToString()).ToString()
            except:pass


class FamilieTypFactory:
    uiapp = __revit__

    @staticmethod
    def GetAllFamilieType():
        try:
            coll = DB.FilteredElementCollector(FamilieTypFactory.uiapp.ActiveUIDocument.Document).OfClass(DB.FamilySymbol).ToElements()
            liste = [FamilieTyp(el) for el in coll]
            liste.sort(key= lambda x:x.familytypname)
            return ObservableCollection[FamilieTyp](liste)
        except:
            return ObservableCollection[FamilieTyp]()

    @staticmethod
    def GetAllRVTLink():
        coll = DB.FilteredElementCollector(FamilieTypFactory.uiapp.ActiveUIDocument.Document).OfClass(clr.GetClrType(DB.RevitLinkInstance)).ToElements()
        liste = [FamilieTyp(el) for el in coll]
        for el in liste:
            el.name = el.elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        liste.sort(key= lambda x:x.name)
        return ObservableCollection[FamilieTyp](liste)
    
    @staticmethod
    def GetAllModellFamilien():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.CategoryType == 'Model'])

    @staticmethod
    def GetAllModellFamilienType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.CategoryType == 'Model'])
    
    @staticmethod
    def GetAllLuftkanalformteileFamilienType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.BuiltInCategory == 'OST_DuctFitting'])
    
    @staticmethod
    def GetAllRohrformteileFamilienType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.BuiltInCategory == 'OST_PipeFitting'])

    @staticmethod
    def GetAllLuftauslassFamilien():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.BuiltInCategory == 'OST_DuctTerminal'])
    
    @staticmethod
    def GetAllLuftkanalBeschriftung():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.BuiltInCategory == 'OST_DuctTags'])
    
    @staticmethod
    def GetAllLuftkanalZubehoerBeschriftung():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllFamilieType() if el.BuiltInCategory == 'OST_DuctAccessoryTags'])
    
    @staticmethod
    def GetAllSystemType():
        coll = DB.FilteredElementCollector(FamilieTypFactory.uiapp.ActiveUIDocument.Document).OfClass(DB.MEPSystemType).ToElements()
        liste = []
        for systemtyp in coll:
            liste.append(FamilieTyp(systemtyp))
        liste.sort(key= lambda x:x.name)
        return ObservableCollection[FamilieTyp](liste)
    
    @staticmethod
    def GetAllLuftkanalSystemType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllSystemType() if el.BuiltInCategory == 'OST_DuctSystem'])
    
    @staticmethod
    def GetAllVerwendetLuftkanalSystemType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllLuftkanalSystemType() if el.elem.GetDependentElements(DB.ElementClassFilter(DB.MEPSystem)).Count > 0])

    @staticmethod
    def GetAllRohrSystemType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllSystemType() if el.BuiltInCategory == 'OST_PipingSystem'])
    
    @staticmethod
    def GetAllVerwendetRohrSystemType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllRohrSystemType() if el.elem.GetDependentElements(DB.ElementClassFilter(DB.MEPSystem)).Count > 0])
    
    @staticmethod
    def GetAllElektroSystemType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllSystemType() if el.BuiltInCategory == 'OST_PipingSystem'])
    
    @staticmethod
    def GetAllVerwendetElektroSystemType():
        return ObservableCollection[FamilieTyp]([el for el in FamilieTypFactory.GetAllRohrSystemType() if el.elem.GetDependentElements(DB.ElementClassFilter(DB.MEPSystem)).Count > 0])
    
    
