# coding: utf8
from IGF_Klasse import DB, ItemTemplateNurElem
from IGF_Funktionen._Parameter import get_value
from System.Collections.ObjectModel import ObservableCollection
import clr

class Ebene(ItemTemplateNurElem):
    def __init__(self, elem):
        ItemTemplateNurElem.__init__(self, elem)
        self.name = self.elem.Name
        self.height = get_value(self.elem.get_Parameter(DB.BuiltInParameter.LEVEL_ELEV))

class EbeneFactory:
    uiapp = __revit__
    @staticmethod
    def GetAllEbenen():
        try:
            coll = DB.FilteredElementCollector(EbeneFactory.uiapp.ActiveUIDocument.Document).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
            liste = [Ebene(el) for el in coll]
            liste.sort(key= lambda x:x.height)
        except:
            liste = []
        return ObservableCollection[Ebene](liste)

    @staticmethod
    def GetAllEbenenUKRD():
        liste = [el for el in EbeneFactory.GetAllEbenen() if el.name.find('UKRD') != -1]
        return ObservableCollection[Ebene](liste)

    @staticmethod
    def GetAllEbenenOKRB():
        liste = [el for el in EbeneFactory.GetAllEbenen() if el.name.find('OKRB') != -1]
        return ObservableCollection[Ebene](liste)
    
    @staticmethod
    def GetAllEbenenOKFF():
        liste = [el for el in EbeneFactory.GetAllEbenen() if el.name.find('OKFF') != -1]
        return ObservableCollection[Ebene](liste)
    
    @staticmethod
    def GetAllEbenenOKRF():
        liste = [el for el in EbeneFactory.GetAllEbenen() if el.name.find('OKRF') != -1]
        return ObservableCollection[Ebene](liste)
    