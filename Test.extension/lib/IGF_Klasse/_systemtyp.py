# coding: utf8
from IGF_Klasse import ItemTemplateMitName,ObservableCollection,DB

class Systemtyp(ItemTemplateMitName):
    def __init__(self, name, liste):
        ItemTemplateMitName.__init__(self,name)
        self.liste = liste
        self.liste_system = []
        self.liste_elememts = []
        self.liste_familyinstance = []
        self.liste_trassen = []

    def get_all_systeminstance(self):
        for el in self.liste:
            self.liste_system.append(SystemInstance(el))
    
    def get_all_elements(self):
        for el in self.liste_system:
            self.liste_elememts.extend(el.ElementList)
    
    def get_all_familyinstance(self):
        for el in self.liste_system:
            self.liste_familyinstance.extend(el.FamilyInstance)
    
    def get_all_trassen(self):
        for el in self.liste_system:
            self.liste_trassen.extend(el.Trassen)

class SystemInstance(object):
    def __init__(self,elem):
        self.elem = elem
    @property
    def ElementSet(self):
        try:
            return self.elem.PipingNetwork
        except:
            return self.elem.DuctNetwork
    
    @property
    def ElementList(self):
        return list(self.ElementSet)
    
    @property
    def FamilyInstance(self):
        return [elem for elem in self.ElementSet if isinstance(elem,DB.FamilyInstance)]
    
    @property
    def Trassen(self):
        return [elem for elem in self.ElementSet if isinstance(elem,DB.MEPCurve)]

class SystemMethode:
    @staticmethod
    def get_all_verwendete_pipesystem():
        Liste_Systemtyp = ObservableCollection[Systemtyp]()
        coll = DB.FilteredElementCollector(__revit__.ActiveUIDocument.Document).OfCategory(DB.BuiltInCategory.\
                                            OST_PipingSystem).WhereElementIsNotElementType().ToElements()
        Dict = {}
        for el in coll:
            typ = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            if typ not in Dict.keys():
                Dict[typ] = []
            Dict[typ].append(el)

        for el in sorted(Dict.keys()):
            Liste_Systemtyp.Add(Systemtyp(el,Dict[el]))
        return Liste_Systemtyp
    
    @staticmethod
    def get_all_verwendete_ductsystem():
        Liste_Systemtyp = ObservableCollection[Systemtyp]()
        coll = DB.FilteredElementCollector(__revit__.ActiveUIDocument.Document).OfCategory(DB.BuiltInCategory.\
                                            OST_DuctSystem).WhereElementIsNotElementType().ToElements()
        Dict = {}
        for el in coll:
            typ = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            if typ not in Dict.keys():
                Dict[typ] = []
            Dict[typ].append(el)

        for el in sorted(Dict.keys()):
            Liste_Systemtyp.Add(Systemtyp(el,Dict[el]))
        return Liste_Systemtyp
