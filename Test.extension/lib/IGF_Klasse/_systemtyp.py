from IGF_Klasse import ItemTemplateMitName,ObservableCollection,DB

class Systemtyp(ItemTemplateMitName):
    def __init__(self, name, liste):
        ItemTemplateMitName.__init__(self,name)
        self.liste = liste

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

