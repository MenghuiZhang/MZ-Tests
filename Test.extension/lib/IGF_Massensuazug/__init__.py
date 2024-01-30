# coding: utf8
import Autodesk.Revit.DB as DB

class ItemTemplate(object):
    def __init__(self,elem):
        self.elem = elem
    
    @property
    def doc(self):
        return self.elem.Document
    
    @property
    def familiename(self):
        return self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
    
    @property
    def typname(self):
        return self.elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
    
    @property
    def familietyp(self):
        return self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    
    @property
    def systemtyp(self):
        try:return self.elem.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
        except:
            try:return self.elem.get_Parameter(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
            except:return None
