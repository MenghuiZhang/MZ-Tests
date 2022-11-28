from customclass import RFTItem
import clr
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection

class CFI():
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document

    @staticmethod
    def get_AllFamilies():
        temp = ObservableCollection[RFTItem]()
        items = DB.FilteredElementCollector(CFI.doc).\
            OfClass(clr.GetClrType(DB.Family)).ToElements()
        temp_dict = {item.Name:item for item in items}
        for el in sorted(temp_dict.keys()):
            temp.Add(RFTItem(el,temp_dict[el].Id,CFI.doc))
        return temp

    @staticmethod
    def get_alltbType():
        temp = ObservableCollection[RFTItem]()
        items = DB.FilteredElementCollector(CFI.doc).\
            OfClass(clr.GetClrType(DB.Family)).ToElements()
        temp_list = [item for item in items if item.FamilyCategoryId.ToString() == '-2000280']
        temp_ids = []
        for el in temp_list:
            for id in el.GetFamilySymbolIds():
                temp_ids.append(id)
        temp_dict = {CFI.doc.GetElement(id).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString():id\
            for id in temp_ids}
        for el in sorted(temp_dict.keys()):
            temp.Add(RFTItem(el,temp_dict[el],CFI.doc))
        return temp

    @staticmethod
    def get_alllafType():
        temp = ObservableCollection[RFTItem]()
        items = DB.FilteredElementCollector(CFI.doc).\
            OfClass(clr.GetClrType(DB.ElementType)).ToElements()
        temp_dict = {item.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString():item \
            for item in items if item.FamilyName == 'Ansichtsfenster'}
        
        for el in sorted(temp_dict.keys()):
            temp.Add(RFTItem(el,temp_dict[el].Id,CFI.doc))
        return temp

    