from customclass._ansicht import RVItem,RBLItem
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection

def get_ansichts(doc):
    temp = ObservableCollection[RVItem]()
    items = DB.FilteredElementCollector(doc).\
        OfCategory(DB.BuiltInCategory.OST_Views).ToElements()
    temp_dict = {item.Name:item for item in items}
    for el in sorted(temp_dict.keys()):
        temp.Add(RVItem(el,temp_dict[el].Id,doc))
    return temp

class Ansicht():
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    ansichts = get_ansichts(doc)

    @staticmethod
    def get_ansichts():
        temp = ObservableCollection[RVItem]()
        items = DB.FilteredElementCollector(Ansicht.doc).\
            OfCategory(DB.BuiltInCategory.OST_Views).ToElements()
        temp_dict = {item.Name:item for item in items}
        for el in sorted(temp_dict.keys()):
            temp.Add(RVItem(el,temp_dict[el].Id,Ansicht.doc))
        return temp

    @staticmethod
    def get_ansichtstemplates():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is True
        ])

    @staticmethod
    def get_ansichtsisnottemplates():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False
        ])

    @staticmethod
    def get_legende():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Legend
        ])
    
    @staticmethod
    def get_FloorPlan():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.FloorPlan
        ])

    @staticmethod
    def get_CeilingPlan():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.CeilingPlan
        ])

    @staticmethod
    def get_Section():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Section
        ])
    
    @staticmethod
    def get_Schedule():
        temp = {item.Name:item for item in DB.FilteredElementCollector(Ansicht.doc).\
                OfCategory(DB.BuiltInCategory.OST_Schedules).ToElements()}
        return ObservableCollection[RBLItem]([
            RBLItem(name,temp[name].Id,Ansicht.doc) \
                for name in sorted(temp.keys()) \
                if temp[name].IsTemplate is False and \
                temp[name].OwnerViewId.ToString() == '-1'
        ])
    
    @staticmethod
    def get_Section():
        return ObservableCollection[RVItem]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Section
        ])
    
    # @staticmethod
    # def get_Section():
    #     return ObservableCollection[RVItem]([
    #         item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Section
    #     ])


    # @staticmethod
    # def get_CeilingPlan():
    #     return ObservableCollection[RVItem]([
    #         item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Legend
    #     ])

    # @staticmethod
    # def get_Section():
    #     return ObservableCollection[RVItem]([
    #         item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Legend
    #     ])

    # @staticmethod
    # def get_ansichts():
        # temp = ObservableCollection[RVItem]()
        # items = DB.FilteredElementCollector(Ansicht.doc).\
        #     OfCategory(DB.BuiltInCategory.OST_Views).ToElements()
        # temp_dict = {item.Name:item for item in items}
        # for el in sorted(temp_dict.keys()):
        #     temp.Add(RVItem(el,temp_dict[el].Id,Ansicht.doc))
        # return temp





# class RVItem(RMitem):
#     def __init__(self,name,elemid,doc):
#         RMitem.__init__(name,elemid)
#         self.doc = doc
#         self.elem = doc.GetElement(self.elemid)
#         self.typ = self.elem.ViewType

# class RBLItem(RVItem):
#     def __init__(self,name,elemid,doc):
#         RVItem.__init__(name,elemid,doc)
    
#     def get_Params(self):
#         return self.elem.Definition.GetFieldOrder()
    
#     def get_ParamsType(self):
#         params = self.get_Params()
#         dict_params = {self.elem.Definition.GetField(param).ColumnHeading:\
#             self.elem.Definition.GetField(param).CanDisplayMinMax() for param in params}
#         return dict_params
    
#     def get_Data(self):
#         Liste = []
#         tableData = self.elem.GetTableData()
#         sectionBody = tableData.GetSectionData(DB.SectionType.Body)
#         for r in range(sectionBody.NumberOfRows):
#             rowliste_temp = [self.elem.GetCellText(DB.SectionType.Body, r, c) \
#                 for c in range(sectionBody.NumberOfColumns)]
#             Liste.append(rowliste_temp)
#         return Liste

#     def get_Data2(self):
#         Liste = []
#         dict_params = self.get_ParamsType()
#         tableData = self.elem.GetTableData()
#         sectionBody = tableData.GetSectionData(DB.SectionType.Body)
#         headers = [self.elem.GetCellText(DB.SectionType.Body, 0, c) \
#                 for c in range(sectionBody.NumberOfColumns)]
        
#         for r in range(1,sectionBody.NumberOfRows):
#             rowlist = []
#             for c in range(sectionBody.NumberOfColumns):
#                 celltext = self.elem.GetCellText(DB.SectionType.Body, r, c)
#                 if dict_params[headers[c]] == True:
#                     if celltext.find('%') == -1:
#                         try:celltext = float(celltext[:celltext.find(' ')])
#                         except:
#                             try:celltext = float(celltext)
#                             except:pass
#                 rowlist.append(celltext)
#         Liste.append(rowlist)
                
# class RPLItem(RVItem):
#     def __init__(self,name,elemid,doc):
#         RVItem.__init__(name,elemid,doc)