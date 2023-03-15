# coding: utf8
from __init__ import ItemTemplateMitElem,DB
from System.Collections.ObjectModel import ObservableCollection

class AllgemeinerAnsicht(ItemTemplateMitElem):
    def __init__(self,elem):
        ItemTemplateMitElem.__init__(self,elem)
        self.typ = self.elem.ViewType

class BauteilListeCell(object):
    def __init__(self,data,background = [255,255,255],textcolor = [0,0,0],units = '',textalign = 'Center',accuracy = 1.0,width = None):
        self.data = data
        self.background = background
        self.textcolor = textcolor
        self.units = units
        self.textalign = textalign
        self.accuracy = accuracy
        self.width = width

class Bauteilliste(ItemTemplateMitElem):
    def __init__(self,elem):
        ItemTemplateMitElem.__init__(self,elem)
    
    @property
    def Parameters(self):
        return self.elem.Definition.GetFieldOrder()
    
    @property
    def ColumnHeadings(self):
        return {self.elem.Definition.GetField(param).ColumnHeading:self.elem.Definition.GetField(param) for param in self.Parameters}
    
    def get_Data_String(self):
        Liste = []
        tableData = self.elem.GetTableData()
        sectionBody = tableData.GetSectionData(DB.SectionType.Body)
        for r in range(sectionBody.NumberOfRows):
            rowliste_temp = [self.elem.GetCellText(DB.SectionType.Body, r, c) \
                for c in range(sectionBody.NumberOfColumns)]
            Liste.append(rowliste_temp)
        return Liste

    def get_Data_MitFormat(self):
        Liste = []
        units_default = DB.Units(DB.UnitSystem.Metric)
        ColumnHeadings = self.ColumnHeadings
        tableData = self.elem.GetTableData()
        sectionBody = tableData.GetSectionData(DB.SectionType.Body)
        headers = [self.elem.GetCellText(DB.SectionType.Body, 0, c) \
                for c in range(sectionBody.NumberOfColumns)]
        
        Headerlist = []
        for c in range(sectionBody.NumberOfColumns):
            celltext = self.elem.GetCellText(DB.SectionType.Body, 0, c)
            headerformat = sectionBody.GetTableCellStyle(0,c)
            Background = [headerformat.BackgroundColor.Red,headerformat.BackgroundColor.Green,headerformat.BackgroundColor.Blue]
            Textcolor = [headerformat.TextColor.Red,headerformat.TextColor.Green,headerformat.TextColor.Blue]
            width = ColumnHeadings[headers[c]].SheetColumnWidth *304.8
            Headerlist.append(BauteilListeCell(celltext,Background,Textcolor,width=width))
        Liste.append(Headerlist)
        for r in range(1,sectionBody.NumberOfRows):
            rowlist = []
            for c in range(sectionBody.NumberOfColumns):
                celltext = self.elem.GetCellText(DB.SectionType.Body, r, c)
                cellformat = sectionBody.GetTableCellStyle(r,c)
                Background = [cellformat.BackgroundColor.Red,cellformat.BackgroundColor.Green,cellformat.BackgroundColor.Blue]
                Textcolor = [cellformat.TextColor.Red,cellformat.TextColor.Green,cellformat.TextColor.Blue]
                
                textalignment = ColumnHeadings[headers[c]].HorizontalAlignment.ToString()
                width = ColumnHeadings[headers[c]].SheetColumnWidth *304.8
                Headerformat = ColumnHeadings[headers[c]].GetFormatOptions()
                if Headerformat.UseDefault:
                    param = self.doc.GetElement(ColumnHeadings[headers[c]].ParameterId)
                    if param:
                        try:Headerformat = units_default.GetFormatOptions(param.GetDefinition().UnitType)
                        except:
                            try:Headerformat = units_default.GetFormatOptions(param.GetDefinition().GetSpecTypeId())
                            except:Headerformat = None
                
                try:
                    Accuracy = Headerformat.Accuracy
                    if Accuracy < 0.000001:
                        Accuracy = ''
                except:
                    Accuracy = ''
                try:
                    unit = DB.LabelUtils.GetLabelFor(Headerformat.UnitSymbol)
                except:
                    unit = ''
                if ColumnHeadings[headers[c]].CanDisplayMinMax() == True:
                    if unit:
                        try:
                            celltext = float(celltext.replace(unit,''))
                        except:
                            try:
                                celltext = float(celltext)
                            except:celltext = celltext
                    else:
                        try:
                            celltext = float(celltext)
                        except:celltext = celltext
                cellformat.Dispose()
                

                rowlist.append(BauteilListeCell(celltext,Background,Textcolor,unit,textalignment,Accuracy,width))
            Liste.append(rowlist)
        return Liste

class Ansicht(object):
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    ansichts = ObservableCollection[AllgemeinerAnsicht]([AllgemeinerAnsicht(el) for el in \
        DB.FilteredElementCollector(__revit__.ActiveUIDocument.Document).OfCategory(DB.BuiltInCategory.OST_Views).ToElements()])
    

    @staticmethod
    def get_ansichtstemplates():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is True
        ])

    @staticmethod
    def get_ansichtsisnottemplates():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False
        ])

    @staticmethod
    def get_legende():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Legend
        ])
    
    @staticmethod
    def get_FloorPlan():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.FloorPlan
        ])

    @staticmethod
    def get_CeilingPlan():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.CeilingPlan
        ])

    @staticmethod
    def get_Section():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.Section
        ])
    
    @staticmethod
    def get_Schedule():
        temp = {item.Name:item for item in DB.FilteredElementCollector(Ansicht.doc).\
                OfCategory(DB.BuiltInCategory.OST_Schedules).ToElements()}
        return ObservableCollection[Bauteilliste]([
            Bauteilliste(temp[name]) for name in sorted(temp.keys()) \
                if temp[name].IsTemplate is False and \
                temp[name].OwnerViewId.IntegerValue == -1
        ])
    
    @staticmethod
    def get_ThreeD():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.ansichts if item.elem.IsTemplate is False and item.typ == DB.ViewType.ThreeD
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






                
# class RPLItem(RVItem):
#     def __init__(self,name,elemid,doc):
#         RVItem.__init__(self,name,elemid,doc)