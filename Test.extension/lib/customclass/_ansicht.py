import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
from __init__ import RFIItem,RMitem

class RVItem(RMitem):
    def __init__(self,name,elemid,doc):
        RMitem.__init__(self,name,elemid)
        self.doc = doc
        self.elem = doc.GetElement(self.elemid)
        self.typ = self.elem.ViewType

class Data:
    def __init__(self,data,background = [255,255,255],textcolor = [0,0,0],units = '',textalign = 'Center',accuracy = 1.0,width = None):
        self.data = data
        self.background = background
        self.textcolor = textcolor
        self.units = units
        self.textalign = textalign
        self.accuracy = accuracy
        self.width = width

class RBLItem(RVItem):
    def __init__(self,name,elemid,doc):
        RVItem.__init__(self,name,elemid,doc)
        
    
    def get_Params(self):
        return self.elem.Definition.GetFieldOrder()
    
    def get_column_format(self):
        params = self.get_Params()
        dict_params = {self.elem.Definition.GetField(param).ColumnHeading:self.elem.Definition.GetField(param) for param in params}
        return dict_params
    
    def get_Data(self):
        Liste = []
        tableData = self.elem.GetTableData()
        sectionBody = tableData.GetSectionData(DB.SectionType.Body)
        for r in range(sectionBody.NumberOfRows):
            rowliste_temp = [self.elem.GetCellText(DB.SectionType.Body, r, c) \
                for c in range(sectionBody.NumberOfColumns)]
            Liste.append(rowliste_temp)
        return Liste

    def get_Data2(self):
        Liste = []
        units_default = DB.Units(DB.UnitSystem.Metric)
        dict_params = self.get_column_format()
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
            width = dict_params[headers[c]].SheetColumnWidth *304.8
            Headerlist.append(Data(celltext,Background,Textcolor,width=width))
        Liste.append(Headerlist)
        for r in range(1,sectionBody.NumberOfRows):
            rowlist = []
            for c in range(sectionBody.NumberOfColumns):
                celltext = self.elem.GetCellText(DB.SectionType.Body, r, c)
                cellformat = sectionBody.GetTableCellStyle(r,c)
                Background = [cellformat.BackgroundColor.Red,cellformat.BackgroundColor.Green,cellformat.BackgroundColor.Blue]
                Textcolor = [cellformat.TextColor.Red,cellformat.TextColor.Green,cellformat.TextColor.Blue]
                
                textalignment = dict_params[headers[c]].HorizontalAlignment.ToString()
                width = dict_params[headers[c]].SheetColumnWidth *304.8
                Headerformat = dict_params[headers[c]].GetFormatOptions()
                if Headerformat.UseDefault:
                    param = self.doc.GetElement(dict_params[headers[c]].ParameterId)
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
                if dict_params[headers[c]].CanDisplayMinMax() == True:
                    if unit:
                        # if unit == '%':
                        #     try:
                        #         celltext = float(celltext.replace(unit,''))/100
                        #     except:
                        #         try:
                        #             celltext = float(celltext)/100
                        #         except:celltext = celltext
                        try:
                            celltext = float(celltext.replace(unit,''))
                        except:
                            try:
                                celltext = float(celltext)
                            except:celltext = celltext
                    

                        # if celltext.find('%') == -1:
                        #     try:celltext = float(celltext[:celltext.find(' ')])
                        #     except:
                        #         try:celltext = float(celltext)
                        #         except:pass
                    else:
                        try:
                            celltext = float(celltext)
                        except:celltext = celltext
                cellformat.Dispose()
                

                rowlist.append(Data(celltext,Background,Textcolor,unit,textalignment,Accuracy,width))
            Liste.append(rowlist)
        return Liste

# class RBLItem(RVItem):
#     def __init__(self,name,elemid,doc):
#         RVItem.__init__(self,name,elemid,doc)
    
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
#         Liste.append(headers)
#         for r in range(1,sectionBody.NumberOfRows):
#             rowlist = []
#             for c in range(sectionBody.NumberOfColumns):
#                 celltext = self.elem.GetCellText(DB.SectionType.Body, r, c)
#                 if dict_params[headers[c]] == True:
#                     if celltext.find('%') == -1:
#                         try:celltext = float(celltext)
#                         except:
#                             try:celltext = float(celltext[:celltext.find(' ')])
#                             except:pass
#                 rowlist.append(celltext)
#             Liste.append(rowlist)
#         return Liste
                
class RPLItem(RVItem):
    def __init__(self,name,elemid,doc):
        RVItem.__init__(self,name,elemid,doc)