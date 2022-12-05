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
    def __init__(self,data,background = [255,255,255],textcolor = [0,0,0],units = '',textalign = 'Center',accuracy = 1.0):
        self.data = data
        self.background = background
        self.textcolor = textcolor
        self.units = units
        self.textalign = textalign
        self.accuracy = accuracy

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
            Headerlist.append(Data(celltext,Background,Textcolor))
        Liste.append(Headerlist)
        for r in range(1,sectionBody.NumberOfRows):
            rowlist = []
            for c in range(sectionBody.NumberOfColumns):
                celltext = self.elem.GetCellText(DB.SectionType.Body, r, c)
                cellformat = sectionBody.GetTableCellStyle(r,c)
                Background = [cellformat.BackgroundColor.Red,cellformat.BackgroundColor.Green,cellformat.BackgroundColor.Blue]
                Textcolor = [cellformat.TextColor.Red,cellformat.TextColor.Green,cellformat.TextColor.Blue]
                
                textalignment = dict_params[headers[c]].HorizontalAlignment.ToString()
                Headerformat = dict_params[headers[c]].GetFormatOptions()
                try:
                    Accuracy = Headerformat.Accuracy
                except:
                    Accuracy = ''
                try:
                    unit = DB.LabelUtils.GetLabelFor(Headerformat.UnitSymbol)
                except:
                    unit = ''
                if dict_params[headers[c]].CanDisplayMinMax() == True:
                    
                    if unit:
                        if unit == '%':
                            try:
                                celltext = float(celltext.replace(unit,''))/100
                            except:
                                try:
                                    celltext = float(celltext)/100
                                except:celltext = celltext
                        try:
                            celltext = float(celltext.replace(' '+unit,''))
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
                

                rowlist.append(Data(celltext,Background,Textcolor,unit,textalignment,Accuracy))
            Liste.append(rowlist)
        return Liste
                
class RPLItem(RVItem):
    def __init__(self,name,elemid,doc):
        RVItem.__init__(self,name,elemid,doc)
        self.snummer = self.elem.SheetNumber
        self.sname = self.elem.Name
        self.plankopf = ''
        self.plankopfname = ''
        self.vp_hauptviews = []
        self.vp_hauptview = ''
        self.vp_hauptlegends = []
        self.vp_hauptlegend = ''
        self.vp_schnitts = []
        self.vp_schnitt = ''

     
    def get_plankopf(self):
        p = self.elem.GetDependentElements(DB.ElementCategoryFilter\
            (DB.BuiltInCategory.OST_TitleBlocks))
        temp = self.doc.GetElement(p[0]) if p.Count > 0 else ''
        return temp
    
    def get_plankopfname(self):
        p = self.get_plankopf()
        if p == '':return p
        return p.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    
    def get_ansichts(self):
        vpts = self.elem.GetAllViewports()
        self.vp_hauptviews.clear()
        self.vp_hauptlegends.clear()
        self.vp_schnitts.clear()
        for vpid in vpts:
            vp = self.doc.GetElement(vpid)
            vptyp = vp.get_Parameter(DB.BuiltInParameter.VIEW_FAMILY).AsString()
            if vptyp == 'Grundrisse':self.vp_hauptviews.append(vp)
            elif vptyp == 'Legenden':self.vp_hauptlegends.append(vp)
            elif vptyp == 'Schnitte':self.vp_schnitts.append(vp)
        self.vp_hauptview = self.vp_hauptviews[0] if len(self.vp_hauptviews) == 1 else ''
        self.vp_hauptlegend = self.vp_hauptlegends[0] if len(self.vp_hauptlegends) == 1 else ''
        self.vp_schnitt = self.vp_schnitts[0] if len(self.vp_schnitts) == 1 else ''