import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
from __init__ import RFIItem,RMitem

class RVItem(RMitem):
    def __init__(self,name,elemid,doc):
        RMitem.__init__(self,name,elemid)
        self.doc = doc
        self.elem = doc.GetElement(self.elemid)
        self.typ = self.elem.ViewType

class RBLItem(RVItem):
    def __init__(self,name,elemid,doc):
        RVItem.__init__(self,name,elemid,doc)
    
    def get_Params(self):
        return self.elem.Definition.GetFieldOrder()
    
    def get_ParamsType(self):
        params = self.get_Params()
        dict_params = {self.elem.Definition.GetField(param).ColumnHeading:\
            self.elem.Definition.GetField(param).CanDisplayMinMax() for param in params}
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
        dict_params = self.get_ParamsType()
        tableData = self.elem.GetTableData()
        sectionBody = tableData.GetSectionData(DB.SectionType.Body)
        headers = [self.elem.GetCellText(DB.SectionType.Body, 0, c) \
                for c in range(sectionBody.NumberOfColumns)]
        Liste.append(headers)
        for r in range(1,sectionBody.NumberOfRows):
            rowlist = []
            for c in range(sectionBody.NumberOfColumns):
                celltext = self.elem.GetCellText(DB.SectionType.Body, r, c)
                if dict_params[headers[c]] == True:
                    if celltext.find('%') == -1:
                        try:celltext = float(celltext[:celltext.find(' ')])
                        except:
                            try:celltext = float(celltext)
                            except:pass
                rowlist.append(celltext)
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