# coding: utf8
from IGF_Klasse import ItemTemplateMitElem,DB,ItemTemplateMitElemUndName,ObservableCollection,List
import Autodesk.Revit.UI as UI
from IGF_Funktionen._Geometrie import get_ClosestPoints


class AllgemeinerAnsicht(ItemTemplateMitElem):
    def __init__(self,elem):
        ItemTemplateMitElem.__init__(self,elem)
        self.typ = self.elem.ViewType
    
    @staticmethod
    def get_Raster_In_Ansicht(View):
        return DB.FilteredElementCollector(View.Document,View.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()

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
        units_default = self.doc.GetUnits()
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
                    
                    try:Headerformat = units_default.GetFormatOptions(ColumnHeadings[headers[c]].UnitType)
                    except:
                        try:Headerformat = units_default.GetFormatOptions(ColumnHeadings[headers[c]].GetSpecTypeId())
                        except:
                            try:
                                Headerformat = units_default.GetFormatOptions(ColumnHeadings[headers[c]].GetDataType())
                            except:

                                Headerformat = None
                
                try:
                    Accuracy = Headerformat.Accuracy
                    if Accuracy < 0.00001:
                        Accuracy = ''
                except:
                    Accuracy = ''
                try:
                    unit = DB.LabelUtils.GetLabelFor(Headerformat.UnitSymbol)
                except:
                    try:unit = DB.LabelUtils.GetLabelForSymbol(Headerformat.GetSymbolTypeId())
                        
                    except:unit = ''
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

    @staticmethod
    def get_AllAnsichts():
        try:return ObservableCollection[AllgemeinerAnsicht]([AllgemeinerAnsicht(el) for el in \
        DB.FilteredElementCollector(Ansicht.uiapp.ActiveUIDocument.Document).OfCategory(DB.BuiltInCategory.OST_Views).ToElements()])
        except:
            return ObservableCollection[AllgemeinerAnsicht]()

    @staticmethod
    def get_ansichtstemplates():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is True
        ])

    @staticmethod
    def get_ansichtsisnottemplates():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is False
        ])

    @staticmethod
    def get_legende():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is False and item.typ == DB.ViewType.Legend
        ])
    
    @staticmethod
    def get_FloorPlan():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is False and item.typ == DB.ViewType.FloorPlan
        ])

    @staticmethod
    def get_CeilingPlan():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is False and item.typ == DB.ViewType.CeilingPlan
        ])

    @staticmethod
    def get_Section():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is False and item.typ == DB.ViewType.Section
        ])
    
    @staticmethod
    def get_Schedule():
        temp = {item.Name:item for item in DB.FilteredElementCollector(Ansicht.uiapp.ActiveUIDocument.Document).\
                OfCategory(DB.BuiltInCategory.OST_Schedules).ToElements()}
        return ObservableCollection[Bauteilliste]([
            Bauteilliste(temp[name]) for name in sorted(temp.keys()) \
                if temp[name].IsTemplate is False and \
                temp[name].OwnerViewId.IntegerValue == -1
        ])
    
    @staticmethod
    def get_ThreeD():
        return ObservableCollection[AllgemeinerAnsicht]([
            item for item in Ansicht.get_AllAnsichts() if item.elem.IsTemplate is False and item.typ == DB.ViewType.ThreeD
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

class AllgemeinerAnsichtMitRegionShapeLinien(AllgemeinerAnsicht):
    def __init__(self,elem):
        AllgemeinerAnsicht.__init__(self,elem)
        self._ShapeMittelPoint = None
        self._ShapeMaxPoint = None
        self._ShapeMinPoint = None
    
    @property
    def RegionShapeLinien(self):
        try:return self.elem.GetCropRegionShapeManager().GetCropShape()[0]
        except:return None
    
    @property
    def ShapeMittelPoint(self):
        if not self._ShapeMittelPoint:
            lines = self.RegionShapeLinien
            if lines:
                p = DB.XYZ(0,0,0)
                n = 0
                for curve in lines:
                    n += 2
                    p += curve.GetEndPoint(0)
                    p += curve.GetEndPoint(1)
                p = p / n if n > 0 else p
                self._ShapeMittelPoint = p
        return self._ShapeMittelPoint
    
    @property
    def ShapeMaxPoint(self):
        if not self._ShapeMaxPoint:
            lines = self.RegionShapeLinien
            if lines:
                max_x,max_y,max_z = -10000000,-10000000,-10000000
                for curve in lines:
                    max_x = max([max_x,curve.GetEndPoint(0).X,curve.GetEndPoint(1).X])
                    max_y = max([max_y,curve.GetEndPoint(0).Y,curve.GetEndPoint(1).Y])
                    max_z = max([max_z,curve.GetEndPoint(0).Z,curve.GetEndPoint(1).Z])
                self._ShapeMaxPoint = DB.XYZ(max_x,max_y,max_z)
        return self._ShapeMaxPoint
    
    @property
    def ShapeMinPoint(self):
        if not self._ShapeMinPoint:
            lines = self.RegionShapeLinien
            if lines:
                min_x,min_y,min_z = -10000000,-10000000,-10000000
                for curve in lines:
                    min_x = min([min_x,curve.GetEndPoint(0).X,curve.GetEndPoint(1).X])
                    min_y = min([min_y,curve.GetEndPoint(0).Y,curve.GetEndPoint(1).Y])
                    min_z = min([min_z,curve.GetEndPoint(0).Z,curve.GetEndPoint(1).Z])
                self._ShapeMinPoint = DB.XYZ(min_x,min_y,min_z)
        return self._ShapeMinPoint
    
    @property
    def Raster(self):
        return AllgemeinerAnsicht.get_Raster_In_Ansicht(self.elem)

    def Raster_anpassen(self):
        rasters = self.Raster
        if rasters.Count == 0:return
        if not self.ShapeMaxPoint:return
        Lines = self.RegionShapeLinien
        Liste_Lines = [curve for curve in Lines]
        Liste_Points = [Liste_Lines[0].GetEndPoint(0),Liste_Lines[0].GetEndPoint(1)]
        for n in range(1,len(Liste_Lines)):
            if self.Parallel_Test(Liste_Lines[0],Liste_Lines[n]):
                Liste_Points.extend([Liste_Lines[n].GetEndPoint(0),Liste_Lines[n].GetEndPoint(1)])
                break
        for raster in rasters:
            try:
                raster.Pinned = False
                Line_raster = raster.GetCurvesInView(DB.DatumExtentType.ViewSpecific, self.elem)[0]
                Liste_ClosestPointsPair = []
                for Line_Shape in Liste_Lines:
                    if not self.Parallel_Test(Line_raster,Line_Shape):
                        try:Liste_ClosestPointsPair.append(get_ClosestPoints(Line_Shape,Line_raster))
                        except Exception as e:
                            pass
                if len(Liste_ClosestPointsPair) == 4:
                    for p0 in Liste_Points:
                        for ClosestPointsPair in Liste_ClosestPointsPair:
                            if p0.IsAlmostEqualTo(ClosestPointsPair.XYZPointOnFirstCurve):
                                Liste_ClosestPointsPair.remove(ClosestPointsPair)
                                ClosestPointsPair.Dispose()
                                break
                # print(Liste_Point)
                if len(Liste_ClosestPointsPair) == 2:
                    P0 = Liste_ClosestPointsPair[0].XYZPointOnSecondCurve
                    P1 = Liste_ClosestPointsPair[1].XYZPointOnSecondCurve
                    Liste_ClosestPointsPair[0].Dispose()
                    Liste_ClosestPointsPair[1].Dispose()
                    vektor = P0 - P1
                    if vektor.Normalize().IsAlmostEqualTo(Line_raster.Direction):
                        raster.SetCurveInView(DB.DatumExtentType.ViewSpecific, self.elem, DB.Line.CreateBound(P0,P1))
                    else:
                        raster.SetCurveInView(DB.DatumExtentType.ViewSpecific, self.elem, DB.Line.CreateBound(P1,P0))
                raster.Pinned = True
            except Exception as e:pass
            
    def Parallel_Test(self,Line0,Line1):
        if Line0.Direction.IsAlmostEqualTo(Line1.Direction) or Line0.Direction.IsAlmostEqualTo(Line1.Direction.Negate()):return True
        return False

class SchnittInGrundriss(ItemTemplateMitElemUndName):
    def __init__(self,elem,name = ''):
        ItemTemplateMitElemUndName.__init__(self,name,elem)
        self.name = self.elem.Name
        self._SchnittAnsicht = None
        self._SchnittAnsichtKlass = None
    
    @property
    def doc(self):
        return self.elem.Document
    
    @property
    def SchnittAnsicht(self):
        if not self._SchnittAnsicht:
            self._SchnittAnsicht = self.doc.GetElement(self.elem.GetDependentElements(DB.ElementClassFilter(DB.ViewSection))[0])
        return self._SchnittAnsicht
    
    @property
    def SchnittAnsichtKlass(self):
        if not self._SchnittAnsichtKlass:
            self._SchnittAnsichtKlass = AllgemeinerAnsichtMitRegionShapeLinien(self.SchnittAnsicht)
        return self._SchnittAnsichtKlass

    @property
    def MittelPointLinie(self):
        if self.SchnittAnsichtKlass:
            return self.SchnittAnsichtKlass.ShapeMittelPoint
    
    @property
    def MittelPoint(self):
        if self.MittelPointLinie:
            p = self.MittelPointLinie - self.SchnittAnsicht.ViewDirection * \
            (self.Param_SchnittTiefe.AsDouble()) / 2
            return p
    
    @property
    def Param_SchnittTiefe(self):
        return self.elem.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
    
    def MINMaxBerechnen(self,Liste_elems):
        max_x,max_y = -10000000,-10000000
        min_x,min_y = 10000000,10000000
        for el in Liste_elems:
            box_temp = el.get_BoundingBox(None)
            if box_temp:
                max_x = max(max_x,box_temp.Max.X)
                max_y = max(max_y,box_temp.Max.Y)
                min_x = min(min_x,box_temp.Min.X)
                min_y = min(min_y,box_temp.Min.Y)
                box_temp.Dispose()
        return max_x,max_y,min_x,min_y
    
    def verschieben(self,Liste_elems):
        pinned = self.elem.Pinned
        max_x,max_y,min_x,min_y = self.MINMaxBerechnen(Liste_elems)        
        cx = (max_x + min_x) /2
        cy = (max_y + min_y) /2
        p = self.MittelPoint
        self.elem.Pinned = False
        self.elem.Location.Move(DB.XYZ((cx- p.X),(cy-p.Y),0))
        self.elem.Pinned = pinned
    
    def AnsichtsTiefeBerechnen(self,max_x,max_y,min_x,min_y):
        p_vektor0 = DB.XYZ(max_x-min_x,max_y-min_y,0)
        p_vektor1 = DB.XYZ(min_x-max_x,max_y-min_y,0)
        length0 = abs(p_vektor0.DotProduct(self.SchnittAnsicht.ViewDirection.Normalize()))  + 2
        length1 = abs(p_vektor1.DotProduct(self.SchnittAnsicht.ViewDirection.Normalize()))  + 2
        length = max(length0,length1)
        return length
    
    def verschieben_AnsichtsTiefe(self,Liste_elems):
        pinned = self.elem.Pinned
        max_x,max_y,min_x,min_y = self.MINMaxBerechnen(Liste_elems)    
        length = self.AnsichtsTiefeBerechnen(max_x,max_y,min_x,min_y)    
        self.Param_SchnittTiefe.Set(length)
        self.doc.Regenerate()
        cx = (max_x + min_x) /2
        cy = (max_y + min_y) /2
        p = self.MittelPoint
        self.elem.Pinned = False
        self.elem.Location.Move(DB.XYZ((cx- p.X),(cy-p.Y),0))
        self.elem.Pinned = pinned

    
    def verschieben_dynamische(self,Liste_elems):
        max_x,max_y = -1100000000,-10000000
        min_x,min_y = 10000000,10000000
        for el in Liste_elems:
            box_temp = el.get_BoundingBox(None)
            max_x = max(max_x,box_temp.Max.X)
            max_y = max(max_y,box_temp.Max.Y)
            min_x = min(min_x,box_temp.Min.X)
            min_y = min(min_y,box_temp.Min.Y)
            box_temp.Dispose()
        
        
        cx = (max_x + min_x) /2
        cy = (max_y + min_y) /2

        box1 = self.elem.get_BoundingBox(UI.UIDocument(self.elem.Document).ActiveView)
        origin_0 = (box1.Max + box1.Min) /2
        box1.Dispose()

        p_vektor0 = DB.XYZ((max_x-min_x),(max_y-min_y),0)
        p_vektor1 = DB.XYZ((min_x-max_x),(max_y-min_y),0)
        length0 = abs(p_vektor0.DotProduct(self.schnittanschit.ViewDirection.Normalize()))  + 2
        length1 = abs(p_vektor1.DotProduct(self.schnittanschit.ViewDirection.Normalize()))  + 2
        length = max(length0,length1)
        self.elem.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR).Set(length)

        o = origin_0 - self.schnittanschit.ViewDirection * \
            (length+ 0.5) / 2
        
        self.elem.Location.Move(DB.XYZ((cx- o.X),(cy-o.Y),0))
    
    def verschieben_dynamische2(self,Liste_elems):
        Liste_X = []
        Liste_Y = []
        Liste_Z = []
        for el in Liste_elems:
            box_temp = el.get_BoundingBox(None)
            Liste_X.append(box_temp.Max.X)
            Liste_X.append(box_temp.Min.X)
            Liste_Y.append(box_temp.Max.Y)
            Liste_Y.append(box_temp.Min.Y)
            Liste_Z.append(box_temp.Max.Z)
            Liste_Z.append(box_temp.Min.Z)
            box_temp.Dispose()
        
        max_point = DB.XYZ(max(Liste_X),max(Liste_Y),max(Liste_Z))
        min_point = DB.XYZ(min(Liste_X),min(Liste_Y),min(Liste_Z))

        p_vektor0 = DB.XYZ((max(Liste_X)-min(Liste_X)),(max(Liste_Y)-min(Liste_Y)),0)
        p_vektor1 = DB.XYZ((min(Liste_X)-max(Liste_X)),(min(Liste_Y)-max(Liste_Y)),0)
        length0 = abs(p_vektor0.DotProduct(self.schnittanschit.ViewDirection.Normalize()))  + 2
        length1 = abs(p_vektor1.DotProduct(self.schnittanschit.ViewDirection.Normalize()))  + 2
        length = max(length0,length1)
        self.elem.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR).Set(length)

        transform = self.schnittanschit.CropBox.Transform
        box = self.schnittanschit.CropBox

        point1 = transform.Inverse.OfPoint(DB.XYZ(max(Liste_X)+0.5,max(Liste_Y)+0.5,max(Liste_Z)+0.5))
        point2 = transform.Inverse.OfPoint(DB.XYZ(max(Liste_X)+0.5,min(Liste_Y)-0.5,max(Liste_Z)+0.5))
        point3 = transform.Inverse.OfPoint(DB.XYZ(min(Liste_X)-0.5,max(Liste_Y)+0.5,max(Liste_Z)+0.5))
        point4 = transform.Inverse.OfPoint(DB.XYZ(min(Liste_X)-0.5,min(Liste_Y)-0.5,max(Liste_Z)+0.5))
        point5 = transform.Inverse.OfPoint(DB.XYZ(max(Liste_X)+0.5,max(Liste_Y)+0.5,min(Liste_Z)-0.5))
        point6 = transform.Inverse.OfPoint(DB.XYZ(max(Liste_X)+0.5,min(Liste_Y)-0.5,min(Liste_Z)-0.5))
        point7 = transform.Inverse.OfPoint(DB.XYZ(min(Liste_X)-0.5,max(Liste_Y)+0.5,min(Liste_Z)-0.5))
        point8 = transform.Inverse.OfPoint(DB.XYZ(min(Liste_X)-0.5,min(Liste_Y)-0.5,min(Liste_Z)-0.5))

        max_x = max([point1.X,point2.X,point3.X,point4.X,point5.X,point6.X,point7.X,point8.X])
        max_y = max([point1.Y,point2.Y,point3.Y,point4.Y,point5.Y,point6.Y,point7.Y,point8.Y])
        max_z = max([point1.Z,point2.Z,point3.Z,point4.Z,point5.Z,point6.Z,point7.Z,point8.Z])
        min_x = min([point1.X,point2.X,point3.X,point4.X,point5.X,point6.X,point7.X,point8.X])
        min_y = min([point1.Y,point2.Y,point3.Y,point4.Y,point5.Y,point6.Y,point7.Y,point8.Y])
        min_z = min([point1.Z,point2.Z,point3.Z,point4.Z,point5.Z,point6.Z,point7.Z,point8.Z])

        
        # box = DB.BoundingBoxXYZ()
        box.Max = DB.XYZ(max_x,max_y,max_z)
        box.Min = DB.XYZ(min_x,min_y,min_z)
        # box.Transform = transform
        
        self.schnittanschit.CropBox = box

class Grundriss(object):
    def __init__(self,elem):
        self.elem = elem
        self.schnitt_Liste = ObservableCollection[SchnittInGrundriss]()
    
    def getAllSchnitt(self):
        self.schnitt_Liste.Clear()
        coll = DB.FilteredElementCollector(self.elem.Document,self.elem.Id).OfCategory(DB.BuiltInCategory.OST_Viewers).WhereElementIsNotElementType().ToElements()
        _dict = {el.Name:el for el in coll}
        for name in sorted(_dict.keys()):
            self.schnitt_Liste.Add(SchnittInGrundriss(_dict[name]))
        return self.schnitt_Liste
    
    @staticmethod
    def get_schnitt_ansciht(elem):
        return Grundriss(elem).getAllSchnitt()


                
# class RPLItem(RVItem):
#     def __init__(self,name,elemid,doc):
#         RVItem.__init__(self,name,elemid,doc)