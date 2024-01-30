# coding: utf8
from System import Enum
from IGF_Klasse import DB, ItemTemplateNurElem
from IGF_Funktionen._Parameter import get_value
from System.Collections.ObjectModel import ObservableCollection
import clr


class LinkedWandDecken(object):
    def __init__(self,elem,revitlink):
        self.elem = elem
        self.revitlink = revitlink
        self.elemid = self.elem.Id
        self.TypName = self.elem.Name
        self.Solids = []
        self._Type = None        
        self.Brandschutz = self.get_Brandschutz()
        self._Geometrie  = None
        if self.Kategorie == 'Geschossdecken' and self.ElementTyp == '???':
            self.Brandschutz = None

    @property
    def doc(self):
        return self.elem.Document
    
    @property
    def Kategorie(self):
        return self.elem.Category.Name
    
    @property
    def TextBrandschutz(self):
        try:return self.elem.LookupParameter('IGF_A_Brandschutz').AsString()
        except:return ''
    
    @property
    def ElementTyp(self):
        if not self._Type:
            if self.TypName.upper().find('TREPPE') != -1:
                self._Type = 'TREPPE'
                return self._Type
            elif self.TypName.upper().find('BETON') != -1:
                self._Type  = 'Beton'
                return self._Type
            elif self.TypName.upper().find('GIPS') != -1 or self.TypName.upper().find('TROCKENBAU') != -1 or self.TypName.upper().find('METALL') != -1 or self.TypName.upper().find('RASTER') != -1 or self.TypName.upper().find('TK') != -1:
                self._Type  = 'Trockenbau'
                return self._Type
            elif self.TypName.upper().find('MAUERWERK') != -1 or self.TypName.upper().find('ZIEGEL') != -1:
                self._Type  = 'Mauerwerk'
                return self._Type
            
            else:
                Materiels = self.elem.GetMaterialIds(False)
                for m in Materiels:
                    try:
                        name = self.doc.GetElement(m).Name
                        if name.upper().find('BETON') != -1:
                            self._Type  = 'Beton'
                            return self._Type
                        elif name.upper().find('GIPS') != -1 or self.TypName.upper().find('TROCKENBAU') != -1 or self.TypName.upper().find('METALL') != -1:
                            self._Type  = 'Trockenbau'
                            return self._Type
                        elif name.upper().find('MAUERWERK') != -1 or self.TypName.upper().find('ZIEGEL') != -1:
                            self._Type  = 'Mauerwerk'
                            return self._Type
                    except:
                        pass
        if not self._Type:
            self._Type = '???'
        return self._Type

    def get_Brandschutz(self):
        if self.Kategorie == 'Geschossdecken':
            if self.TypName.upper().find('ENSCAPE') != -1:
                return None
            try:
                if self.elem.LookupParameter('Dicke').AsDouble()*0.3048 <= 0.061:
                    return None
            except:
                return None

            if self.TypName.upper().find('F120') != -1 or self.TextBrandschutz.upper().find('F120') != -1:
                return 'F120'
            elif self.TypName.upper().find('F90') != -1 or self.TextBrandschutz.upper().find('F90') != -1:
                return 'F90'
            elif self.TypName.upper().find('F60') != -1 or self.TextBrandschutz.upper().find('F60') != -1:
                return 'F60'
            elif self.TypName.upper().find('F30') != -1 or self.TextBrandschutz.upper().find('F30') != -1:
                return 'F30'
            elif self.TypName.upper().find('BW') != -1 or self.TextBrandschutz.upper().find('BW') != -1:
                return 'BW'
            elif self.TypName.upper().find('BAUART BRANDWAND') != -1 or self.TextBrandschutz.upper().find('BAUART BRANDWAND') != -1:
                return 'Bauart BW'
            else:
                return 'F90'
        else:
            if self.TypName.upper().find('F120') != -1 or self.TextBrandschutz.upper().find('F120') != -1:
                return 'F120'
            elif self.TypName.upper().find('F90') != -1 or self.TextBrandschutz.upper().find('F90') != -1:
                return 'F90'
            elif self.TypName.upper().find('F60') != -1 or self.TextBrandschutz.upper().find('F60') != -1:
                return 'F60'
            elif self.TypName.upper().find('F30') != -1 or self.TextBrandschutz.upper().find('F30') != -1:
                return 'F30'
            elif self.TypName.upper().find('BW') != -1 or self.TextBrandschutz.upper().find('BW') != -1:
                return 'BW'
            elif self.TypName.upper().find('BAUART BRANDWAND') != -1 or self.TextBrandschutz.upper().find('BAUART BRANDWAND') != -1:
                return 'Bauart BW'
            
            else:
                return None

    @property
    def BrandschutzWand(self):
        if self.Brandschutz:
            return True
        else:
            return False

    @property
    def Geometrie(self):
        if not self._Geometrie:
            if self.BrandschutzWand:
                try:self._Geometrie = self.TransformSolid()
                except Exception as e:print(e)
        return self._Geometrie
    
    def GetSolids(self):
        def GetSolid(GeoElement):
            lstSolid = []
            for el in GeoElement:
                if el.GetType().ToString() == 'Autodesk.Revit.DB.Solid':
                    if el.SurfaceArea > 0 and el.Volume > 0 and el.Faces.Size > 1 and el.Edges.Size > 1:
                        lstSolid.append(el)
                elif el.GetType().ToString() == 'Autodesk.Revit.DB.GeometryInstance':
                    ge = el.GetInstanceGeometry()
                    lstSolid.extend(GetSolid(ge))
            return lstSolid
        lstSolid = []
        opt = DB.Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        ge = self.elem.get_Geometry(opt)
        if ge != None:
            lstSolid.extend(GetSolid(ge))
        opt.Dispose()
        return lstSolid
    
    def TransformSolid(self):
        m_lstModels = []
        listSolids = self.GetSolids()
        for solid in listSolids:
            tempSolid = DB.SolidUtils.CreateTransformed(solid,self.revitlink.GetTransform())
            m_lstModels.append(tempSolid)
        for solid in listSolids:
            solid.Dispose()
        out_solid = None
        if len(m_lstModels) > 1:
            for solid in m_lstModels:
                if not out_solid:
                    out_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(solid,solid,DB.BooleanOperationsType.Union)
                else:
                    temp = out_solid
                    out_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(temp,solid,DB.BooleanOperationsType.Union)
                    temp.Dispose()
            for solid in m_lstModels:
                solid.Dispose()
            return out_solid
        elif len(m_lstModels) == 1:
            return m_lstModels[0]
       
class ZuPruefendElement(object):
    def __init__(self,elem,BrandschutzElement):
        self.BSK_Liste = ['BEK,' 'BSK', 'BRANDSCHUTZKLAPPE', 'F30', 'F60', 'F90', 'RSK', 'RAUCHSCGUTZKLAPPE', 'ENTRAUCHUNGSKLAPPE','DRM_DAMPER','FDM_FIRE_DAMPER']
        self.elem = elem
        self.BrandschutzElement = BrandschutzElement
        self.elemid = self.elem.Id
        self.Brandklass = {}
        self.Brandklass_Text = '' 
        self.Anzahl = 0
        self.Pruefen = 'Ja'
        self._Linie = None
        self._Anzahl_BSK = None
        self._Liste_BSK_Connector = []
        self.Liste_Internal_Connector = []
        self.BrandschutzPruefen()
        self.BSK_Pruefen()
    
    @property
    def doc(self):
        return self.elem.Document
    
    def Conn_pruefen(self,conn):
        refs = conn.AllRefs
        for ref in refs:
            try:
                famname = ref.Owner.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
                for e in self.BSK_Liste:
                    if famname.upper().find(e) != -1:
                        self._Liste_BSK_Connector.append(conn.Origin)
                        return 1
            except:
                pass
        return 0
    @property
    def Anzahl_BSK(self):
        if self._Anzahl_BSK is None:
            self._Liste_BSK_Connector = []
            anzahl = 0
            self._Anzahl_BSK = 0
            if self.elem.Category.Id.ToString() in ['-2008000','-2008010']:
                conns = None
                try:conns = list(self.elem.ConnectorManager.Connectors)
                except:
                    try:conns = list(self.elem.MEPModel.ConnectorManager.Connectors)
                    except:pass
                if conns:
                    for conn in conns:
                        anzahl += self.Conn_pruefen(conn)
            self._Anzahl_BSK = anzahl
        return self._Anzahl_BSK

    @property
    def Linie(self):
        if not self._Linie:
            try:self._Linie = self.elem.Location.Curve
            except:pass    
        return self._Linie
    
    @property
    def NichtPruefen(self):
        try:return self.elem.LookupParameter('IGF_HLS_Brandschott_prüfen').AsString().upper().find('NEIN') != -1
        except:pass    
  
    def get_Brandklass_Text(self):
        try:
            anzahl = 0
            for cate in self.Brandklass:
                for typ in self.Brandklass[cate]:
                    for klass in self.Brandklass[cate][typ]:
                        self.Brandklass_Text += str(self.Brandklass[cate][typ][klass]) + 'x' + cate+'-'+typ+'-'+klass +', '
                        anzahl += self.Brandklass[cate][typ][klass]
            self.Brandklass_Text = self.Brandklass_Text[:-2] 
            if anzahl != self.Anzahl and anzahl != 0:
                self.Pruefen += ' Warnung: Anzahl unterschiedlich.'
        except:pass
        
    def wert_schreiben(self):
        if self.Brandklass_Text == None:
            self.Brandklass_Text = ""

        try:self.elem.LookupParameter('IGF_HLS_Brandschutz').Set(self.Brandklass_Text)
        except:pass
        
        try:self.elem.LookupParameter('IGF_HLS_Brandschott').Set(self.Anzahl)
        except:pass
        
        try:self.elem.LookupParameter('IGF_HLS_Brandschott_prüfen').Set(self.Pruefen)
        except:pass

    def BrandschutzPruefen(self):
        if not self.NichtPruefen:
            if len(self.BrandschutzElement) > 0:
                if self.Linie:
                    opt_outside = DB.SolidCurveIntersectionOptions()
                    opt_outside.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsOutside
                    opt_inside = DB.SolidCurveIntersectionOptions()
                    opt_inside.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside
                    if len(self.BrandschutzElement) == 1:
                        result_Inside = self.BrandschutzElement[0].Geometrie.IntersectWithCurve(self.Linie,opt_inside)
                        result_Outside = self.BrandschutzElement[0].Geometrie.IntersectWithCurve(self.Linie,opt_outside)
                        if result_Inside.SegmentCount > 0 and result_Outside.SegmentCount > 0:
                            self.Anzahl = result_Inside.SegmentCount
                            if result_Outside.SegmentCount - result_Inside.SegmentCount != 1:
                                self.Pruefen = 'Warnung: Leitung endet in Objekt.'
                            self.Brandklass.setdefault(self.BrandschutzElement[0].Kategorie,{}).setdefault(self.BrandschutzElement[0].ElementTyp,{}).setdefault(self.BrandschutzElement[0].Brandschutz,0)
                            self.Brandklass[self.BrandschutzElement[0].Kategorie][self.BrandschutzElement[0].ElementTyp][self.BrandschutzElement[0].Brandschutz] = 1

                        elif result_Outside.SegmentCount == 0 and result_Inside.SegmentCount > 0:
                            self.Anzahl = result_Inside.SegmentCount
                            self.Pruefen = 'Warnung: Leitung in Objekt.'
                            self.Brandklass.setdefault(self.BrandschutzElement[0].Kategorie,{}).setdefault(self.BrandschutzElement[0].ElementTyp,{}).setdefault(self.BrandschutzElement[0].Brandschutz,0)
                            self.Brandklass[self.BrandschutzElement[0].Kategorie][self.BrandschutzElement[0].ElementTyp][self.BrandschutzElement[0].Brandschutz] = 1

                        else:
                            self.Anzahl = 1
                            self.Pruefen = 'Warnung: manuell prüefen. Wahrscheinlich Kollision'
                        if result_Inside.SegmentCount > 0:
                            for n in range(result_Inside.SegmentCount):
                                self.Liste_Internal_Connector.append(result_Inside.GetCurveSegment(n).GetEndPoint(0))
                                self.Liste_Internal_Connector.append(result_Inside.GetCurveSegment(n).GetEndPoint(1))

                        result_Inside.Dispose()
                        result_Outside.Dispose()
                        opt_inside.Dispose()
                        opt_outside.Dispose()
                        self.get_Brandklass_Text()
                        self.Linie.Dispose()
                    
                    else:
                        gesamt = None
                        KeinProblem = True
                        for element in self.BrandschutzElement:
                            if not gesamt:
                                gesamt = DB.BooleanOperationsUtils.ExecuteBooleanOperation(element.Geometrie,element.Geometrie,DB.BooleanOperationsType.Union)
                            else:
                                temp = gesamt
                                try:
                                    gesamt = DB.BooleanOperationsUtils.ExecuteBooleanOperation(element.Geometrie,temp,DB.BooleanOperationsType.Union)
                                    temp.Dispose()
                                except:
                                    KeinProblem = False
                                    temp.Dispose()
                                    break
                        
                        if not KeinProblem:
                            self.Anzahl = 1
                            try:gesamt.Dispose()
                            except:pass
                        else:
                            result_Inside = gesamt.IntersectWithCurve(self.Linie,opt_inside)
                            result_Outside = gesamt.IntersectWithCurve(self.Linie,opt_outside)
                            if result_Inside.SegmentCount > 0 and result_Outside.SegmentCount > 0:
                                self.Anzahl = result_Inside.SegmentCount
                                if result_Outside.SegmentCount - result_Inside.SegmentCount != 1:
                                    self.Pruefen = 'Warnung: Leitung endet in Objekt.'
                            elif result_Outside.SegmentCount == 0 and result_Inside.SegmentCount > 0:
                                self.Anzahl = result_Inside.SegmentCount
                                self.Pruefen = 'Warnung: Leitung in Objekt.'
                            else:
                                self.Anzahl = 1
                                self.Pruefen = 'Warnung: manuell prüefen. Wahrscheinlich Kollision.'
                            if result_Inside.SegmentCount > 0:
                                for n in range(result_Inside.SegmentCount):
                                    self.Liste_Internal_Connector.append(result_Inside.GetCurveSegment(n).GetEndPoint(0))
                                    self.Liste_Internal_Connector.append(result_Inside.GetCurveSegment(n).GetEndPoint(1))
                            result_Inside.Dispose()
                            result_Outside.Dispose()
                            gesamt.Dispose()
                            for element in self.BrandschutzElement:
                                result_Inside = element.Geometrie.IntersectWithCurve(self.Linie,opt_inside)
                                if result_Inside.SegmentCount > 0:
                                    self.Brandklass.setdefault(element.Kategorie,{}).setdefault(element.ElementTyp,{}).setdefault(element.Brandschutz,0)
                                    self.Brandklass[element.Kategorie][element.ElementTyp][element.Brandschutz] += 1
                                result_Inside.Dispose()
                            opt_inside.Dispose()
                            opt_outside.Dispose()
                            self.get_Brandklass_Text()
                            self.Linie.Dispose()

                else:
                    self.Anzahl = 1
                    try:self.Liste_Internal_Connector = [conn.Origin for conn in self.elem.MEPModel.ConnectorManager.Connectors]
                    except:pass
                    for element in self.BrandschutzElement:
                        self.Brandklass.setdefault(element.Kategorie,{}).setdefault(element.ElementTyp,{}).setdefault(element.Brandschutz,0)
                        self.Brandklass[element.Kategorie][element.ElementTyp][element.Brandschutz] += 1
                    self.get_Brandklass_Text()

    def BSK_Pruefen(self):
        if self.Anzahl_BSK:
            if self.Anzahl == 0:
                self.Pruefen = 'BSK bereits eingesetzt.'
            else:
                if len(self.Liste_Internal_Connector) > 0:
                    for BSK_point in self._Liste_BSK_Connector:
                        for point in self.Liste_Internal_Connector:
                            if point.IsAlmostEqualTo(BSK_point):
                                self.Anzahl -= 1
                                break
                    if self.Anzahl <= 0:
                        self.Anzahl = 0
                        self.Pruefen = 'BSK bereits eingesetzt.'
                    else:
                        if self.Pruefen == 'Ja':self.Pruefen = 'BSK bereits eingesetzt.'
                        else:self.Pruefen += 'BSK bereits eingesetzt.'
                

