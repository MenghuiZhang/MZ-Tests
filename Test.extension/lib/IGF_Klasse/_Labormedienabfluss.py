# coding: utf8
from IGF_Klasse import DB,ItemTemplate,TemplateItemBase
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from IGF_Funktionen._Geometrie import LinkedPointTransform,LinkedVectorTransform,CurrentPointTransformToLink
from IGF_Funktionen._Parameter import wert_schreibenbase
from IGF_Klasse._familie import Familie
from IGF_Klasse._parameter import LaboranschlussItemBase
from IGF_Funktionen._Allgemein import Nummer_Korrigieren
            
class LaborMedien_IGF(TemplateItemBase):
    def __init__(self,elem):
        TemplateItemBase.__init__(self)
        self.elem = elem
        self.doc = elem.Document
        self.LaborMedien_Heinekamp = None
    @property
    def Connector(self):
        return list(self.elem.MEPmodel.ConnectorManager.Connectors)[0]
    
    @property
    def Loaction(self):
        return self.Connector.Origin
    
    @property
    def Richtung(self):
        return self.Connector.CoordinateSystem.BasisZ
    
    def Trennen(self):
        if self.Connector.IsConnected:
            for ref in self.Connector.AllRefs:
                if ref.Owner.Category.CategoryType.ToString() == 'Model':
                    self.Connector.DisconnectFrom(ref)

    
    def Move(self,Location):
        if not self.Location.IsAlmostEqualTo(Location):
            self.Trennen()
            self.elem.Location.Move(Location-self.Location)
    
    def Rotate(self,Richtung):
        if self.Richtung.IsAlmostEqualTo(Richtung):
            angle = 0
        elif self.Richtung.IsAlmostEqualTo(Richtung.Negate()):
            if self.Richtung.Y==self.Richtung.Z==self.Richtung.X:
                v0 = self.Richtung + DB.XYZ(0,0,1)
            else:
                v0 = self.Richtung + DB.XYZ(self.Richtung.Y,self.Richtung.Z,self.Richtung.X)
            vector = v0-v0.DotProduct(self.Richtung)*self.Richtung
            angle = self.Richtung.AngleTo(Richtung)
        else:
            vector = self.Richtung.CrossProduct(Richtung)
            angle = self.Richtung.AngleTo(Richtung)
        
        if angle > 0:
            self.Trennen()
            DB.ElementTransformUtils.RotateElement(self.doc,self.elem.Id,DB.Line.CreateUnbound(self.Location,vector),angle)
            self.doc.Regenerate()
            if not self.Richtung.IsAlmostEqualTo(Richtung):
                DB.ElementTransformUtils.RotateElement(self.doc,self.elem.Id,DB.Line.CreateUnbound(self.Location,vector),- 2 * angle)

class LaborMedien_Heinekamp_Abfluss(TemplateItemBase):
    '''
    elem:revit element
    rvtlink:revit Link
    Liste_Familie:IGF_Klasse._labormedien.LaboranschlussType
    '''
    def __init__(self,elem,rvtlink,Familietype,Hoehe = 10):
        TemplateItemBase.__init__(self)
        self.Hoehe = Hoehe
        self.elem = elem
        self.rvtlink = rvtlink
        self.currectdoc = self.rvtlink.Document
        self.rvtdoc = self.elem.doc
    
    @property
    def OriginaleLocation(self):
        return self.elem.Location.Point
    
    @property
    def CurrenteLocation(self):
        return LinkedPointTransform(self.rvtlink, self.OriginaleLocation)
    
    @property
    def MEPRaum(self):
        return self.currectdoc.GetSpaceAtPoint(self.CurrenteLocation) + DB.XYZ(0,0,1)
    
    @property
    def Richtung(self):
        return DB.XYZ(0,0,-1)
    
    @property
    def BasisLevel(self):
        if self.MEPRaum:
            return self.MEPRaum.Level
        return None
    
    def PositionLocation(self):
        return DB.XYZ(self.CurrenteLocation.X,self.CurrenteLocation.Y,self.Hoehe/304.8+self.BasisLevel.get_Parameter(DB.BuiltInParameter.LEVEL_ELEV).AsDouble())
     
    def CreateLaborAnschluss(self):
        self.Familietype.Activate()
        newfamilie = self.currectdoc.Create.NewFamilyInstance(self.PositionLocation,self.Familietype,self.BasisLevel,DB.Structure.StructuralType.NonStructural)
        self.currectdoc.Regenerate()
        laborMedienIGF = LaborMedien_IGF(newfamilie)
        laborMedienIGF.LaborMedien_Heinekamp = self
        self.LaborMedien_IGF = laborMedienIGF
        laborMedienIGF.Move(self.PositionLocation)
        self.currectdoc.Regenerate()
        laborMedienIGF.Rotate(self.Richtung)

    def LaborAnschlussAnpassen(self):
        pinned = self.LaborMedien_IGF.elem.Pinned
        self.LaborMedien_IGF.elem.Pinned = False
        self.LaborMedien_IGF.Move(self.PositionLocation)
        self.currectdoc.Regenerate()
        self.LaborMedien_IGF.Rotate(self.Richtung)
        self.LaborMedien_IGF.elem.Pinned = pinned
    
    def werte_schreiben(self):
        if self.ConnectorCount > 1:
            for el in self.SUB_Labor:
                el.werte_schreiben()
        else:
            if self.LaborMedien_IGF:
                def wertschreiben(paramname,wert):
                    wert_schreibenbase(self.LaborMedien_IGF.elem.LookupParameter(paramname), wert)
                wertschreiben('IGF_L_Connector_DN', self.Durchmesser)
                wertschreiben('IGF_L_Connector_LuftmengeMin', self.Luftmengemin)
                wertschreiben('IGF_L_Connector_LuftmengeMax', self.Luftmengemax)
                wertschreiben('IGF_L_Connector_LuftmengeNacht', self.Luftmengemin)
                wertschreiben('IGF_L_Connector_LuftmengeTiefenacht', self.Luftmengemin)
                wertschreiben('IGF_L_Connector_Druckverlust', self.Druckverlust)
                wertschreiben('IGF_L_Connector_Druckverlust', self.Druckverlust)
                wertschreiben('IGF_L_Connector_Breite', self.Breite)
                wertschreiben('DIN_a', self.Durchmesser)
                wertschreiben('IGF_L_Connector_Höhe', self.Hoehe)
                wertschreiben('IGF_L_Connector_FamilieTyp', self.familietyp)
                wertschreiben('IGF_X_Einbauort', self.MEPRaumNummer)
                wertschreiben('IGF_X_EinbauDetails', 'Zeile ' + self.Zeile + ', Nr ' + str(self.Nummer))
                wertschreiben('IGF_X_Auftragnehmer_Exemplar',self.Auftragnehmer)
                wertschreiben('IGF_X_Relevant',self.Relevant)
                if self.Relevant:
                    wert_schreibenbase(self.LaborMedien_IGF.elem.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM), self.Luftmengemax)
                else:
                    wert_schreibenbase(self.LaborMedien_IGF.elem.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM), self.Luftmengemin)
                

class LaborMedien_Heinekamp_SUB(object):
    '''
    elem:revit element
    rvtlink:revit Link
    Liste_Familie:IGF_Klasse._labormedien.LaboranschlussType
    '''
    def __init__(self,Parent,Connector):
        self.Parent = Parent
        self._Connector = Connector      
        self._auslasstyp = None  
        self.Subnummer = 0
        self.LaborMedien_IGF = None
        self.currectdoc = self.Parent.currectdoc
    
    @property
    def Zuordnung(self):
        return self.Parent.Zuordnung
    
    @property
    def Kanalnetz(self):
        return self.Parent.Kanalnetz
    
    @property
    def Ist24h(self):
        return self.Parent.Ist24h
    
    @property
    def Art(self):
        return self.Parent.Art

    @property
    def Connector(self):
        return self._Connector

    @property
    def Richtung(self):
        return LinkedVectorTransform(self.Parent.rvtlink, self.Connector.CoordinateSystem.BasisZ)

    @property
    def Location(self):
        return LinkedPointTransform(self.Parent.rvtlink, self.Connector.Origin)
    
    @property
    def Shape(self):
        if self.Connector.Shape.ToString() == 'Round':
            return 'RU'
        return 'RE'
    
    @property
    def ZeileKoreper(self):
        return self.Parent.ZeileKoreper
    
    @property
    def Zeile(self):
        return self.Parent.Zeile

    @property
    def Position(self):
        if self.ZeileKoreper:
            return self.ZeileKoreper.HandOrientation.DotProduct(self.Location)
        return 0

    @property
    def Auslassstyp(self):
        if not self._auslasstyp:
            if self.Parent.Liste_Familie:
                for el in self.Parent.Liste_Familie:
                    if el.Art == self.Art and el.Shape == self.Shape and self.Ist24h == el.Ist24h:
                        for typ in el.Dict_Types.keys():
                            laborparam = LaboranschlussItemBase(typ)
                            if laborparam.luftmengemin == self.Parent.Luftmengemin and laborparam.luftmengemax == self.Parent.Luftmengemax \
                                and laborparam.druckverlust == self.Druckverlust and laborparam.durchmesser == self.Durchmesser\
                                    and laborparam.breite == self.Breite and laborparam.hoehe == self.Hoehe and laborparam.typenameohnedetail == self.Zuordnung and laborparam.IstSuB:
                                self._auslasstyp = el.Dict_Types[typ]
                                return self._auslasstyp
                            del laborparam
                        try:
                            return el.Dict_Types['Allgemein']
                        except:
                            return el.defaulttype
        return self._auslasstyp
    
    @property
    def Auslasstypname(self):
        if self.Auslassstyp:
            return self.Auslassstyp.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    
    @property  
    def ParameterName(self):
        if self.Auslasstypname.find('_'+str(self.Druckverlust)+'Pa') != -1 and self.Auslasstypname.find(str(self.Luftmengemin)) != -1 and self.Auslasstypname.find('_unter') != -1:
            return self.Auslasstypname
        else:
            text = 'IGF_RLT_Laboranschluss_'
            if self.Ist24h == '24h':
                text += '24h_'
            elif self.Art == 'Zuluft':
                text += 'LZU_'
            else:
                text += 'LAB_'
            text += self.Zuordnung + '_' + str(self.Druckverlust) + 'Pa_'
            if self.Shape == 'RU':
                text += 'DN' + str(self.Durchmesser)
            else:
                text += str(self.Breite) + 'x' + str(self.Hoehe)
            text += '_min' +  str(self.Parent.Luftmengemin) + '_max' + str(self.Parent.Luftmengemax) + '_unter'
            return text
    
    @property
    def Luftmengemin(self):
        return self.Parent.Luftmengemin / self.Parent.ConnectorCount
    
    @property
    def Luftmengemax(self):
        return self.Parent.Luftmengemax / self.Parent.ConnectorCount

    @property
    def Durchmesser(self):
        if self.Shape == 'RU':
            return int(round(self.Connector.Radius * 2 * 304.8))
        return None
    
    @property
    def Auftragnehmer(self):
        return self.Parent.Auftragnehmer
    
    @property
    def Breite(self):
        if self.Shape == 'RE':
            return int(round(self.Connector.Width * 304.8))
        return None
    
    @property
    def Hoehe(self):
        if self.Shape == 'RE':
            return int(round(self.Connector.Height * 304.8))
        return None
    
    @property
    def MEPRaum(self):
        return self.Parent.MEPRaum
    
    @property
    def Druckverlust(self):
        return self.Parent.Druckverlust
    
    def CreateLaborAnschluss(self):
        if self.Auslasstypname != self.ParameterName:
            self._auslasstyp = self.Auslassstyp.Duplicate(self.ParameterName)
            self.currectdoc.Regenerate()
            LaboranschlussType.Anpassen(self._auslasstyp,  LaboranschlussItemBase(self.ParameterName))
        self.Auslassstyp.Activate()
        newfamilie = self.currectdoc.Create.NewFamilyInstance(self.Location,self.Auslassstyp,self.MEPRaum.Level,DB.Structure.StructuralType.NonStructural)
        self.currectdoc.Regenerate()
        laborMedienIGF = LaborMedien_IGF(newfamilie)
        laborMedienIGF.LaborMedien_Heinekamp = self
        self.LaborMedien_IGF = laborMedienIGF
        laborMedienIGF.Move(self.Location)
        self.currectdoc.Regenerate()
        laborMedienIGF.Rotate(self.Richtung)
    
    def LaborAnschlussAnpassen(self):
        pinned = self.LaborMedien_IGF.elem.Pinned
        self.LaborMedien_IGF.elem.Pinned = False
        self.LaborMedien_IGF.Move(self.Location)
        self.Parent.currectdoc.Regenerate()
        self.LaborMedien_IGF.Rotate(self.Richtung)
        self.LaborMedien_IGF.elem.Pinned = pinned
    
    def werte_schreiben(self):
        if self.LaborMedien_IGF:
            def wertschreiben(paramname,wert):
                wert_schreibenbase(self.LaborMedien_IGF.elem.LookupParameter(paramname), wert)
            wertschreiben('IGF_L_Connector_DN', self.Durchmesser)
            wertschreiben('IGF_L_Connector_LuftmengeMin', self.Luftmengemin)
            wertschreiben('IGF_L_Connector_LuftmengeMax', self.Luftmengemax)
            wertschreiben('IGF_L_Connector_LuftmengeNacht', self.Luftmengemin)
            wertschreiben('IGF_L_Connector_LuftmengeTiefenacht', self.Luftmengemin)
            wertschreiben('IGF_L_Connector_Druckverlust', self.Druckverlust)
            wertschreiben('DIN_a', self.Durchmesser)
            wertschreiben('IGF_L_Connector_Breite', self.Breite)
            wertschreiben('IGF_L_Connector_Höhe', self.Hoehe)
            wertschreiben('IGF_L_Connector_FamilieTyp', self.Parent.familietyp)
            wertschreiben('IGF_X_Einbauort', self.Parent.MEPRaumNummer)
            wertschreiben('IGF_X_EinbauDetails', 'Zeile ' + self.Zeile + ', Nr ' + str(self.Parent.Nummer) + self.Subnummer)
            wertschreiben('IGF_X_Auftragnehmer_Exemplar',self.Auftragnehmer)
            wertschreiben('IGF_X_Relevant',self.Parent.Relevant)
            if self.Parent.Relevant:
                wert_schreibenbase(self.LaborMedien_IGF.elem.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM), self.Luftmengemax)
            else:
                wert_schreibenbase(self.LaborMedien_IGF.elem.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM), self.Luftmengemin)
            
  
class LaborMedien_MEPRaum:
    """ 
    MEPRaum = RVT MEP Raum
    rvtlink = RVT Link
    Liste_Familiename = [LaboranschlussFamilieName]
    Liste_Familien = [IGF_Klasse._labormedien.LaboranschlussType]
    """
    def __init__(self,MEPRaum = None,rvtlink = None,Liste_Familiename = [],Liste_Familien = []):
        self.MEPRaum = MEPRaum
        self.doc = self.MEPRaum.Document
        self.Liste_Familiename = Liste_Familiename
        self.Liste_Familien = Liste_Familien
        if not self.Liste_Familiename:self.Liste_Familiename = []
        self.boundingboxmin = self.MEPRaum.get_BoundingBox(None).Min
        self.boundingboxmax = self.MEPRaum.get_BoundingBox(None).Max
        self.rvtlink = rvtlink
        self.Liste_LaborMedien_Heinekamp = []
        self.Liste_LaborMedien_IGF = []
        self.Dict_LaborMedien_Heinekamp = {}
  
    def set_up(self):
        if not self.MEPRaum:return
        self.Liste_LaborMedien_Heinekamp = []
        self.Liste_LaborMedien_IGF = []
        self.Dict_LaborMedien_Heinekamp = {}
        self.get_Liste_LaborMedien_Heinekamp()
        self.get_Liste_LaborMedien_IGF()
        self.get_Data_RVLink()
        self.Match()
    
    def get_Data_RVLink(self):
        dict0 = {}
        dict1 = {}
        for laborMedien in self.Liste_LaborMedien_Heinekamp:
            dict1.setdefault(laborMedien.Zeile,[])
            dict0.setdefault(laborMedien.ParameterName,[])
            dict0[laborMedien.ParameterName].append(laborMedien)
            dict1[laborMedien.Zeile].append(laborMedien)
        self.Dict_LaborMedien_Heinekamp = dict0
        for zeile in dict1.keys():
            liste = dict1[zeile]
            liste.sort(key = lambda x:x.Position)
            for n,labormedien in enumerate(liste):
                labormedien.Nummer = Nummer_Korrigieren(n+1,2)

    def get_Liste_LaborMedien_Heinekamp(self):
        if not (self.MEPRaum and self.rvtlink and self.Liste_Familien):return
        try:
            outline = DB.Outline(CurrentPointTransformToLink(self.rvtlink,self.boundingboxmin),CurrentPointTransformToLink(self.rvtlink,self.boundingboxmax))
            filter0 = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM),'Medieneinspeisung',False))
            filter1 = DB.BoundingBoxIntersectsFilter(outline)
            filter2 = DB.BoundingBoxIsInsideFilter(outline)
            filter3 = DB.LogicalOrFilter(List[DB.ElementFilter]([filter0,filter1,filter2]))
            coll = DB.FilteredElementCollector(self.rvtlink.GetLinkDocument()).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WherePasses(filter3).WhereElementIsNotElementType().ToElements()
            filter0.Dispose()
            filter1.Dispose()
            filter2.Dispose()
            filter3.Dispose()
            if len(coll) == 0:return
            for laborMedien in coll:
                tmp = LaborMedien_Heinekamp(laborMedien, self.rvtlink, self.Liste_Familien)
                if tmp.MEPRaumId == self.MEPRaum.Id.IntegerValue:
                    self.Liste_LaborMedien_Heinekamp.append(tmp)
        except:
            pass
          
    def get_Liste_LaborMedien_IGF(self):
        if not (self.MEPRaum and self.Liste_Familiename):return
        outline = DB.Outline(self.MEPRaum.get_BoundingBox(None).Min,self.MEPRaum.get_BoundingBox(None).Max)
        filter1 = DB.BoundingBoxIntersectsFilter(outline)
        filter2 = DB.BoundingBoxIsInsideFilter(outline)
        filter3 = DB.LogicalOrFilter(List[DB.ElementFilter]([filter1,filter2]))
        coll = DB.FilteredElementCollector(self.doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WherePasses(filter3).WhereElementIsNotElementType().ToElements()
        filter1.Dispose()
        filter2.Dispose()
        filter3.Dispose()
        if len(coll) == 0:return
        for laborMedien in coll:
            if laborMedien.Symbol.FamilyName in self.Liste_Familiename:
                tmp = LaborMedien_IGF(laborMedien)
                if tmp.MEPRaumId == self.MEPRaum.Id.IntegerValue:
                    self.Liste_LaborMedien_IGF.append(tmp)
    
    def Pruefen(self,labor_Heinekamp,labor_IGF):
        if labor_Heinekamp.ParameterName == labor_IGF.typname:
            return True
        else:
            return False

    def Match(self):
        for labor_igf in self.Liste_LaborMedien_IGF:
            for labor_Heinekamp in self.Liste_LaborMedien_Heinekamp:
                if labor_Heinekamp.ConnectorCount == 1:
                    if labor_igf.Location.IsAlmostEqualTo(labor_Heinekamp.Location):
                        if self.Pruefen(labor_Heinekamp, labor_igf):
                            if labor_igf.LaborMedien_Heinekamp:
                                labor_igf.LaborMedien_Heinekamp.LaborMedien_IGF = None
                            labor_igf.LaborMedien_Heinekamp = labor_Heinekamp
                            labor_Heinekamp.LaborMedien_IGF = labor_igf
                        break
                else:
                    for sub in labor_Heinekamp.SUB_Labor:
                        if labor_igf.Location.IsAlmostEqualTo(sub.Location):
                            if self.Pruefen(sub, labor_igf):
                                if labor_igf.LaborMedien_Heinekamp:
                                    labor_igf.LaborMedien_Heinekamp.LaborMedien_IGF = None
                                labor_igf.LaborMedien_Heinekamp = sub
                                sub.LaborMedien_IGF = labor_igf
                            break
            if not labor_igf.LaborMedien_Heinekamp:
                for labor_Heinekamp in self.Liste_LaborMedien_Heinekamp:
                    if labor_Heinekamp.ConnectorCount == 1:
                        if not labor_Heinekamp.LaborMedien_IGF:
                            if self.Pruefen(labor_Heinekamp, labor_igf):
                                labor_igf.LaborMedien_Heinekamp = labor_Heinekamp
                                labor_Heinekamp.LaborMedien_IGF = labor_igf
                                break          
                    else:
                        for sub in labor_Heinekamp.SUB_Labor:
                            if not sub.LaborMedien_IGF:
                                if self.Pruefen(sub, labor_igf):
                                    labor_igf.LaborMedien_Heinekamp = sub
                                    sub.LaborMedien_IGF = labor_igf
                                    break 

