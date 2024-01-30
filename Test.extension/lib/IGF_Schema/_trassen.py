# coding: utf8
from System import Enum
from IGF_Klasse import DB, ItemTemplateNurElem
from IGF_Funktionen._Parameter import get_value
from System.Collections.ObjectModel import ObservableCollection
import clr
from IGF_Funktionen._Geometrie import get_ClosestPoints
from IGF_Funktionen._Parameter import wert_schreibenbase

class Uebergang:
    def __init__(self,elem):
        self.elem = elem
        self._Connectors = None
        self._Connector = None
    
    @property
    def doc(self):
        return self.elem.Document
    
    @property
    def Connectors(self):
        if not self. _Connectors:
            self. _Connectors = list(self.elem.MEPModel.ConnectorManager.Connectors)
        return self._Connectors
    
    @property
    def Connector(self):
        if not self._Connector:
            self._Connector = self.Connectors[0]
        return self._Connector
    
    @property
    def Location(self):
        return self.elem.Location.Point
    
    @property
    def Richtung(self):
        return self.Connector.CoordinateSystem.BasisZ
    
    def Move(self,Point):
        Point0 = (self.Connectors[0].Origin + self.Connectors[1].Origin)/2
        self.elem.Location.Move(Point-Point0)
    
    def ChangeDimension(self,durchmesser):
        mepinfo = self.Connector.GetMEPConnectorInfo()
        d = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_DIAMETER))
        r = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_RADIUS))
        if d != DB.ElementId.InvalidElementId:
            try:wert_schreibenbase(self.elem.get_Parameter(self.doc.GetElement(d).GetDefinition()), durchmesser)
            except Exception as e:print(e)

        if r != DB.ElementId.InvalidElementId:
            try:wert_schreibenbase(self.elem.get_Parameter(self.doc.GetElement(r).GetDefinition()), float(durchmesser)/2)
            except Exception as e:print(e)
    
    def Rotate(self,Richtung):
        if self.Richtung.IsAlmostEqualTo(Richtung) or self.Richtung.IsAlmostEqualTo(Richtung.Negate()):
            angle = 0
        # elif self.Richtung.IsAlmostEqualTo(Richtung.Negate()):
        #     if self.Richtung.Y==self.Richtung.Z==self.Richtung.X:
        #         v0 = self.Richtung + DB.XYZ(0,0,1)
        #     else:
        #         v0 = self.Richtung + DB.XYZ(self.Richtung.Y,self.Richtung.Z,self.Richtung.X)
        #     vector = v0-v0.DotProduct(self.Richtung)*self.Richtung
        #     angle = self.Richtung.AngleTo(Richtung)
        else:
            vector = self.Richtung.CrossProduct(Richtung)
            angle = self.Richtung.AngleTo(Richtung)
        
        if angle > 0:
            DB.ElementTransformUtils.RotateElement(self.doc,self.elem.Id,DB.Line.CreateUnbound(self.Location,vector),angle)
            self.doc.Regenerate()
            if not (self.Richtung.IsAlmostEqualTo(Richtung) or self.Richtung.IsAlmostEqualTo(Richtung.Negate())):
                DB.ElementTransformUtils.RotateElement(self.doc,self.elem.Id,DB.Line.CreateUnbound(self.Location,vector),- 2 * angle)
    
    def Verbinden(self,elem):
        conns = elem.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected == False:
                if conn.CoordinateSystem.BasisZ.IsAlmostEqualTo(self.Richtung.Negate()):
                    conn.Origin = self.Connector.Origin
                    conn.ConnectTo(self.Connector)
                else:
                    conn.Origin = self.Connectors[1].Origin
                    conn.ConnectTo(self.Connectors[1])

class Rohr:
    def __init__(self,elem):
        self.elem = elem
        self._Connectors = None
    
    @property
    def doc(self):
        return self.elem.Document
    
    @property
    def Connectors(self):
        if not self. _Connectors:
            self. _Connectors = list(self.elem.ConnectorManager.Connectors)
        return self._Connectors
    
    @property
    def Connector0(self):
        return self.Connectors[0]
    
    @property
    def Connector1(self):
        return self.Connectors[1]
        
    def get_ClosestPoints(self,Rohr):
        line0 = self.elem.Location.Curve
        line1 = Rohr.Location.Curve
        if line0.Direction.IsAlmostEqualTo(line1.Direction) or line0.Direction.IsAlmostEqualTo(line1.Direction.Negate()):
            print('Die Rohre sind Parallel!')
            return
        a = get_ClosestPoints(line0,line1)
        point = a.XYZPointOnFirstCurve
        line0.Dispose()
        line1.Dispose()
        a.Dispose()
        if point.IsAlmostEqualTo(self.Connector0.Origin) == False and point.IsAlmostEqualTo(self.Connector1.Origin) == False:
            return point
        else:
            print('Der nächstgelegene Punkt ist der Endpunkt des Segments')
    
    def CreateLine(self,Point):
        bearbeitungsbereich = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
        self.doc.GetWorksetTable().SetActiveWorksetId(DB.WorksetId(bearbeitungsbereich))
        elem = self.doc.GetElement(DB.ElementTransformUtils.CopyElement(self.doc,self.elem.Id,DB.XYZ(0,0,0))[0])
        self.doc.Regenerate()
        elem.Location.Curve = DB.Line.CreateBound(Point,self.Connector0.Origin)
        if self.Connector0.IsConnected:
            for ref in self.Connector0.AllRefs:
                if ref.Owner.Category.Name in ['Rohre','Rohrformteile','Rohrzubehör']:
                    if ref.IsConnected:
                        ref.DisconnectFrom(self.Connector0)
                        for conn in elem.ConnectorManager.Connectors:
                            if conn.Origin.IsAlmostEqualTo(self.Connector0.Origin):
                                conn.ConnectTo(self.Connector0)
          
        self.Connector0.Origin = Point
        return elem
    
    def CreateUebergang(self,Typ,Point):
        elem = self.doc.Create.NewFamilyInstance(Point,Typ,self.elem.ReferenceLevel,DB.Structure.StructuralType.NonStructural)
        self.doc.Regenerate()
        tmp = Uebergang(elem)

        tmp.Rotate(self.Connector0.CoordinateSystem.BasisZ)
        self.doc.Regenerate()
        tmp.Move(Point)
        tmp.ChangeDimension(self.Connector0.Radius*304.8*2)
        del tmp
        return elem
    
    def Verbinden(self,Rohr,Ubergang):
        connspaar = []
        for conn in Rohr.ConnectorManager.Connectors:
            if conn.IsConnected == False:
                for conn1 in Ubergang.MEPModel.ConnectorManager.Connectors:
                    if conn1.IsConnected == False:
                        if conn.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn1.CoordinateSystem.BasisZ.Negate()):
                            connspaar.append([conn,conn1])
        if len(connspaar) == 1:
            connspaar[0][0].Origin = connspaar[0][1].Origin
            connspaar[0][0].ConnectTo(connspaar[0][1])
        elif len(connspaar) == 2:
            if connspaar[0][0].Origin.DistanceTo(connspaar[0][1].Origin) < connspaar[1][0].Origin.DistanceTo(connspaar[1][1].Origin):
                connspaar[0][0].Origin = connspaar[0][1].Origin
                connspaar[0][0].ConnectTo(connspaar[0][1])#
            else:
                connspaar[1][0].Origin = connspaar[1][1].Origin
                connspaar[1][0].ConnectTo(connspaar[1][1])

    
    def CreateUebergangAndVerbinden(self,Typ,elem):
        point = self.get_ClosestPoints(elem)
        if point:
            rohr1 = self.CreateLine(point)
            self.doc.Regenerate()
            uebergang = self.CreateUebergang(Typ,point)
            # print(uebergang)
            self.doc.Regenerate()
            self.Verbinden(self.elem,uebergang)
            self.doc.Regenerate()
            self.Verbinden(rohr1,uebergang)
            self.doc.Regenerate()
