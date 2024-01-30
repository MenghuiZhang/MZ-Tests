# coding: utf8
from IGF_Klasse import DB, ItemTemplateMitName, ItemTemplateMitElem
from IGF_Funktionen._Parameter import get_value,wert_schreibenbase
from IGF_Funktionen._Geometrie import CurrentPointTransformToLink

class ElementBase(ItemTemplateMitElem):
    def __init__(self,elem):
        ItemTemplateMitElem.__init__(self, elem)

    @property
    def Bearbeitungsbereich(self):
        try:return self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
        except:return None
    
    @property
    def Muster(self):
        return self.Bearbeitungsbereich in ['KG4xx_Musterbereich','K4xx_Musterbereich']

class FamilieInstanceBase(ElementBase):
    def __init__(self,elem):
        ElementBase.__init__(self, elem)

    @property
    def doc(self):
        return self.elem.Document

    @property
    def CreatedPhase(self):
        try:return self.doc.GetElement(self.elem.CreatedPhaseId)
        except:return None

class FamilienInstance(FamilieInstanceBase):
    def __init__(self,elem,Link = None):
        FamilieInstanceBase.__init__(self, elem)
        self.Link = Link
    
    @property
    def MEP_Raum(self):
        try: return self.elem.Space[self.CreatedPhase]
        except:return None

    @property
    def Linked_MEP_Raum(self):
        try: return self.Link.GetLinkDocument().GetSpaceAtPoint(CurrentPointTransformToLink(self.Link,self.elem.Location.Point))
        except:     return None

    
    @property
    def Linked_MEP_Raum_Number(self):
        try: return self.Linked_MEP_Raum.Number
        except:return None
    
    @property
    def Linked_MEP_Raum_Name(self):
        try: self.Linked_MEP_Raum.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        except:return None

    @property
    def MEP_Raum_Id(self):
        if self.MEP_Raum:return self.MEP_Raum.Id.IntegerValue
        else:return -1

    @property
    def MEP_Raum_Number(self):
        if self.MEP_Raum:return self.MEP_Raum.Number
        else:return None
    
    @property
    def MEP_Raum_Name(self):
        if self.MEP_Raum:return self.MEP_Raum.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        else:return None
    
    @property
    def MEP_Raum_Name_Number(self):
        if self.MEP_Raum:return self.MEP_Raum_Number + ' - ' + self.MEP_Raum_Name
        else:return None
    
    def wert_schreiben(self,paramname,wert):
        param = self.elem.LookupParameter(paramname)
        wert_schreibenbase(param,wert)
 
    def Schreiben_Einbauort(self):
        try:self.wert_schreiben('IGF_X_Einbauort', self.MEP_Raum_Number)
        except:pass
    
    def Schreiben_Linked_Einbauort(self):
        try:self.wert_schreiben('IGF_X_Einbauort', self.Linked_MEP_Raum_Number)
        except:pass


class Brandschutzklappe(FamilieInstanceBase):
    def __init__(self, elem):
        FamilieInstanceBase.__init__(self, elem)


    @property
    def Connectors(self):

        liste = []
        for conn in self.elem.MEPModel.ConnectorManager.Connectors:
            if conn.Domain.ToString() == 'DomainHvac':
                liste.append(conn)

        return liste

    @property
    def BedienConnector(self):


            for conn in self.Connectors:
                if conn.Description == 'Bedienseite':
                    return conn


    @property
    def Connector(self):
        for conn in self.Connectors:
            if conn.Description != 'Bedienseite':
                return conn

    @property
    def Seite(self):

        if self.Connector:
            richtung = self.Connector.CoordinateSystem.BasisZ
            point = self.Connector.Origin
            n = 0

            _space = self.doc.GetSpaceAtPoint(point, self.CreatedPhase)
            while (_space is None and n < 5):
                if _space:
                    break
                if n == 5:
                    break
                n += 1
                point = point + richtung * n/2.5

                _space = self.doc.GetSpaceAtPoint(point, self.CreatedPhase)
            return _space



    @property
    def Bedienseite(self):

        if self.BedienConnector:
            richtung = self.BedienConnector.CoordinateSystem.BasisZ
            point = self.BedienConnector.Origin
            n = 0

            _space = self.doc.GetSpaceAtPoint(point, self.CreatedPhase)
            while (_space is None and n < 5):
                if _space:
                    break
                if n == 5:
                    break
                n += 1
                point = point + richtung * n/2.5

                _space = self.doc.GetSpaceAtPoint(point, self.CreatedPhase)
            return _space



    def wert_schreiben(self,paramname,wert):
        param = self.elem.LookupParameter(paramname)
        if param:
            if wert is not None:
                if param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                elif param.StorageType.ToString() == 'Integer':
                    param.Set(int(wert))
                elif param.StorageType.ToString() == 'String':
                    param.Set(str(wert))
                else:
                    param.Set(wert)

    def Ortschreiben(self):
        try:self.wert_schreiben('IGF_X_Einbauort', self.Bedienseite.Number)
        except:pass
        try:self.wert_schreiben('IGF_X_Wirkungsort', self.Seite.Number)
        except:pass


class Brandschutzventile(FamilieInstanceBase):
    def __init__(self, elem, Tellerventil = True):
        FamilieInstanceBase.__init__(self, elem)
        self.Tellerventil = Tellerventil

    @property
    def Connectors(self):

        liste = []
        for conn in self.elem.MEPModel.ConnectorManager.Connectors:
            if conn.Domain.ToString() == 'DomainHvac':
                liste.append(conn)

        return liste



    @property
    def Connector(self):
        if not self.Connectors:
            return self.Connectors[0]


    @property
    def Seite(self):


        return self.elem.Space[self.CreatedPhase]



    @property
    def Bedienseite(self):

        if self.Tellerventil:
            if self.Connector:

                richtung = self.Connector.CoordinateSystem.BasisZ
                point = self.Connector.Origin
                n = 0

                _space = self.doc.GetSpaceAtPoint(point, self.CreatedPhase)
                while (_space is None and n < 5):
                    if _space:
                        break
                    if n == 5:
                        break
                    n += 1
                    point = point + richtung * n / 2.5
                    _space = self.doc.GetSpaceAtPoint(point, self.CreatedPhase)
                return _space
        else:
            return self.Seite


    def wert_schreiben(self, paramname, wert):
        param = self.elem.LookupParameter(paramname)
        if param:
            if wert is not None:
                if param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                elif param.StorageType.ToString() == 'Integer':
                    param.Set(int(wert))
                elif param.StorageType.ToString() == 'String':
                    param.Set(str(wert))
                else:
                    param.Set(wert)

    def Ortschreiben(self):
        try:
            self.wert_schreiben('IGF_X_Einbauort', self.Bedienseite.Number)
        except:
            pass
        try:
            self.wert_schreiben('IGF_X_Wirkungsort', self.Seite.Number)
        except:
            pass







  
