# coding: utf8
from System import Enum
from IGF_Klasse import DB, ItemTemplateNurElem
from IGF_Funktionen._Parameter import get_value
from System.Collections.ObjectModel import ObservableCollection
import clr

class TrassenType(ItemTemplateNurElem):
    def __init__(self, elem):
        ItemTemplateNurElem.__init__(self, elem)
        self.name = self.elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        self.familiename = self.elem.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
        self.familietyp = self.familiename + ': ' + self.name

class Trassen(ItemTemplateNurElem):
    def __init__(self, elem,Anzahl = 1):
        ItemTemplateNurElem.__init__(self, elem)
        self._Connectors = None
        self._Shape = None
        self._StartConnector = None
        self._EndConnector = None
        self.BeschrifterAnzahl = Anzahl
    
    @property
    def Connectors(self):
        if not self._Connectors: self._Connectors =  self.elem.ConnectorManager.Connectors
        return self._Connectors

    @property
    def StartPoint(self):
        return self.elem.Location.Curve.GetEndPoint(0)
    
    @property
    def EndPoint(self):
        return self.elem.Location.Curve.GetEndPoint(1)

    @property
    def MittePoint(self):
        return (self.StartPoint + self.EndPoint) / 2
    
    @property
    def StartConnector(self):
        if not self._StartConnector:
            liste = list(self.Connectors)
            conn0 = liste[0]
            conn1 = liste[1]
            if conn0.Origin.IsAlmostEqualTo(self.StartPoint):
                self._StartConnector = conn0
                self._EndConnector = conn1
            else:
                self._StartConnector = conn1
                self._EndConnector = conn0
        return self._StartConnector

    @property
    def EndConnector(self):
        if not self._EndConnector:
            liste = list(self.Connectors)
            conn0 = liste[0]
            conn1 = liste[1]
            if conn0.Origin.IsAlmostEqualTo(self.StartPoint):
                self._StartConnector = conn0
                self._EndConnector = conn1
            else:
                self._StartConnector = conn1
                self._EndConnector = conn0
        return self._EndConnector
    
    @property
    def Connector(self):
        return list(self.Connectors)[0]
    
    @property
    def Shape(self):
        if not self._Shape:
            if self.Connector.Shape.ToString() == 'Round':
                self._Shape = 'RU'
            else:
                self._Shape = 'RE'
        return self._Shape
    
    @property
    def Durchmesser(self):
        if self.Shape == 'RU':
            return int(round(self.Connector.Radius * 304.8 * 2))
    
    @property
    def Breite(self):
        if self.Shape == 'RE':
            return int(round(self.Connector.Width * 304.8))
    
    @property
    def Hoehe(self):
        if self.Shape == 'RE':
            return int(round(self.Connector.Height * 304.8))
    
    @property
    def StartLinkedElement(self):
        if self.StartConnector.IsConnected:
            refs = self.StartConnector.AllRefs
            for ref in refs:
                if ref.Owner.Category.Id.ToString() in ['-2008010','-2008049']:
                    return Formteile(ref.Owner)
                elif ref.Owner.Category.Id.ToString() in ['-2008016','-2008055']:
                    return Zubehoer(ref.Owner)
        return None
    @property
    def EndLinkedElement(self):
        if self.EndConnector.IsConnected:
            refs = self.EndConnector.AllRefs
            for ref in refs:
                if ref.Owner.Category.Id.ToString() in ['-2008010','-2008049']:
                    return Formteile(ref.Owner)
                elif ref.Owner.Category.Id.ToString() in ['-2008016','-2008055']:
                    return Zubehoer(ref.Owner)
        return None
    
    def CreatedBeschriftung_Start(self,Beschrifter,View,Basisz = None):
        if Basisz:
            DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, DB.XYZ(self.StartPoint.X,self.StartPoint.Y,Basisz))
        else:
            DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, self.StartPoint)
    
    def CreatedBeschriftung_End(self,Beschrifter,View,Basisz = None):
        if Basisz:
            DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, DB.XYZ(self.EndPoint.X,self.EndPoint.Y,Basisz))
        else:
            DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, self.EndPoint)
    
    def CreatedBeschriftung_Mitte(self,Beschrifter,View,Basisz = None):
        if Basisz:
            DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, DB.XYZ(self.MittePoint.X,self.MittePoint.Y,Basisz))
        else:
            DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, self.MittePoint)
    
    def CreatedBeschriftung_Point(self,Beschrifter,View,Point,Linie = False):
        if Linie:
            tag = DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, Point)
            self.elem.Document.Regenerate()
            tag.LeadersPresentationMode  = DB.LeadersPresentationMode.ShowAll
        else:DB.IndependentTag.Create(self.elem.Document, Beschrifter, View.Id,  DB.Reference(self.elem), False, DB.TagOrientation.Horizontal, Point)

class Zubehoer(ItemTemplateNurElem):
    def __init__(self, elem):
        ItemTemplateNurElem.__init__(self, elem)
        self.elemid = self.elem.Id.ToString()
        self._Connectors = None
        self._Shape = None
        self.Art = 'Zubehör'
    
    @property
    def Connectors(self):
        if not self._Connectors: self._Connectors =  self.elem.MEPModel.ConnectorManager.Connectors
        return self._Connectors
    
    @property
    def Connector(self):
        return list(self.Connectors)[0]
    
    @property
    def Shape(self):
        if not self._Shape:
            if self.Connector.Shape.ToString() == 'Round':
                self._Shape = 'RU'
            else:
                self._Shape = 'RE'
        return self._Shape
      
class Formteile(ItemTemplateNurElem):
    def __init__(self, elem):
        ItemTemplateNurElem.__init__(self, elem)
        self.elemid = self.elem.Id.ToString()
        self._Connectors = None
        self._Shape = None
        self.Art = 'Formteil'
    
    @property
    def Connectors(self):
        if not self._Connectors: self._Connectors =  self.elem.MEPModel.ConnectorManager.Connectors
        return self._Connectors
    
    @property
    def Connector(self):
        return list(self.Connectors)[0]
    
    @property
    def Shape(self):
        if not self._Shape:
            if self.Connector.Shape.ToString() == 'Round':
                self._Shape = 'RU'
            else:
                self._Shape = 'RE'
        return self._Shape
    
    @property
    def IstBogen(self):
        return self.elem.MEPModel.PartType.ToString() == 'Elbow'

    @property
    def IstAsymBogen(self):
        if self.IstBogen:
            liste_conns = list(self.Connectors)
            conn0 = liste_conns[0]
            conn1 = liste_conns[1]
            try:
                if self.Shape == 'RE':
                    if conn0.Width == conn1.Width and conn0.Height == conn1.Height:
                        return False
                    else:
                        return True
                else:
                    if conn0.Radius == conn1.Radius:
                        return False
                    else:
                        return True
            except:
                return False
        return   

class TrassenFactory(object):
    uiapp = __revit__
    @staticmethod
    def GetAllLuftkanalType():
        try:
            coll = DB.FilteredElementCollector(TrassenFactory.uiapp.ActiveUIDocument.Document).OfClass(DB.Mechanical.DuctType).WhereElementIsElementType().ToElements()
            liste = [TrassenType(el) for el in coll]
            liste.sort(key= lambda x:x.familietyp)
        except:
            liste = []
        return ObservableCollection[TrassenType](liste)

    @staticmethod
    def GetAllRohrType():
        try:
            coll = DB.FilteredElementCollector(TrassenFactory.uiapp.ActiveUIDocument.Document).OfClass(DB.Plumbing.PipeType).WhereElementIsElementType().ToElements()
            liste = [TrassenType(el) for el in coll]
            liste.sort(key= lambda x:x.familietyp)
        except:
            liste = []
        return ObservableCollection[TrassenType](liste)