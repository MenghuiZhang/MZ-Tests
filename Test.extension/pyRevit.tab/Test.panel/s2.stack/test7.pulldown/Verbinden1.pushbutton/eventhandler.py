# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB

class HLSBauteile(Selection.ISelectionFilter):
    # Formteil
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2001140':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Rohrzubehoer(Selection.ISelectionFilter):
    # Formteil
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008055':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter1(Selection.ISelectionFilter):
    def __init__(self,elemid):
        self.elemid = elemid
    # Formteil
    def AllowElement(self,element):
        if element.Id.ToString() != self.elemid:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class ExternalEventListe(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.name = 'verbinden'
        self.executeapp = self.CREATE
    
    def GetName(self):
        return self.name
    
    def Execute(self,uiapp):
        try:
            self.executeapp(uiapp)
        except:
            pass
    
    def GetparterConnector(self,elem1,elem2):
        try:conns0 = elem1.ConnectorManager.Connectors
        except:conns0 = elem1.MEPModel.ConnectorManager.Connectors
        try:conns1 = elem2.ConnectorManager.Connectors
        except:conns1 = elem2.MEPModel.ConnectorManager.Connectors

        Distance = 10000
        conn_0 = None
        conn_1 = None
        for conn1 in conns1:
            if not conn1.IsConnected:
                for conn0 in conns0:
                    if not conn0.IsConnected:
                        distance = conn0.Origin.DistanceTo(conn1.Origin)
                        if distance < Distance:
                            Distance = distance
                            conn_0 = conn0
                            conn_1 = conn1

        return conn_0,conn_1
    
    def conns_pruefen(self,conn0,conn1):
        vector = conn0.Origin-conn1.Origin
        if vector.Normalize().IsAlmostEqualTo(conn0.CoordinateSystem.BasisZ) or  vector.Normalize().IsAlmostEqualTo(conn0.CoordinateSystem.BasisZ.Negate()):
            return True
        else:
            return False
    
    def conns_pruefen_0(self,conn0,conn1):
        if conn0.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn1.CoordinateSystem.BasisZ.Negate()):
            return True
        else:
            return False
    
    def connect(self,conn0,conn1):
        # conn1.Origin = conn0.Origin
        conn0.ConnectTo(conn1)
        
    def verschieben(self,conn0,conn1,elem1):
        # vector = conn0.Origin-conn1.Origin
        # length = abs(vector.DotProduct(conn0.CoordinateSystem.BasisZ))
        # point = conn0.Origin + conn0.CoordinateSystem.BasisZ * length
        elem1.Location.Move(conn0.Origin-conn1.Origin)

    def CREATE(self,uiapp):
        n = -1
        raum = None
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        while (True):
            try:
                ref1 = uidoc.Selection.PickObject(Selection.ObjectType.Element)
                n += 1
                elem1 = doc.GetElement(ref1.ElementId)
                space = elem1.Space[doc.GetElement(elem1.CreatedPhaseId)].Number
                if space != raum:
                    n = 0
                    raum = space
            
                t = DB.Transaction(doc,self.name)
                t.Start()
                elem1.LookupParameter('SBI_Bauteilnummerierung').Set('DSP_'+raum+'.'+str(n))
                t.Commit()

                t.Dispose()

                
            except Exception as e:
                return
    
