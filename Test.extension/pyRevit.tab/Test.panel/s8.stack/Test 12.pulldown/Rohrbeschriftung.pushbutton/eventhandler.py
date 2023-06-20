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
        conn1.Origin = conn0.Origin
        conn0.ConnectTo(conn1)
        
    def verschieben(self,conn0,conn1,elem1):
        vector = conn0.Origin-conn1.Origin
        length = abs(vector.DotProduct(conn0.CoordinateSystem.BasisZ))
        point = conn0.Origin + conn0.CoordinateSystem.BasisZ * length
        elem1.Location.Move(point - conn1.Origin)

    def CREATE(self,uiapp):
        while (True):
            try:
                uidoc = uiapp.ActiveUIDocument
                doc = uidoc.Document
                ref1 = uidoc.Selection.PickObject(Selection.ObjectType.Element)
                ref2 = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter1(ref1.ElementId.ToString()))
                elem1 = doc.GetElement(ref1.ElementId)
                elem2 = doc.GetElement(ref2.ElementId)
            
                conn0,conn1 = self.GetparterConnector(elem1,elem2)
                t = DB.Transaction(doc,self.name)
                t.Start()
                try:
                    if self.conns_pruefen_0(conn0, conn1):
                        if self.conns_pruefen(conn0, conn1):
                            self.connect(conn0, conn1)
                        else:
                            self.verschieben(conn0, conn1, elem2)
                            doc.Regenerate()
                            self.connect(conn0, conn1)
                    else:
                        print(0)
                except:
                    t.RollBack()
                    break
                t.Commit()

                t.Dispose()

                
            except Exception as e:
                print(e)
                break
    
