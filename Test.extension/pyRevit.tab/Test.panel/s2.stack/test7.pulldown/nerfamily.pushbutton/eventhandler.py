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
        if element.Category.Id.ToString() == '-2008044':
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
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        ref = uidoc.Selection.PickObject(Selection.ObjectType.PointOnElement)#,Rohrzubehoer())
        rohr = doc.GetElement(ref.ElementId)
        symbol = doc.GetElement(DB.ElementId(31678762))
        point = ref.GlobalPoint
        t = DB.Transaction(doc,'Klappe')
        t.Start()
        doc.Create.NewFamilyInstance(point,symbol,rohr,DB.Structure.StructuralType.NonStructural)
        t.Commit()
        t.Dispose()
    

cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id)
t = DB.Transaction(doc,'1')
t.Start()
for el in cl:
    el.LookupParameter('IGF_X_Bauteil_ID_Text').Set(el.LookupParameter('SBI_Bauteilnummerierung').AsString())
t.Commit()


for el in cl:
    conns = el.MEPModel.ConnectorManager.Connectors
    for conn in conns:
        if conn.IsConnected:
            for ref in conn.AllRefs:
                if ref.Owner.Category.Name == 'Rohrformteile':
                    
                    print(el.Id,ref.Owner.Name)

_dict = {}
for el in cl:
    if el.Space[doc.GetElement(el.CreatedPhaseId)]:
        nummer = el.Space[doc.GetElement(el.CreatedPhaseId)].Number
        _dict.setdefault(nummer,0)
        _dict[nummer] += 1
for el in sorted(_dict.keys()):
    print(el,_dict[el])









class Rohrformteil:
    def __init__(self,elem):
        self.doc = None
        self.elem = elem
        self.conns0 = ''
        self.conns1 = ''
        self.conns2 = ''
        self.liste = []
        self.art = ''
        # self.get_conns()
    
    def get_conns(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        if conns.Size == 2:
            conns = list(conns)
            conn_sizes = [int(conn.Radius*304.8*2) for conn in conns]
            if conn_sizes[0] != conn_sizes[1]:
                self.art = 'Übergang'
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Category.Id.ToString() in ['-2008044','-2008055','-2001140','-2001160']: # Rohre, Zubehör, HLS, Sanitärinstallationen
                                try:
                                    conns_owner = owner.ConnectorManager.Connectors
                                except:
                                    conns_owner = owner.MEPModel.ConnectorManager.Connectors
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp):
                                        if self.conns0 == '':
                                            self.conns0 = conn_temp
                                        else:
                                            self.conns1 = conn_temp
            else:
                self.art = 'Bogen'
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            elemid = owner.Id.ToString()
                            if elemid in self.liste:
                                continue
                            self.liste.append(elemid)
                            if owner.Category.Id.ToString() == '-2008044':
                                conns_owner = owner.ConnectorManager.Connectors
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp):
                                        if self.conns0 == '':
                                            self.conns0 = conn_temp
                                        else:
                                            self.conns1 = conn_temp
                            elif owner.Category.Id.ToString() == '-2008049':
                                conns_owner = owner.MEPModel.ConnectorManager.Connectors
                                conn_ander = ''
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp) == False:
                                        conn_ander = conn_temp
                                        break 
                                allrefs2 = conn_ander.AllRefs
                                for ref2 in allrefs2:
                                    owner2 = ref2.Owner
                                    if owner2.Category.Id.ToString() == '-2008044':
                                        conns_owner2 = owner2.ConnectorManager.Connectors
                                        for conn_temp2 in conns_owner2:
                                            if conn_ander.IsConnectedTo(conn_temp2):
                                                if self.conns0 == '':
                                                    self.conns0 = conn_temp2
                                                else:
                                                    self.conns1 = conn_temp2
        else:
            self.art = 'T-Stück'
            for conn in conns:
                if conn.IsConnected:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        elemid = owner.Id.ToString()
                        if elemid in self.liste:
                            continue
                        self.liste.append(elemid)
                        if owner.Category.Id.ToString() == '-2008044':
                            conns_owner = owner.ConnectorManager.Connectors
                            for conn_temp in conns_owner:
                                if conn.IsConnectedTo(conn_temp):
                                    if self.conns0 == '':
                                        self.conns0 = conn_temp
                                    elif self.conns1 == '':
                                        self.conns1 = conn_temp
                                    else:
                                        self.conns2 = conn_temp
                        elif owner.Category.Id.ToString() == '-2008049':
                            conns_owner = owner.MEPModel.ConnectorManager.Connectors
                            if conns_owner.Size == 1:
                                for conn_temp in conns_owner:
                                    if self.conns0 == '':
                                        self.conns0 = conn_temp
                                    elif self.conns1 == '':
                                        self.conns1 = conn_temp
                                    else:
                                        self.conns2 = conn_temp
                            else:
                                conn_ander = ''
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp) == False:
                                        conn_ander = conn_temp
                                        break 
                                allrefs2 = conn_ander.AllRefs
                                for ref2 in allrefs2:
                                    owner2 = ref2.Owner
                                    if owner2.Category.Id.ToString() == '-2008044':
                                        conns_owner2 = owner2.ConnectorManager.Connectors
                                        for conn_temp2 in conns_owner2:
                                            if conn_ander.IsConnectedTo(conn_temp2):
                                                if self.conns0 == '':
                                                    self.conns0 = conn_temp2
                                                elif self.conns1 == '':
                                                    self.conns1 = conn_temp2
                                                else:
                                                    self.conns2 = conn_temp2


