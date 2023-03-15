# coding: utf8
from rpw import revit,DB
from pyrevit import script, forms
from System.Collections.Generic import List



__title__ = "Strömungsteiler entfernen"
__doc__ = """

[2022.11.30]
Version: 1.1
"""
__authors__ = "Menghui Zhang"


logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc

def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if not param:return ''
    if param.StorageType.ToString() == 'ElementId':
        return param.AsValueString()
    elif param.StorageType.ToString() == 'Integer':
        value = param.AsInteger()
    elif param.StorageType.ToString() == 'Double':
        value = param.AsDouble()
    elif param.StorageType.ToString() == 'String':
        value = param.AsString()
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value

class Verteiler:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.IsVerteiler = False
        self.liste = []
        try:
            if self.elem.Symbol.FamilyName != '_S_IGF_412_Strömungsteiler':
                return
            self.IsVerteiler = True

        except:
            return
        
        self.rohr = None
        self.bogen = None
        self.conn_ein0 = None
        self.conn_ein1 = None
        self.conn_aus0 = None
        self.conn_aus1 = None
        self.rohr_ein0 = None
        self.rohr_ein1 = None
        self.rohr_aus0 = None
        self.rohr_aus1 = None
        self.liste0 = []
    
    def get_liste0(self):
        for conn in self.elem.MEPModel.ConnectorManager.Connectors:
            if conn.IsConnected:
                refs = conn.AllRefs
                for ref in refs:
                    if ref.Owner.Category.Name == 'Rohrformteile':
                        self.liste0.append(ref.Owner.Id)

    def getconns(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        conn_ein0 = None
        conn_ein1 = None
        conn_aus0 = None
        conn_aus1 = None
        for conn in conns:
            if conn.IsConnected:
                if conn.Direction.ToString() == 'Out':
                    if not conn_aus0: conn_aus0 = conn
                    else:conn_aus1 = conn
                else:
                    if not conn_ein1: conn_ein1 = conn
                    else:conn_ein0 = conn
        distance0 = conn_ein1.Origin.DistanceTo(conn_aus0.Origin)
        distance1 = conn_ein1.Origin.DistanceTo(conn_aus1.Origin)
        distance2 = conn_ein0.Origin.DistanceTo(conn_aus0.Origin)
        distance3 = conn_ein0.Origin.DistanceTo(conn_aus1.Origin)
        if distance0 == max(distance0,distance1,distance2,distance3):
            self.conn_ein0 = conn_ein1
            self.conn_ein1 = conn_ein0
            self.conn_aus0 = conn_aus0
            self.conn_aus1 = conn_aus1
        elif distance1 == max(distance0,distance1,distance2,distance3):
            self.conn_ein0 = conn_ein1
            self.conn_ein1 = conn_ein0
            self.conn_aus0 = conn_aus1
            self.conn_aus1 = conn_aus0
        elif distance2 == max(distance0,distance1,distance2,distance3):
            self.conn_ein0 = conn_ein0
            self.conn_ein1 = conn_ein1
            self.conn_aus0 = conn_aus0
            self.conn_aus1 = conn_aus1
        elif distance3 == max(distance0,distance1,distance2,distance3):
            self.conn_ein0 = conn_ein0
            self.conn_ein1 = conn_ein1
            self.conn_aus0 = conn_aus1
            self.conn_aus1 = conn_aus0
        
    def get_Liste(self,elem):
        elemid = elem.Id.IntegerValue
        self.liste.append(elemid)
        if len(self.liste) > 200:return
        try:
            conns = elem.MEPModel.ConnectorManager.Connectors
        except:
            conns = elem.ConnectorManager.Connectors
        if conns:
            for conn in conns:
                if elemid == self.elemid.IntegerValue:
                    if conn.Id != self.conn_ein1.Id:
                        continue
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner
                    if owner.Category.Name in ['Rohre','Rohrformteile','Rohrzubehör']:
                        if owner.Id.IntegerValue not in self.liste:
                            self.get_Liste(owner)
    
    def getrohr(self):
        def getrohrintener(conn):
            refs = conn.AllRefs
            for ref in refs:
                owner = ref.Owner
                if owner.Category.Name == 'Rohre':
                    return owner
                elif owner.Category.Name == 'Rohrformteile':
                    conns = owner.MEPModel.ConnectorManager.Connectors
                    for conn1 in conns:
                        if conn1.IsConnectedTo(conn):
                            continue
                        else:
                            return getrohrintener(conn1)


        self.rohr_ein0 = getrohrintener(self.conn_ein0)
        self.rohr_ein1 = getrohrintener(self.conn_ein1)
        self.rohr_aus0 = getrohrintener(self.conn_aus0)
        self.rohr_aus1 = getrohrintener(self.conn_aus1)
    
    def getrohr1(self):
        conns0 = list(self.rohr_ein0.ConnectorManager.Connectors)
        conns3 = list(self.rohr_ein1.ConnectorManager.Connectors)
        conns2 = list(self.rohr_aus0.ConnectorManager.Connectors)
        conns1 = list(self.rohr_aus1.ConnectorManager.Connectors)
                    
        distance = 10000000

        for con in conns0:
            for con1 in conns1:
                dis = con.Origin.DistanceTo(con1.Origin)
                if dis < distance:
                    distance = dis
                    self.conn_ein0 = con
                    self.conn_aus1 = con1

        for con in conns3:
            for con1 in conns2:
             
                dis = con.Origin.DistanceTo(con1.Origin)
                if dis < distance:
                    distance = dis
                    self.conn_ein1 = con
                    self.conn_aus0 = con1
    
    def verbinden(self):
        # print(self.conn_ein0.Owner.Name,self.conn_aus1.Owner.Name)
        # print(self.conn_ein1.Owner.Name,self.conn_aus0.Owner.Name)
        try:
            doc.Regenerate()
            doc.Create.NewElbowFitting(self.conn_ein0, self.conn_aus1)
            doc.Regenerate()
        except:
            try:
                doc.Regenerate()
                doc.Create.NewElbowFitting(self.conn_aus1, self.conn_ein0)
                doc.Regenerate()
            except:
                print('Manuell: Rohr 1: {}, Rohr 2: {}'.format(self.rohr_ein0.Id.ToString(),self.rohr_aus1.Id.ToString()))

        try:
            doc.Regenerate()
            doc.Create.NewElbowFitting(self.conn_ein1, self.conn_aus0)
            doc.Regenerate()
        except:
            try:
                doc.Regenerate()
                doc.Create.NewElbowFitting(self.conn_aus0, self.conn_ein1)
                doc.Regenerate()
            except:
                print('Manuell: Rohr 1: {}, Rohr 2: {}'.format(self.rohr_ein1.Id.ToString(),self.rohr_aus0.Id.ToString()))



class Element:
    def __init__(self,elemid,dimension):
        self.elemid = elemid
        self.elem = doc.GetElement(DB.ElementId(self.elemid))
        self.isrohr = (self.elem.Category.Name == 'Rohre')
        self.dimension = dimension
    def Dimensionieren_Formteile(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            mepinfo = conn.GetMEPConnectorInfo()
            d = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_DIAMETER))
            r = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_RADIUS))
            if d != DB.ElementId.InvalidElementId:
                try:
                    param = doc.GetElement(d)
                    self.elem.get_Parameter(param.GetDefinition()).SetValueString(str(self.dimension))
                except:pass
            if r != DB.ElementId.InvalidElementId:
                try:
                    param = doc.GetElement(r)
                    self.elem.get_Parameter(param.GetDefinition()).SetValueString(str(self.dimension/2.0))
                except:pass
    def Dimensionieren_Rohr(self):
        self.elem.LookupParameter('Durchmesser').SetValueString(str(self.dimension))  
    def dimensionieren(self):
        if self.isrohr:
            self.Dimensionieren_Rohr()
        else:
            self.Dimensionieren_Formteile()



# Liste = uidoc.Selection.GetElementIds()
Liste = []
Liste_elemid = []
for el in uidoc.Selection.GetElementIds():
    fs = Verteiler(el)
    if fs.IsVerteiler:
        Liste.append(fs)

t = DB.Transaction(doc,'Verteiler')
t.Start()
# doc.Delete(List[DB.ElementId](Liste_elemid))
# doc.Regenerate()

with forms.ProgressBar(title="{value}/{max_value} Verteiler",cancellable=True, step=1) as pb:

    for n,fs in enumerate(Liste):
        if pb.cancelled:
            t.RollBack()
            script.exit()
        
        pb.update_progress(n+1, len(Liste))
        fs.getconns()
        fs.get_Liste(fs.elem)
        fs.getrohr()

        dim = get_value(fs.rohr_ein0.LookupParameter('Durchmesser'))
        for el in fs.liste:
            elem = Element(el,dim)
            try:elem.dimensionieren()
            except Exception as e:print(e)
            doc.Regenerate()
        fs.getrohr()
        
        fs.get_liste0()
        doc.Delete(fs.elemid)
        doc.Delete(List[DB.ElementId](fs.liste0))
        doc.Regenerate()
        fs.getrohr1()
        fs.verbinden()


t.Commit()
