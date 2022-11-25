# coding: utf8
import Autodesk.Revit.DB as DB

__title__ = "ULK"
__doc__ = """"""
__author__ = "Menghui Zhang"


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document 
active_view = uidoc.ActiveView

def Element_TypFilter(Fam_name = None, Typname = None):

    param_equality=DB.FilterStringContains()

    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)

    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)

    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,Fam_name,True)

    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    param_equality1=DB.FilterStringEquals()
    Fam_typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)

    Fam_typ_prov = DB.ParameterValueProvider(Fam_typ_id)

    Fam_typ_value_rule = DB.FilterStringRule(Fam_typ_prov, param_equality1, Typname, True)

    Fam_typ_filter = DB.ElementParameterFilter(Fam_typ_value_rule)


    coll = DB.FilteredElementCollector(doc).WherePasses(Fam_name_filter).WherePasses(Fam_typ_filter).WhereElementIsNotElementType().ToElements()

    return coll

Bauteile = Element_TypFilter('MC_K_IGF_FG_Flex Geko','BG8')

class Umluft:
    def __init__(self,elem):
        self.elem = elem
        self.rohrliste = []
        self.conns0 = []
        self.conns1 = []
    def get_rohr_liste(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.PipeSystemType.ToString() == 'ReturnHydronic':
                self.conns0.append(conn)
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner
                    if owner.Category.Name == 'Rohre':
                        self.rohrliste.append(owner)
                        conns_rohr = owner.ConnectorManager.Connectors
                        for conn_r in conns_rohr:
                            if conn_r.IsConnectedTo(conn):
                                self.conns0.append(conn_r)
                                break
                        break
            elif conn.PipeSystemType.ToString() == 'SupplyHydronic':
                self.conns1.append(conn)
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner
                    if owner.Category.Name == 'Rohre':
                        self.rohrliste.append(owner)
                        conns_rohr = owner.ConnectorManager.Connectors
                        for conn_r in conns_rohr:
                            if conn_r.IsConnectedTo(conn):
                                self.conns1.append(conn_r)
                                break
                        break
    def disconnect(self):
        try:self.conns0[0].DisconnectFrom(self.conns0[1])
        except:print(self.elem.Id.ToString())
        try:self.conns1[0].DisconnectFrom(self.conns1[1])
        except:print(self.elem.Id.ToString())
                        
class Rohr:
    def __init__(self,elem):
        self.elem = elem
        self.bogen = ''
        self.location = ''
        self.conns0 = []
        self.conn = ''
        self.conn1 = ''
        self.conn1_location = ''
        self.conn1_temp = ''

    def get_rohr_liste(self):
        conns = self.elem.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner
                    if owner.Category.Name == 'Rohrformteile':
                        self.conns0.append(conn)
                        self.conn = conn
                        self.bogen = owner
                        conns_rohr = owner.MEPModel.ConnectorManager.Connectors
                        for conn_r in conns_rohr:
                            if conn_r.IsConnectedTo(conn):
                                self.conns0.append(conn_r)
        # for conn in conns:
        #     if conn.Id != self.conn.Id:
        #         self.conn1 = conn
        #         self.conn1_location = self.conn1.Origin
        # l0 = DB.Line.CreateBound(self.conn.Origin,self.conn1.Origin)
        # ln = l0.Direction.Normalize()
        # self.conn1_temp = self.conn.Origin + ln*2000/304.8
            
    def dimension(self):
        self.elem.LookupParameter('Durchmesser').SetValueString('20')
        self.bogen.LookupParameter('Nennradius').SetValueString('10')
        # doc.Regenerate()
        # try:self.bogen.LookupParameter('Nennradius').SetValueString('10')
        # except:print(self.bogen.Id.ToString())
        # doc.Regenerate()
    
    # def verschieben0(self):
    #     self.conn1.Origin = self.conn1_temp
    #     doc.Regenerate()
    # def verschieben1(self):
    #     try:self.conn1.Origin = self.conn1_location
    #     except Exception as e:print(e)
    #     doc.Regenerate()

bauteil_liste = []
rohr_liste = []
for el in Bauteile:
    ulk = Umluft(el)
    ulk.get_rohr_liste()
    bauteil_liste.append(ulk)
    for r in ulk.rohrliste:
        rohr = Rohr(r)
        rohr.get_rohr_liste()
        rohr_liste.append(rohr)

t = DB.Transaction(doc,'ULK & Rohr')
t.Start()
for el in bauteil_liste:
    el.disconnect()
doc.Regenerate()
t.Commit()

t1 = DB.Transaction(doc,'Rohr & Rohrformteil')
t1.Start()
for el in rohr_liste:
    # rohr.verschieben0()
    el.dimension()
#     rohr.verschieben1()
# doc.Regenerate()
t1.Commit()