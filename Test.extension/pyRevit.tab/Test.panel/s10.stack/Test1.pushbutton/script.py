# coding: utf8
from IGF_log import getlog,getloglocal
from System.Collections.Generic import List
from rpw import revit,DB


__title__ = "Dimension VE"
__doc__ = """

Dimension von VE
"""

__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp
active_view = uidoc.ActiveView
revitversion = app.VersionNumber


coll_temp = DB.FilteredElementCollector(doc,active_view.Id).OfCategory(DB.BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()
coll = []
for elem in coll_temp:
    if elem.Name == 'VE (kalt)' and elem.Symbol.FamilyName == '_S_IGF_xxx_MC_Laborzeile':
        coll.append(elem)


class Verbraucher:
    def __init__(self,elem):
        self.elem = elem
        self.pipes = []
        self.Liste = []
        self.bogen = []
        self.Liste_Hidden = List[DB.ElementId]()
        self.firstt = None
        self.vorlauf = None
        self.ruecklauf = None
        self.get_first_T(elem)
        if self.firstt:
            self.get_start_p(self.firstt)
        else:
            
            print('{}: T-Stück nicht gefunden'.format(self.elem.Id.IntegerValue))
            return
        if self.vorlauf and self.ruecklauf:
            self.get_elems(self.vorlauf)
            self.get_elems(self.ruecklauf)
        else:
            print('{}: Rohrzubehör nicht gefunden'.format(self.elem.Id.IntegerValue))
            return
        
        self.Liste = set(self.Liste)
        self.Liste = list(self.Liste)
        for Id in self.Liste:
            self.Liste_Hidden.Add(DB.ElementId(Id))
        self.dimension()

    
    def get_first_T(self,elem):
        Id = elem.Id.IntegerValue
        self.Liste.append(Id)
        if len(self.Liste) > 6:
            return
        if self.firstt:return
        cate = elem.Category.Name
        
        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2:
                    self.firstt = elem
                    return
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.IntegerValue not in self.Liste:                                                                        
                            self.get_first_T(owner)

    def get_start_p(self,elem):
        Id = elem.Id.IntegerValue
        self.Liste.append(Id)
        if self.vorlauf and self.ruecklauf:return
        cate = elem.Category.Name
        
        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and Id != self.firstt.Id.IntegerValue:
                    return
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.IntegerValue not in self.Liste:   
                            try:
                                if owner.Symbol.FamilyName == '_X_IGF_MC_FlowMeter':
                                    self.ruecklauf = owner
                                    return
                                elif owner.Symbol.FamilyName.find('_HK_IGF_481_MC_') != -1:
                                    self.vorlauf = owner
                                    return
                            except:
                                pass

                            self.get_start_p(owner)
    
    def get_elems(self,elem):
        Id = elem.Id.IntegerValue
        self.Liste.append(Id)
        cate = elem.Category.Name
        
        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2:
                    self.Liste.remove(Id)
                    return
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.IntegerValue not in self.Liste:   
                            if owner.Category.Name == 'Rohre':
                                self.pipes.append(owner)
                            elif owner.Category.Name == 'Rohrformteile':
                                self.bogen.append(owner)

                            self.get_elems(owner)
    
    def dimension(self):
        for elem in self.pipes:
            elem.LookupParameter('Durchmesser').SetValueString('12')
        for elem in self.bogen:
            conns = elem.MEPModel.ConnectorManager.Connectors
            try:
                for conn in conns:
                    mepinfo = conn.GetMEPConnectorInfo()
                    d = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_DIAMETER))
                    r = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_RADIUS))
                    if d != DB.ElementId.InvalidElementId:
                        try:
                            param = doc.GetElement(d)
                            elem.get_Parameter(param.GetDefinition()).SetValueString('12')
                        except:pass
                    if r != DB.ElementId.InvalidElementId:
                        try:
                            param = doc.GetElement(r)
                            elem.get_Parameter(param.GetDefinition()).SetValueString('6')
                        except:pass
                   
            except:pass
        if len(self.pipes) != 0 and len(self.bogen) != 0:
            try:
                active_view.HideElements(self.Liste_Hidden)
            except:
                print('{}: Hide Elements failed'.format(self.elem.Id.IntegerValue))


t = DB.Transaction(doc,'Dimension')
t.Start()
for elem in coll:
    verbraucher = Verbraucher(elem)
    # verbraucher.dimension()

t.Commit()