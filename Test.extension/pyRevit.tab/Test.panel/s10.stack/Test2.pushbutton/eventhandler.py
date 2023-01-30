# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class ZubFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        return True
        if element.Category.Id.ToString() in ['-2008055', '-2001140', '-2008044','-2001160', '-2008049','-2008099', '-2008050']:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Verbraucher:
    def __init__(self,elem,doc,active_view):
        self.elem = elem
        self.doc = doc
        self.active_view = active_view
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
                            param = self.doc.GetElement(d)
                            elem.get_Parameter(param.GetDefinition()).SetValueString('12')
                        except:pass
                    if r != DB.ElementId.InvalidElementId:
                        try:
                            param = self.doc.GetElement(r)
                            elem.get_Parameter(param.GetDefinition()).SetValueString('6')
                        except:pass
                   
            except:pass
        try:
            self.active_view.HideElements(self.Liste_Hidden)
        except:
            print('{}: Hide Elements failed'.format(self.elem.Id.IntegerValue))


class VERBINDEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
                elem = doc.GetElement(el0_ref)
                t = DB.Transaction(doc,'Dimension')
                t.Start()
                # for elem in coll:
                #     verbraucher = Verbraucher(elem)
                #     verbraucher.dimension()
                
                verbraucher = Verbraucher(elem,doc,view)
                verbraucher.dimension()
                t.Commit()
                t.Dispose()
            except:
                break



    def GetName(self):
        return "Dimension"