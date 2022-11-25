# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class HLSFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2001140':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class RohrFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class VERBINDEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView

        el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,HLSFilter(),'Wählt den ULK aus')
        el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,RohrFilter(),'Wählt den ersten Rohr aus')
        el2_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,RohrFilter(),'Wählt den zweiten Rohr aus')
        el0 = doc.GetElement(el0_ref)
        el1 = doc.GetElement(el1_ref)
        el2 = doc.GetElement(el2_ref)
        conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        conns1 = list(el1.ConnectorManager.Connectors)
        conns2 = list(el2.ConnectorManager.Connectors)
        
        co01 = None
        co02 = None
        co10 = None
        co20 = None
        for con in conns0:
            if con.PipeSystemType.ToString() == 'ReturnHydronic':
                co01 = con
            elif con.PipeSystemType.ToString() == 'SupplyHydronic':
                co02 = con
        for con in conns1:
            if con.IsConnected == False:
                if con.PipeSystemType.ToString() == 'ReturnHydronic':
                    co10 = con
                else:
                    co20 = con
        for con in conns2:
            if con.IsConnected == False:
                if con.PipeSystemType.ToString() == 'ReturnHydronic':
                    co10 = con
                else:
                    co20 = con

        
        t = DB.Transaction(doc,'Verbinden')
        t.Start()
        try:
            co10.Origin = co01.Origin
        except Exception as e:
            print(e)
        
        try:
            co20.Origin = co02.Origin
        except Exception as e:
            print(e)
        doc.Regenerate()
        try:
            co02.ConnectTo(co20)
        except Exception as e:
            print(e)
        try:
            co01.ConnectTo(co10)
        except Exception as e:
            print(e)
        view.HideElements(List[DB.ElementId]([el0.Id]))
        

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Verbinden"