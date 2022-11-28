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
        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohr aus')
                el2_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohr aus')

                el0 = doc.GetElement(el0_ref)
                el1 = doc.GetElement(el1_ref)
                el2 = doc.GetElement(el2_ref)

                try:

                    conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
                except:
                    conns0 = list(el0.ConnectorManager.Connectors)
                try:

                    conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
                except:
                    conns1 = list(el1.ConnectorManager.Connectors)
                
                

                conn0 = ''
                conn1 = ''
                dis = 1000
                for co0 in conns0:
                    for co1 in conns1:
                        if co1.IsConnected == False and co0.IsConnected == False:
                            l0 = co0.Origin.DistanceTo(co1.Origin)
                            if l0 < dis:
                                dis = l0
                                conn0 = co0
                                conn1 = co1

                if not(conn0 and conn1):
                    print('fehler')
                    return

                t = DB.Transaction(doc,'Verbinden')
                t.Start()
                try:
                    el1.Location.Move(conn0.Origin-conn1.Origin)
                    conn1.Origin = conn0.Origin
                except Exception as e:
                    pass
                try:
                    conn0.ConnectTo(conn1)
                except Exception as e:
                    print(e)
                
                doc.Regenerate()
                try:

                    conns0 = list(el1.MEPModel.ConnectorManager.Connectors)
                except:
                    conns0 = list(el1.ConnectorManager.Connectors)
                try:

                    conns1 = list(el2.MEPModel.ConnectorManager.Connectors)
                except:
                    conns1 = list(el2.ConnectorManager.Connectors)
                
                conn0 = ''
                conn1 = ''
                dis = 1000
                for co0 in conns0:
                    for co1 in conns1:
                        if co1.IsConnected == False and co0.IsConnected == False:
                            l0 = co0.Origin.DistanceTo(co1.Origin)
                            if l0 < dis:
                                dis = l0
                                conn0 = co0
                                conn1 = co1

                if not(conn0 and conn1):
                    print('fehler')
                    return

                # t = DB.Transaction(doc,'Verbinden')
                # t.Start()
                try:
                    el2.Location.Move(conn0.Origin-conn1.Origin)
                    conn1.Origin = conn0.Origin
                except Exception as e:
                    pass
                try:
                    conn0.ConnectTo(conn1)
                except Exception as e:
                    print(e)

                
                t.Commit()
                t.Dispose()
            except:break


    def GetName(self):
        return "Verbinden"

class TRENNEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        while(True):
            try:

                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohr aus')

                el0 = doc.GetElement(el0_ref)
                el1 = doc.GetElement(el1_ref)

                conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
                conns1 = list(el1.ConnectorManager.Connectors)
                conn0 = ''
                conn1 = ''
                for co0 in conns0:
                    for co1 in conns1:
                        if co0.IsConnectedTo(co1):
                            conn0 = co0
                            conn1 = co1

                if not(conn0 and conn1):
                    print('fehler')
                    return

                t = DB.Transaction(doc,'Trennen')
                t.Start()

                try:
                    conn0.DisconnectFrom(conn1)
                except Exception as e:
                    print(e)
                
                t.Commit()
                t.Dispose()
            except:break


    def GetName(self):
        return "Trennen"

class ROHRERSTELLEN(IExternalEventHandler):        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView.GenLevel.Id
       
        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den ersten Teil aus')
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den zweiten Teil aus')
                el0 = doc.GetElement(el0_ref)
                el1 = doc.GetElement(el1_ref)
                try:
                    conns0 = list(el0.ConnectorManager.Connectors)
                except:
                    conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
       
                try:
                    conns1 = list(el1.ConnectorManager.Connectors)
                except:
                    conns1 = list(el1.MEPModel.ConnectorManager.Connectors)

                
                co0 = None
                co1 = None
         
                distance = 1000

                for con in conns0:
                    for con1 in conns1:
                        if con.IsConnected == False and con1.IsConnected == False:
                            dis = con.Origin.DistanceTo(con1.Origin)
                            if dis < distance:
                                distance = dis
                                co0 = con
                                co1 = con1
                if not (co0 and co1):
                    return
                                
                t = DB.Transaction(doc,'Übergang')
                t.Start()
                try:
                    rohr = DB.Plumbing.Pipe.Create(doc,DB.ElementId(9759272),view,co1,co0)
                except:
                    try: rohr = DB.Plumbing.Pipe.Create(doc,DB.ElementId(9759272),view,co0,co1)
                    except:print('nicht geklappt')
                doc.Regenerate()
                mep = rohr.MEPSystem
                doc.Delete(mep.Id)
                t.Commit()
                t.Dispose()
            except:
                break

    def GetName(self):
        return "Rohr erstellen"