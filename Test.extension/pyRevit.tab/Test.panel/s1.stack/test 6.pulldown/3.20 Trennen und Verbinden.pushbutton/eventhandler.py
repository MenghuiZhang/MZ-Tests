# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
import math

class ZubFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Name == 'Luftkanalzubehör':
            return True
        if element.Category.Id.ToString() in ['-2008055', '-2001140', '-2008044','-2001160', '-2008049','-2008099', '-2008050','-2008020','-2008000','-2008010']:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class ZubFilter1(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Name == 'Luftkanalformteile':
            return True
        if element.Category.Id.ToString() in ['-2008055', '-2001140', '-2008044','-2001160', '-2008049','-2008099', '-2008050','-2008020']:
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
        elemids = uidoc.Selection.GetElementIds()
        el0 = None
        el1 = None
        # el2 = None
        for elid in elemids:
            el = doc.GetElement(elid)
            if el.Category.Name == 'Luftdurchlässe':
                el0 = el
            else:
                if not el1:
                    el1 = el
                # else:
                #     el2 = el
        # el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
        # el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter1(),'Wählt den Rohr aus')
        # el2_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter1(),'Wählt den Rohr aus')

        # el0 = doc.GetElement(el0_ref)
        # el1 = doc.GetElement(el1_ref)
        # el2 = doc.GetElement(el2_ref)

        # conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        try:
            conns0 = list(el0.ConnectorManager.Connectors)
        except:
            conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        try:
            conns1 = list(el1.ConnectorManager.Connectors)
        except:
            conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
        # try:
        #     conns2 = list(el2.ConnectorManager.Connectors)
        # except:
        #     conns2 = list(el2.MEPModel.ConnectorManager.Connectors)
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
        
        # conn2 = ''
        # conn3 = ''
        # dis = 1000
        # for co0 in conns0:
        #     for co1 in conns2:
        #         if co1.IsConnected == False and co0.IsConnected == False:
        #             l0 = co0.Origin.DistanceTo(co1.Origin)
        #             if l0 < dis:
        #                 dis = l0
        #                 conn2 = co0
        #                 conn3 = co1

        if not(conn0 and conn1):
            print('fehler')
            return

        t = DB.Transaction(doc,'Verbinden')
        t.Start()
        try:
            conn0.ConnectTo(conn1)
            conn1.Origin = conn0.Origin
        except Exception as e:
            print(e)
        # try:
        #     conn2.ConnectTo(conn3)
        # except Exception as e:
        #     print(e)
        
        t.Commit()
        t.Dispose()
        # while(True):
        #     try:
        #         el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
        #         el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter1(),'Wählt den Rohr aus')
        #         el2_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter1(),'Wählt den Rohr aus')

        #         el0 = doc.GetElement(el0_ref)
        #         el1 = doc.GetElement(el1_ref)
        #         el2 = doc.GetElement(el2_ref)

        #         conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        #         try:
        #             conns1 = list(el1.ConnectorManager.Connectors)
        #         except:
        #             conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
        #         try:
        #             conns2 = list(el2.ConnectorManager.Connectors)
        #         except:
        #             conns2 = list(el2.MEPModel.ConnectorManager.Connectors)
        #         conn0 = ''
        #         conn1 = ''
        #         dis = 1000
        #         for co0 in conns0:
        #             for co1 in conns1:
        #                 if co1.IsConnected == False and co0.IsConnected == False:
        #                     l0 = co0.Origin.DistanceTo(co1.Origin)
        #                     if l0 < dis:
        #                         dis = l0
        #                         conn0 = co0
        #                         conn1 = co1
                
        #         conn2 = ''
        #         conn3 = ''
        #         dis = 1000
        #         for co0 in conns0:
        #             for co1 in conns2:
        #                 if co1.IsConnected == False and co0.IsConnected == False:
        #                     l0 = co0.Origin.DistanceTo(co1.Origin)
        #                     if l0 < dis:
        #                         dis = l0
        #                         conn2 = co0
        #                         conn3 = co1

        #         if not(conn0 and conn1):
        #             print('fehler')
        #             return

        #         t = DB.Transaction(doc,'Verbinden')
        #         t.Start()
        #         try:
        #             conn0.ConnectTo(conn1)
        #         except Exception as e:
        #             print(e)
        #         try:
        #             conn2.ConnectTo(conn3)
        #         except Exception as e:
        #             print(e)
                
        #         t.Commit()
        #         t.Dispose()
        #     except:break


    def GetName(self):
        return "Verbinden"

class TRENNEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        # uidoc = app.ActiveUIDocument
        # doc = uidoc.Document
        # try:
        #     # el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
        #     el0 = doc.GetElement(uidoc.Selection.GetElementIds()[0])
        #     conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        #     conn_0 = None
        #     conn_1 = None
        #     # conn_2 = None
        #     # conn_3 = None
        #     for conn in conns0:
        #         if not conn_0: conn_0 = conn
        #         else:conn_1 = conn
        #         # allrefs = conn.AllRefs
        #         # for ref in allrefs:
        #         #     owner = ref.Owner
        #         #     if owner.Category.Name in ['Luftkanalformteile','Luftkanäle']:

        #         #         if not conn_0:
        #         #             conn_0 = conn
        #         #             conn_1 = ref
        #         #         else:
        #         #             conn_2 = conn
        #         #             conn_3 = ref
        #     t = DB.Transaction(doc,'Test')
        #     t.Start()
        #     # conn_0.DisconnectFrom(conn_1)
        #     # conn_2.DisconnectFrom(conn_3)
        #     doc.Regenerate()
        #     o0 = conn_0.Origin
        #     o1 = conn_1.Origin
        #     A = o0.X-o1.X
        #     B = o0.Y-o1.Y
        #     C = o0.Z-o1.Z
        #     D = A*(o0.X+o1.X)/2.0+B*(o0.Y+o1.Y)/2.0+C*(o0.Z+o1.Z)/2.0
        #     X2 = 0
        #     Y2 = 0
        #     Z2 = 0
        #     if A != 0:
        #         X2 = D/A
        #     elif B != 0:
        #         Y2 = D/B
        #     elif C != 0:
        #         Z2 = D/C
            
        #     Line = DB.Line.CreateBound(DB.XYZ((o0.X+o1.X)/2,(o0.Y+o1.Y)/2,(o0.Z+o1.Z)/2),DB.XYZ(X2,Y2,Z2))
        #     Line1 = DB.Line.CreateBound(o1,o0)
        #     DB.ElementTransformUtils.RotateElement(doc, el0.Id, Line,float(math.pi))
        #     # print(DB.XYZ((o0.X+o1.X)/2,(o0.Y+o1.Y)/2,(o0.Z+o1.Z)/2),DB.XYZ(X2,Y2,Z2))
        #     # print(o0,o1)
        #     # el0.Location.Rotate(Line,math.pi)
        #     # doc.Regenerate()
        #     # conn_0.ConnectTo(conn_3)
        #     # conn_2.ConnectTo(conn_1)
        #     t.Commit()
        #     t.Dispose()
        # except Exception as e:
        #     print(e)

        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        while(True):
            try:

                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohr aus')

                el0 = doc.GetElement(el0_ref)
                el1 = doc.GetElement(el1_ref)

                conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
                try:
                    conns1 = list(el1.ConnectorManager.Connectors)
                except:
                    conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
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
        try:
            # el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den Rohrzubehör aus')
            el0 = doc.GetElement(uidoc.Selection.GetElementIds()[0])
            conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
            conn_0 = None
            conn_1 = None
            conn_2 = None
            conn_3 = None
            for conn in conns0:
                allrefs = conn.AllRefs
                for ref in allrefs:
                    owner = ref.Owner
                    if owner.Category.Name in ['Luftkanalformteile','Luftkanäle']:

                        if not conn_0:
                            conn_0 = conn
                            conn_1 = ref
                        else:
                            conn_2 = conn
                            conn_3 = ref
            t = DB.Transaction(doc,'Test')
            t.Start()
            conn_0.DisconnectFrom(conn_1)
            conn_2.DisconnectFrom(conn_3)
            doc.Regenerate()
            o0 = conn_0.Origin
            o1 = conn_2.Origin
            A = o0.X-o1.X
            B = o0.Y-o1.Y
            C = o0.Z-o1.Z
            D = A*(o0.X+o1.X)/2.0+B*(o0.Y+o1.Y)/2.0+C*(o0.Z+o1.Z)/2.0
            X2 = 0
            Y2 = 0
            Z2 = 0
            if A != 0:
                X2 = D/A
            elif B != 0:
                Y2 = D/B
            elif C != 0:
                Z2 = D/C
            
            # Line = DB.Line.CreateUnbound(DB.XYZ((o0.X+o1.X)/2,(o0.Y+o1.Y)/2,(o0.Z+o1.Z)/2),DB.XYZ(X2,Y2,Z2))
            # DB.ElementTransformUtils.RotateElement(doc, el0.Id, Line,float(math.pi/2))
            # print(DB.XYZ((o0.X+o1.X)/2,(o0.Y+o1.Y)/2,(o0.Z+o1.Z)/2),DB.XYZ(X2,Y2,Z2))
            # print(o0,o1)
            # el0.Location.Rotate(Line,math.pi)
            # doc.Regenerate()
            # conn_0.ConnectTo(conn_3)
            # conn_2.ConnectTo(conn_1)
            t.Commit()
            t.Dispose()
        except Exception as e:
            print(e)

        # view = uidoc.ActiveView.GenLevel.Id
       
        # while(True):
        #     try:
        #         el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den ersten Teil aus')
        #         el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZubFilter(),'Wählt den zweiten Teil aus')
        #         el0 = doc.GetElement(el0_ref)
        #         el1 = doc.GetElement(el1_ref)
        #         try:
        #             conns0 = list(el0.ConnectorManager.Connectors)
        #         except:
        #             conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
       
        #         try:
        #             conns1 = list(el1.ConnectorManager.Connectors)
        #         except:
        #             conns1 = list(el1.MEPModel.ConnectorManager.Connectors)

                
        #         co0 = None
        #         co1 = None
         
        #         distance = 1000

        #         for con in conns0:
        #             for con1 in conns1:
        #                 if con.IsConnected == False and con1.IsConnected == False:
        #                     dis = con.Origin.DistanceTo(con1.Origin)
        #                     if dis < distance:
        #                         distance = dis
        #                         co0 = con
        #                         co1 = con1
        #         if not (co0 and co1):
        #             return
                                
        #         t = DB.Transaction(doc,'Übergang')
        #         t.Start()
        #         try:
        #             rohr = DB.Plumbing.Pipe.Create(doc,DB.ElementId(2170325),view,co1,co0)
        #         except:
        #             try: rohr = DB.Plumbing.Pipe.Create(doc,DB.ElementId(2170325),view,co0,co1)
        #             except:print('nicht geklappt')
        #         doc.Regenerate()
        #         mep = rohr.MEPSystem
        #         doc.Delete(mep.Id)
        #         t.Commit()
        #         t.Dispose()
        #     except:
        #         break

    def GetName(self):
        return "Rohr erstellen"