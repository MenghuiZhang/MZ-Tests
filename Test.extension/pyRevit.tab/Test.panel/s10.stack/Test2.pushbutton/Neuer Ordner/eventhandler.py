# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit
from Autodesk.Revit.Exceptions import InvalidOperationException  


class TransitionFitting(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.co0 = ''
        self.co1 = ''
        self.l0 = ''
        
    def Execute(self,app):
        # print(0)
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        # processor = FailuresProcessor()
        # app.Application.RegisterFailuresProcessor(processor)
        Liste = []
        for elid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(elid)
            if el.Category.Name == 'Luftkanäle':
                 Liste.append(el)
        if len(Liste) != 2:
            TaskDialog.Show('Fehler','Bitte zwei Luftkanäle auswählen!')
            return
       
        el0 = Liste[0]
        el1 = Liste[1]
        # print(el0.Id,el1.Id)
        conns0 = list(el0.ConnectorManager.Connectors)
        conns1 = list(el1.ConnectorManager.Connectors)
        
        co0 = None
        co1 = None
        distance = 1000
        l0 = DB.Line.CreateBound(conns0[0].Origin,conns0[1].Origin)
        l1 = DB.Line.CreateBound(conns1[0].Origin,conns1[1].Origin)
        if not ((l0.Direction+l1.Direction).GetLength() < 0.00001 or (l0.Direction-l1.Direction).GetLength() < 0.00001):
            TaskDialog.Show('Fehler','Die Luftkanäle sind nicht Parallel!')
            return
        for con in conns0:
            for con1 in conns1:
                if con.IsConnected == False and con1.IsConnected == False:
                    dis = con.Origin.DistanceTo(con1.Origin)
                    if dis < distance:
                        distance = dis
                        co0 = con
                        co1 = con1
        # print(co0.Id,co1.Id)
        if not (co0 and co1):
            TaskDialog.Show('Fehler','Kein oder nur ein offener Anschluss gefunden!')
            return
                
        o0 = co0.Origin
        

        o = co0.Origin-co1.Origin
        ln = l0.Direction.Normalize()
        length = o.DotProduct(ln)

        t = DB.Transaction(doc,'Übergangsstück')
        # failureHandlingOptions = t.GetFailureHandlingOptions()
        # failureHandler = FailureHandler()
        # failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        # failureHandlingOptions.SetClearAfterRollback(True)
        # t.SetFailureHandlingOptions(failureHandlingOptions)
        t.Start()
        try:
            fi = doc.Create.NewTransitionFitting(co0, co1)
            doc.Regenerate()
        except:
            try:fi = doc.Create.NewTransitionFitting(co1, co0)
            except:
                try:
                    elem1 = DB.Mechanical.Duct.Create(doc,el0.LookupParameter('Systemtyp').AsElementId(),el0.GetTypeId(),el0.ReferenceLevel.Id ,conns0[0].Origin,conns0[1].Origin)
                    elem2 = DB.Mechanical.Duct.Create(doc,el1.LookupParameter('Systemtyp').AsElementId(),el1.GetTypeId(),el1.ReferenceLevel.Id ,conns1[0].Origin,conns1[1].Origin)
                    doc.Regenerate()
       
                    conns0_temp = list(elem1.ConnectorManager.Connectors)
                    conns1_temp  = list(elem2.ConnectorManager.Connectors)
                    
                    co0_temp  = None
                    co1_temp  = None
                    distance_temp = 1000
                    for con in conns0_temp :
                        # try:
                        #     print(con.Width)
                        # except:
                        #     pass
                        # try:
                        #     print(co0.Width)
                        # except:
                        #     pass
                        # con.Width = co0.Width
                        # con.Height = co0.Height
                        if (con.Origin-co0.Origin).GetLength() < 0.00001:
                            co0_temp  = con
                            con.Width = co0.Width
                            con.Height = co0.Height
                    for con in conns1_temp :
                        print(1)
                        # con.Width = co1.Width
                        # con.Height = co1.Height
                        if (con.Origin-co1.Origin).GetLength() < 0.00001:
                            co1_temp  = con
                            con.Width = co1.Width
                            con.Height = co1.Height
                    
                    try:
                        fi = doc.Create.NewTransitionFitting(co1_temp, co0_temp)
                    except:
                        try:
                            fi = doc.Create.NewTransitionFitting(co0_temp, co1_temp)
                        except Exception as e:
                            print(e)
                            print('.....')
                            t.RollBack()
                            return
                except Exception as e:
                    print('.......')
                    print(e)
                    t.RollBack()
                    return
                doc.Delete(elem1.Id)
                doc.Delete(elem2.Id)
                doc.Regenerate()
                try:fi.LookupParameter('L').Set(length)
                except Exception as e:print(e)
                doc.Regenerate()
                try:fi.Location.Move(o0-co0.Origin)
                except Exception as e:print(e)    

        try:

            cs = fi.MEPModel.ConnectorManager.Connectors
            for con in cs:
                if con.Origin.DistanceTo(co0.Origin) < 0.001:
                    
                    if not con.IsConnected:
                        con.ConnectTo(co0)
                        con.Width = co0.Width                                    
                        con.Height = co0.Height
                else:

                    if not con.IsConnected:
                        con.ConnectTo(co1)
                        con.Width = co1.Width                                    
                        con.Height = co1.Height
        except:pass
        try:fi.LookupParameter('L').Set(length)
        except Exception as e:print(e)
        doc.Regenerate()
        try:fi.Location.Move(o0-co0.Origin)
        except Exception as e:print(e)    
        t.Commit()

        # if DB.TransactionStatus.Committed == t.Commit():return
        # opt = t.GetFailureHandlingOptions()
        # t.RollBack(opt.SetClearAfterRollback(True))
        # t.Dispose()


        # t1 = DB.Transaction(doc,'Übergangsstück')
        # t1.Start()
        # print(co1.Id,co0.Id)
        # fi = doc.Create.NewTransitionFitting(co1, co0)
        # doc.Regenerate()

        # cs = fi.MEPModel.ConnectorManager.Connectors
        # for con in cs:
        #     if con.Origin.DistanceTo(co0.Origin) < 0.001:
                
        #         if not con.IsConnected:
        #             con.ConnectTo(co0)
        #             con.Width = co0.Width                                    
        #             con.Height = co0.Height
        #     else:

        #         if not con.IsConnected:
        #             con.ConnectTo(co1)
        #             con.Width = co1.Width                                    
        #             con.Height = co1.Height
        # try:fi.LookupParameter('L').Set(length)
        # except Exception as e:print(e)
        # doc.Regenerate()
        # try:fi.Location.Move(o0-co0.Origin)
        # except Exception as e:print(e)    

        # if DB.TransactionStatus.Committed == t1.Commit():return
        # opt = t1.GetFailureHandlingOptions()
        # t1.RollBack(opt.SetClearAfterRollback(True))

        # t1.Dispose()
        # if co0.IsConnected == False and co1.IsConnected == False:
        #     t1 = DB.Transaction(doc,'Übergangsstück')
        #     t1.Start()
        #     fi = doc.Create.NewTransitionFitting(co1, co0)
        #     doc.Regenerate()

        #     cs = fi.MEPModel.ConnectorManager.Connectors
        #     for con in cs:
        #         if con.Origin.DistanceTo(co0.Origin) < 0.001:
                    
        #             if not con.IsConnected:
        #                 con.ConnectTo(co0)
        #                 con.Width = co0.Width                                    
        #                 con.Height = co0.Height
        #         else:

        #             if not con.IsConnected:
        #                 con.ConnectTo(co1)
        #                 con.Width = co1.Width                                    
        #                 con.Height = co1.Height
        #     try:fi.LookupParameter('L').Set(length)
        #     except Exception as e:print(e)
        #     doc.Regenerate()
        #     try:fi.Location.Move(o0-co0.Origin)
        #     except Exception as e:print(e)    

        #     t1.Commit()
        #     t1.Dispose()


    def GetName(self):
        return "Übergansstück erstellen"

class TransitionFitting1(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.o1 = ''
        
    def Execute(self,app):
        print(1)
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not (self.GUI.transitionfitting.co0 and self.GUI.transitionfitting.co1):
            print(2)
            return
        co0 = self.GUI.transitionfitting.co0
        co1 = self.GUI.transitionfitting.co1
        o1 = self.GUI.transitionfitting.co1.Origin
        o = self.GUI.transitionfitting.co0.Origin-self.GUI.transitionfitting.co1.Origin
        ln = self.GUI.transitionfitting.l0.Direction.Normalize()
        length = o.DotProduct(ln)


        if co0.IsConnected and co1.IsConnected:
            return
        

        tran = DB.Transaction(doc,'Übergangsstück1')
        print(co1.Id,co0.Id)
        # failureHandlingOptions = t.GetFailureHandlingOptions()
        # failureHandler = FailureHandler()
        # failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        # failureHandlingOptions.SetClearAfterRollback(False)
        # t.SetFailureHandlingOptions(failureHandlingOptions)
        tran.Start()
        print(co1.Id,co0.Id)
        fi = doc.Create.NewTransitionFitting(co1, co0)
        print(co1.Id,co0.Id)
        doc.Regenerate()
        print(co1.Id,co0.Id)

        cs = fi.MEPModel.ConnectorManager.Connectors
        for con in cs:
            if con.Origin.DistanceTo(co0.Origin) < 0.001:
                
                if not con.IsConnected:
                    con.ConnectTo(co0)
                    con.Width = co0.Width                                    
                    con.Height = co0.Height
            else:

                if not con.IsConnected:
                    con.ConnectTo(co1)
                    con.Width = co1.Width                                    
                    con.Height = co1.Height
        try:fi.LookupParameter('L').Set(length)
        except Exception as e:print(e)
        doc.Regenerate()
        try:fi.Location.Move(o1-co1.Origin)
        except Exception as e:print(e)    

        tran.Commit()
        tran.Dispose()

    def GetName(self):
        return "Übergansstück erstellen"