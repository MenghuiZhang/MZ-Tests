# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class FailureHandler(DB.IFailuresPreprocessor):
    def __init__(self):
        self.ErrorMessage = ""
        self.ErrorSeverity = ""
 
    def PreprocessFailures(self,failuresAccessor):
        failureMessages = failuresAccessor.GetFailureMessages()
        for failureMessageAccessor in failureMessages:
            Id = failureMessageAccessor.GetFailureDefinitionId()
            try:self.ErrorMessage = failureMessageAccessor.GetDescriptionText()
            except:self.ErrorMessage = "Unknown Error"
            try:
                failureSeverity = failureMessageAccessor.GetSeverity()
                self.ErrorSeverity = failureSeverity.ToString()
                if failureSeverity == DB.FailureSeverity.Warning:failuresAccessor.DeleteWarning(failureMessageAccessor)
                else:return DB.FailureProcessingResult.ProceedWithRollBack
            except:pass
        return DB.FailureProcessingResult.Continue


class TransitionFitting(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        sel = uidoc.Selection.GetElementIds()
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
        conns0 = list(el0.ConnectorManager.Connectors)
        conns1 = list(el1.ConnectorManager.Connectors)
        
        co0 = None
        co1 = None
        distance = 1000
        l0 = DB.Line.CreateBound(conns0[0].Origin,conns0[1].Origin)
        l1 = DB.Line.CreateBound(conns1[0].Origin,conns1[1].Origin)
        if not ((l0.Direction+l1.Direction).GetLength() < 1e-10 or (l0.Direction-l1.Direction).GetLength() < 1e-10 ):
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
        if not (co0 and co1):
            TaskDialog.Show('Fehler','Kein oder nur ein offener Anschluss gefunden!')
            return
        print(co0.Origin, co1.Origin)

        
        o0 = co0.Origin
        o = co0.Origin-co1.Origin
        ln = l0.Direction.Normalize()
        length = o.DotProduct(ln)
        t = DB.Transaction(doc,'Übergangsstück erstellen')
        failureHandlingOptions = t.GetFailureHandlingOptions()
        failureHandler = FailureHandler()
        failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        failureHandlingOptions.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(failureHandlingOptions)
        t.Start()
        try:
            if co0.Origin.Y > co1.Origin.Y and co0.Origin.X > co1.Origin.X:fi = doc.Create.NewTransitionFitting(co0, co1)
            elif co1.Origin.Y > co0.Origin.Y and co1.Origin.X > co0.Origin.X:fi = doc.Create.NewTransitionFitting(co1, co0)
            else:fi = doc.Create.NewTransitionFitting(co0, co1)
            doc.Regenerate()
        except:
            try:
                fi = doc.Create.NewTransitionFitting(co1, co0)
                doc.Regenerate()
            except:
                try:
                    origin = (co0.Origin + co1.Origin)/2
                    plane = DB.Plane.CreateByOriginAndBasis(origin,uidoc.ActiveView.ViewDirection,ln)
                    elemliste = DB.ElementTransformUtils.MirrorElements(doc,sel,plane,True)
                    el0_temp = doc.GetElement(elemliste[0])
                    el1_temp = doc.GetElement(elemliste[1])
                    distance_temp = 1000
                    conns0_temp = list(el0_temp.ConnectorManager.Connectors)
                    conns1_temp = list(el1_temp.ConnectorManager.Connectors)
        
                    co0_temp = None
                    co1_temp = None
                    for con in conns0_temp:
                        for con1 in conns1_temp:
                            dis = con.Origin.DistanceTo(con1.Origin)
                            if dis < distance_temp:
                                distance_temp = dis
                                co0_temp = con
                                co1_temp = con1
                    try:
                        if co0_temp.Origin.Y > co1_temp.Origin.Y:fi = doc.Create.NewTransitionFitting(co0_temp, co1_temp)
                        elif co1_temp.Origin.Y > co0_temp.Origin.Y and co1_temp.Origin.X > co0_temp.Origin.X:fi = doc.Create.NewTransitionFitting(co1_temp, co0_temp)
                        else:fi = doc.Create.NewTransitionFitting(co0_temp, co1_temp)
                        doc.Regenerate()
                    except:
                        try:
                            fi = doc.Create.NewTransitionFitting(co1_temp, co0_temp)
                            doc.Regenerate()
                        except Exception as e:
                            sel.Clear()
                            uidoc.Selection.SetElementIds(sel)
                            t.RollBack()
                            t.Dispose()
                            TaskDialog.Show('Fehler',e.ToString())
                            return
                    doc.Delete(elemliste)
                    doc.Regenerate()
                    fitemp = DB.ElementTransformUtils.MirrorElements(doc,List[DB.ElementId]([fi.Id]),plane,True)
                    doc.Delete(fi.Id)
                    print(fitemp)
                    doc.Regenerate()
                    fi = doc.GetElement(fitemp[0])

                except Exception as e:
                    sel.Clear()
                    uidoc.Selection.SetElementIds(sel)
                
                    t.RollBack()
                    t.Dispose()
                    TaskDialog.Show('Fehler',e.ToString())
        sel.Clear()
        uidoc.Selection.SetElementIds(sel)

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
        try:fi.Location.Move(o0-co0.Origin)
        except Exception as e:print(e)

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Übergansstück erstellen"