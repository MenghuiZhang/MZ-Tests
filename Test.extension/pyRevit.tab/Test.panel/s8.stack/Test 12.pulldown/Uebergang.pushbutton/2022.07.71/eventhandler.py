# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB

class UebergangFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008010':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class KanalFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008000':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

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

        el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,KanalFilter(),'Wählt den ersten Luftkanal aus')
        el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,KanalFilter(),'Wählt den zweiten Luftkanal aus')
        el0 = doc.GetElement(el0_ref)
        el1 = doc.GetElement(el1_ref)
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

        o0 = co0.Origin
        o1 = co1.Origin
        o = co0.Origin-co1.Origin
        ln = l0.Direction.Normalize()
        length = abs(o.DotProduct(ln))
        t = DB.Transaction(doc,'Übergangsstück erstellen')
        failureHandlingOptions = t.GetFailureHandlingOptions()
        failureHandler = FailureHandler()
        failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        failureHandlingOptions.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(failureHandlingOptions)
        t.Start()
        try:
            fi = doc.Create.NewTransitionFitting(co1, co0)
            doc.Regenerate()
        except:
            try:
                fi = doc.Create.NewTransitionFitting(co0, co1)
                doc.Regenerate()
            except Exception as e:
                t.RollBack()
                t.Dispose()
                TaskDialog.Show('Fehler','Abstand zu groß!')
                return

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

        t.Commit()
        t.Dispose()
        if not fi:
            return

        t = DB.Transaction(doc,'Länge ändern')
        failureHandlingOptions = t.GetFailureHandlingOptions()
        failureHandler = FailureHandler()
        failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        failureHandlingOptions.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(failureHandlingOptions)
        t.Start()
        co_fi = ''
        for con in cs:
            if not co_fi:
                co_fi = con
            else:
                if co_fi.Id > con.Id:
                    co_fi = con

        if self.GUI.erst.IsChecked:
            if self.GUI.auto.IsChecked:
                # if co_fi.IsConnectedTo(co0):
                try:
                    fi.LookupParameter('L').Set(length)
                    doc.Regenerate()
                except:pass
                # else:
                #     try:
                #         fi.Location.Move(o1-co1.Origin)
                #         doc.Regenerate()
                #         fi.LookupParameter('L').Set(length)
                #         doc.Regenerate()
                #         fi.Location.Move(o0-co0.Origin)
                #     except Exception as e:print(e)
            else:
                # if co_fi.IsConnectedTo(co0):
                try:
                    fi.LookupParameter('L').Set(int(self.GUI.manuell.Content.Text)/304.8)
                    doc.Regenerate()
                except:pass
                # else:
                #     try:
                #         o1_neu_1 = o0 + int(self.GUI.manuell.Content.Text)/304.8*ln
                #         o1_neu_2 = o0 - int(self.GUI.manuell.Content.Text)/304.8*ln
                #         o1_neu_3 = o1 + int(self.GUI.manuell.Content.Text)/304.8*ln
                #         o1_neu_4 = o1 - int(self.GUI.manuell.Content.Text)/304.8*ln
                #         if o1_neu_1.DistanceTo(o1_neu_4) < 0.001 :
                #             fi.Location.Move(o1_neu_1-co1.Origin)
                #         elif o1_neu_2.DistanceTo(o1_neu_3) < 0.001 :
                #             fi.Location.Move(o1_neu_3-co1.Origin)
                #         doc.Regenerate()
                #         fi.LookupParameter('L').Set(int(self.GUI.manuell.Content.Text)/304.8)
                #         doc.Regenerate()
                #         fi.Location.Move(o0-co0.Origin)
                #     except:pass
        else:
            if self.GUI.auto.IsChecked:
                # if co_fi.IsConnectedTo(co1):
                try:
                    fi.LookupParameter('L').Set(length)
                    doc.Regenerate()
                    fi.Location.Move(o1-co1.Origin)
                except:pass
                # else:
                #     try:
                #         fi.Location.Move(o0-co0.Origin)
                #         doc.Regenerate()
                #         fi.LookupParameter('L').Set(length)
                #         doc.Regenerate()
                #         fi.Location.Move(o1-co1.Origin)
                #     except:pass
            else:
                # if co_fi.IsConnectedTo(co1):
                try:
                    fi.LookupParameter('L').Set(int(self.GUI.manuell.Content.Text)/304.8)
                    doc.Regenerate()
                    fi.Location.Move(o1-co1.Origin)
                except:pass
                # else:
                #     try:
                #         o1_neu_1 = o0 + int(self.GUI.manuell.Content.Text)/304.8*ln
                #         o1_neu_2 = o0 - int(self.GUI.manuell.Content.Text)/304.8*ln
                #         o1_neu_3 = o1 + int(self.GUI.manuell.Content.Text)/304.8*ln
                #         o1_neu_4 = o1 - int(self.GUI.manuell.Content.Text)/304.8*ln
                #         abs1 = o1_neu_1-o1_neu_4
                #         abs2 = o1_neu_2-o1_neu_3
                #         print(o0,o1)
                #         print(o1_neu_1,o1_neu_2,o1_neu_3,o1_neu_4)
                #         print(abs1.DotProduct(ln),abs1.DotProduct(ln))
                #         if abs(abs1.DotProduct(ln)) < 0.001 :
                #             fi.Location.Move(o1_neu_1-co0.Origin)
                #         elif abs(abs2.DotProduct(ln)) < 0.001 :
                #             fi.Location.Move(o1_neu_3-co0.Origin)
                #         doc.Regenerate()
                #         fi.LookupParameter('L').Set(int(self.GUI.manuell.Content.Text)/304.8)
                #         doc.Regenerate()
                #         fi.Location.Move(o1-co1.Origin)
                #     except:pass

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Übergansstück erstellen"


class TransitionFitting1(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,KanalFilter(),'Wählt den fixierten Luftkanal aus')
        el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,UebergangFilter(),'Wählt die Übergang aus')
        el0 = doc.GetElement(el0_ref)
        el1 = doc.GetElement(el1_ref)

        conns0 = list(el0.ConnectorManager.Connectors)
        l0 = DB.Line.CreateBound(conns0[0].Origin,conns0[1].Origin)
        
        co0 = None
        co1 = None
        co2_temp = None
        co2 = None

        for con in el1.MEPModel.ConnectorManager.Connectors:
            for con1 in el0.ConnectorManager.Connectors:
                if con.IsConnectedTo(con1):
                    co0 = con
                    co1 = con1
        for con in el1.MEPModel.ConnectorManager.Connectors:
            if con.Id != co0.Id:
                co2_temp = con

        for el in co2_temp.AllRefs:
            if el.Owner.Category.Name == 'Luftkanäle':
                for con in el.ConnectorManager.Connectors:
                    if co2_temp.IsConnectedTo(con):
                        co2 = con

        o0 = co1.Origin
        o1 = co2.Origin
        l_origin = el1.LookupParameter('L').AsDouble()
        l = el0.LookupParameter('Länge').AsDouble()
        lneu = co2.Owner.LookupParameter('Länge').AsDouble()

        ln = l0.Direction.Normalize() 

        t = DB.Transaction(doc,'Länge ändern')
        failureHandlingOptions = t.GetFailureHandlingOptions()
        failureHandler = FailureHandler()
        failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        failureHandlingOptions.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(failureHandlingOptions)
        t.Start()
        if co2_temp.Id > co0.Id:
            if lneu + l_origin > int(self.GUI.length.Text)/304.8:
                try:
                    el1.LookupParameter('L').Set(int(self.GUI.length.Text)/304.8)
                    doc.Regenerate()
                    el1.Location.Move(o0-co1.Origin)
                except Exception as e:print(e)
            else:
                TaskDialog.Show('Fehler','eingegebene Länge zu groß!')
                t.RollBack()
                return
        else:
            try:
                if lneu + l_origin > int(self.GUI.length.Text)/304.8:
                    el1.LookupParameter('L').Set(int(self.GUI.length.Text)/304.8)
                    doc.Regenerate()
                    el1.Location.Move(o0-co1.Origin)
                elif l_origin + l > int(self.GUI.length.Text)/304.8:
                    o1_neu_1 = o0 + int(self.GUI.length.Text)/304.8*ln
                    o1_neu_2 = o0 - int(self.GUI.length.Text)/304.8*ln
                    o1_neu_3 = o1 + int(self.GUI.length.Text)/304.8*ln
                    o1_neu_4 = o1 - int(self.GUI.length.Text)/304.8*ln
                    if o1_neu_1.DistanceTo(o1_neu_4) < 0.001 :
                        el1.Location.Move(o1_neu_1-co1.Origin)
                    elif o1_neu_2.DistanceTo(o1_neu_3) < 0.001 :
                        el1.Location.Move(o1_neu_3-co1.Origin)
                    doc.Regenerate()
                    el1.LookupParameter('L').Set(int(self.GUI.length.Text)/304.8)
                    doc.Regenerate()
                    el1.Location.Move(o0-co1.Origin)
                
                else:
                    TaskDialog.Show('Fehler','eingegebene Länge zu groß!')
            except Exception as e:print(e)
       
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Übergansstück erstellen"