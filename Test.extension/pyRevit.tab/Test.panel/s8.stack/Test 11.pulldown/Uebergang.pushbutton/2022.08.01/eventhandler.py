# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB

class FormteilFilter(Selection.ISelectionFilter):
    # Formteil
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

class KanalFormteilFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008016' or element.Category.Id.ToString() == '-2008010':
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
                TaskDialog.Show('Fehler','Versatzbreite/-höhe zu groß!')
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

        t1 = DB.Transaction(doc,'Länge ändern')
        failureHandlingOptions = t1.GetFailureHandlingOptions()
        failureHandler = FailureHandler()
        failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
        failureHandlingOptions.SetClearAfterRollback(True)
        t1.SetFailureHandlingOptions(failureHandlingOptions)
        t1.Start()
        co_fi = ''
        for con in cs:
            if not co_fi:
                co_fi = con
            else:
                if co_fi.Id > con.Id:
                    co_fi = con
        l_Parameter = fi.LookupParameter('L')
        if l_Parameter == None:l_Parameter = fi.LookupParameter('l')
        if l_Parameter == None:l_Parameter = fi.LookupParameter('DIN_l')
        if l_Parameter == None:
            self.GUI.Close()
            t1.Commit()
            text = 'Das Skript funktionert nicht mit Familie '+fi.Symbol.FamilyName+'. Übergangslänge kann nicht geändert werden. Bitte melden Sie sich bei Menghui.'
            TaskDialog.Show('Fehler',text)
            return

        if self.GUI.erst.IsChecked:
            if self.GUI.auto.IsChecked:
                try:
                    l_Parameter.Set(length)
                    doc.Regenerate()
                except:pass

            else:
                try:
                    l_Parameter.Set(int(self.GUI.manuell.Content.Text)/304.8)
                    doc.Regenerate()
                except:pass

        else:
            if self.GUI.auto.IsChecked:
                try:
                    l_Parameter.Set(length)
                    doc.Regenerate()
                    fi.Location.Move(o1-co1.Origin)
                except:pass
                
            else:
                try:
                    l_Parameter.Set(int(self.GUI.manuell.Content.Text)/304.8)
                    doc.Regenerate()
                    fi.Location.Move(o1-co1.Origin)
                except:pass
                

        t1.Commit()
        t1.Dispose()

    def GetName(self):
        return "Übergansstück erstellen"

class TransitionFitting1(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        if self.GUI.luftkanal.IsChecked == True:
            el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,KanalFilter(),'Wählt den fixierten Luftkanal aus')
            el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,FormteilFilter(),'Wählt die Übergang aus')

            el0 = doc.GetElement(el0_ref)
            el1 = doc.GetElement(el1_ref)

            conns0 = list(el0.ConnectorManager.Connectors)
            conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
            
            co0 = None
            co0_ = None
            co1 = None
            co2 = None
            

            for con1 in conns1:
                for con0 in conns0:
                    if con0.IsConnectedTo(con1):
                        co0 = con0
                        co1 = con1

            for con in conns0:
                if con.Id != co0.Id:co0_ = con

            for con in conns1:
                if con.Id != co1.Id:co2 = con
            
            l0 = DB.Line.CreateBound(co0.Origin,co0_.Origin)
            ln = l0.Direction.Normalize()

            o0 = co1.Origin

            l_Parameter = el1.LookupParameter('L')
            if l_Parameter == None:l_Parameter = el1.LookupParameter('l')
            if l_Parameter == None:l_Parameter = el1.LookupParameter('DIN_l')
            if l_Parameter == None:
                self.GUI.Close()
                text = 'Das Skript funktionert nicht mit Familie '+el1.Symbol.FamilyName+'. Übergangslänge kann nicht geändert werden. Bitte melden Sie sich bei Menghui.'
                TaskDialog.Show('Fehler',text)
                return

            l_origin = l_Parameter.AsDouble()
            l = el0.LookupParameter('Länge').AsDouble()

            t = DB.Transaction(doc,'Länge ändern')
            failureHandlingOptions = t.GetFailureHandlingOptions()
            failureHandler = FailureHandler()
            failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
            failureHandlingOptions.SetClearAfterRollback(True)
            t.SetFailureHandlingOptions(failureHandlingOptions)
            t.Start()
            if co2.Id > co1.Id:
                try:
                    l_Parameter.Set(int(self.GUI.length.Text)/304.8)
                    doc.Regenerate()
                    el1.Location.Move(o0-co1.Origin)
                except Exception as e:print(e)

            else:
                if l_origin + l > int(self.GUI.length.Text)/304.8:
                    try:
                        el1.Location.Move((int(self.GUI.length.Text)/304.8-l_origin)*ln)
                        l_Parameter.Set(int(self.GUI.length.Text)/304.8)
                        doc.Regenerate()
                        el1.Location.Move(o0-co1.Origin)
                    except Exception as e:print(e)
                else:
                    TaskDialog.Show('Fehler','eingegebene Länge zu groß!')
        
            t.Commit()
            t.Dispose()
        
        else:
            el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,KanalFormteilFilter(),'Wählt den fixierten Formteil/Zubehör aus')
            el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,FormteilFilter(),'Wählt die Übergang aus')

            el0 = doc.GetElement(el0_ref)
            el1 = doc.GetElement(el1_ref)

            conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
            conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
            
            co0 = None
            co1 = None
            

            for con1 in conns1:
                for con0 in conns0:
                    if con0.IsConnectedTo(con1):
                        co0 = con0
                        co1 = con1

            o0 = co1.Origin

            l_Parameter = el1.LookupParameter('L')
            if l_Parameter == None:l_Parameter = el1.LookupParameter('l')
            if l_Parameter == None:l_Parameter = el1.LookupParameter('DIN_l')
            if l_Parameter == None:
                self.GUI.Close()
                text = 'Das Skript funktionert nicht mit Familie '+el1.Symbol.FamilyName+'. Übergangslänge kann nicht geändert werden. Bitte melden Sie sich bei Menghui.'
                TaskDialog.Show('Fehler',text)
                return

            t = DB.Transaction(doc,'Länge ändern')
            failureHandlingOptions = t.GetFailureHandlingOptions()
            failureHandler = FailureHandler()
            failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
            failureHandlingOptions.SetClearAfterRollback(True)
            t.SetFailureHandlingOptions(failureHandlingOptions)
            t.Start()

            try:
                l_Parameter.Set(int(self.GUI.length.Text)/304.8)
                doc.Regenerate()
                el1.Location.Move(o0-co1.Origin)
            except Exception as e:print(e)
        
            t.Commit()
            t.Dispose()

    def GetName(self):
        return "Übergansstück aktualisieren"