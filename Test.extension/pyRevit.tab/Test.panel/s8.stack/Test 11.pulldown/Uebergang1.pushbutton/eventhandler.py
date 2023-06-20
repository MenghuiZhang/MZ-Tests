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

class KanalundFlexFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008000' or element.Category.Id.ToString() == '-2008020':
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

        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element)
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element)
                el0 = doc.GetElement(el0_ref)
                el1 = doc.GetElement(el1_ref)
                try:conns0 = list(el0.ConnectorManager.Connectors)
                except:conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
                try:conns1 = list(el1.ConnectorManager.Connectors)
                except:conns1 = list(el1.MEPModel.ConnectorManager.Connectors)
                
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
      

                t = DB.Transaction(doc,'Übergangsstück erstellen')
                # failureHandlingOptions = t.GetFailureHandlingOptions()
                # failureHandler = FailureHandler()
                # failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
                # failureHandlingOptions.SetClearAfterRollback(True)
                # t.SetFailureHandlingOptions(failureHandlingOptions)
                t.Start()
                try:
                    fi = doc.Create.NewTransitionFitting(co1, co0)
                    doc.Regenerate()
                except:
                    try:
                        fi = doc.Create.NewTransitionFitting (co0, co1)
                        doc.Regenerate()
                    except Exception as e:
                        t.RollBack()
                        t.Dispose()
                        TaskDialog.Show('Fehler',e.ToString())
                        
                        return

                t.Commit()
                t.Dispose()

                
            except:

                return

    def GetName(self):
        return "Übergansstück erstellen"

class TransitionFitting1(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        self.GUI.create.IsEnabled = False
        self.GUI.update.IsEnabled = False
        self.GUI.luftkanal.IsEnabled = False
        self.GUI.formteil.IsEnabled = False
        self.GUI.length.IsEnabled = False
        self.GUI.button_2.IsEnabled = False
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        while(True):
            try:
                if self.GUI.luftkanal.IsChecked == True:
                    el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,KanalundFlexFilter(),'Wählt den fixierten Luftkanal/Flexkanal aus')
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
                        self.GUI.create.IsEnabled = True
                        self.GUI.update.IsEnabled = True
                        self.GUI.erst.IsEnabled = True
                        self.GUI.zweit.IsEnabled = True
                        self.GUI.auto.IsEnabled = True
                        self.GUI.manuell.IsEnabled = True
                        self.GUI.button_1.IsEnabled = True
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
                            DB.ElementTransformUtils.MoveElement(doc,el1.Id,o0-co1.Origin)
                            # el1.Location.Move(o0-co1.Origin)
                        except Exception as e:print(e)

                    else:
                        if l_origin + l > int(self.GUI.length.Text)/304.8:
                            try:
                                el1.Location.Move((int(self.GUI.length.Text)/304.8-l_origin)*ln)
                                l_Parameter.Set(int(self.GUI.length.Text)/304.8)
                                doc.Regenerate()
                                # el1.Location.Move(o0-co1.Origin)
                                DB.ElementTransformUtils.MoveElement(doc,el1.Id,o0-co1.Origin)
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
                        self.GUI.create.IsEnabled = True
                        self.GUI.update.IsEnabled = True
                        self.GUI.erst.IsEnabled = True
                        self.GUI.zweit.IsEnabled = True
                        self.GUI.auto.IsEnabled = True
                        self.GUI.manuell.IsEnabled = True
                        self.GUI.button_1.IsEnabled = True
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
                        # el1.Location.Move(o0-co1.Origin)
                        DB.ElementTransformUtils.MoveElement(doc,el1.Id,o0-co1.Origin)
                    except Exception as e:print(e)
                
                    t.Commit()
                    t.Dispose()
            except:
                self.GUI.create.IsEnabled = True
                self.GUI.update.IsEnabled = True
                self.GUI.luftkanal.IsEnabled = True
                self.GUI.formteil.IsEnabled = True
                self.GUI.length.IsEnabled = True
                self.GUI.button_2.IsEnabled = True
                return

    def GetName(self):
        return "Übergansstück aktualisieren"