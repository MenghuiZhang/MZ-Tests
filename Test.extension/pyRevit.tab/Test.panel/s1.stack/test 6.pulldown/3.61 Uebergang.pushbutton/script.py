# coding: utf8
import Autodesk.Revit.DB as DB
from Autodesk.Revit.UI import TaskDialog


__title__ = "Übergang1"
__doc__ = """

[2022.04.13]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

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

def funktion():
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document
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
    print((l0.Direction+l1.Direction).GetLength(), (l0.Direction-l1.Direction).GetLength())
    if not ((l0.Direction+l1.Direction).GetLength() < 1e-10 or (l0.Direction-l1.Direction).GetLength() < 1e-10):
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
    o = co0.Origin-co1.Origin
    ln = l0.Direction.Normalize()
    length = o.DotProduct(ln)
    t = DB.Transaction(doc,'Übergangsstückerstellen')
    # failureHandlingOptions = t.GetFailureHandlingOptions()
    # failureHandler = FailureHandler()
    # failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
    # failureHandlingOptions.SetClearAfterRollback(True)
    # t.SetFailureHandlingOptions(failureHandlingOptions)
    t.Start()
    try:
        fi = doc.Create.NewTransitionFitting(co0, co1)

    except:
        fi = doc.Create.NewTransitionFitting(co1, co0)
        # try:
        #     fi = doc.Create.NewTransitionFitting(co1, co0)
        # except Exception as e:
        #     print(e)
        #     t.RollBack()
        #     return
    doc.Regenerate()
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
    t.Dispose()


def funktion1():
    uidoc = __revit__.ActiveUIDocument
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
            
            elem1 = DB.Mechanical.Duct.Create(doc,el0.LookupParameter('Systemtyp').AsElementId(),el0.GetTypeId(),el0.ReferenceLevel.Id ,conns0[0].Origin,conns0[1].Origin)
            elem2 = DB.Mechanical.Duct.Create(doc,el1.LookupParameter('Systemtyp').AsElementId(),el1.GetTypeId(),el1.ReferenceLevel.Id ,conns1[1].Origin,conns1[0].Origin)
            elem1.LookupParameter('Breite').Set(conns0[0].Width)
            elem1.LookupParameter('Höhe').Set(conns0[0].Height)
            elem2.LookupParameter('Breite').Set(conns1[0].Width)
            elem2.LookupParameter('Höhe').Set(conns1[0].Height)
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
                    # con.Width = co0.Width
                    # con.Height = co0.Height
            for con in conns1_temp :
                print(1)
                # con.Width = co1.Width
                # con.Height = co1.Height
                if (con.Origin-co1.Origin).GetLength() < 0.00001:
                    co1_temp  = con
                    # con.Width = co1.Width
                    # con.Height = co1.Height
            
            try:
                fi = doc.Create.NewTransitionFitting(co1, co0_temp)
            except:
                fi = doc.Create.NewTransitionFitting(co0_temp, co1)
                try:
                    fi = doc.Create.NewTransitionFitting(co0_temp, co1)
                except Exception as e:
                    print(e)
                    print('.....')
                    t.RollBack()
                    return
            # except Exception as e:
            #     print('.......')
            #     print(e)
            #     t.RollBack()
            #     return
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
funktion()