# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import math
from Autodesk.Revit.Exceptions import OperationCanceledException 
import clr

class ElementFilter(Selection.ISelectionFilter):
    def __init__(self,CategoryIds):
        self.cateids = CategoryIds
    def AllowElement(self,element):
        if element.Category.Id.IntegerValue in self.cateids:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False


class ElementFilterFactory:
    @staticmethod
    def CreateLuftkanalFormteilFilter():
        return ElementFilter([-2008010])
    
    @staticmethod
    def CreateLuftkanalFilter():
        return ElementFilter([-2008000])
    
    @staticmethod
    def CreateLuftkanalUndFlexFilter():
        return ElementFilter([-2008000,-2008020])
    
    @staticmethod
    def CreateLuftkanalFormteileUndZubehoerFilter():
        return ElementFilter([-2008016,-2008010])
    
    @staticmethod
    def CreateLuftteileFilter():
        return ElementFilter([-2008000,-2008016,-2008010,-2008013,-2001140])
    
    @staticmethod
    def CreateLuftauslassFilter():
        return ElementFilter([-2008013])
    
    @staticmethod
    def CreateHLSFilter():
        return ElementFilter([-2001140])
        




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

class ExternaleventListe(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.ExecuteApp = None
        self.name = ''
        
    def Execute(self,uiapp):
        try:
            self.ExecuteApp(uiapp)
        except Exception as e:
            TaskDialog.Show('Fehler',e.ToString())
    
    def GetName(self):
        return self.name
    
    def ErmittelnConns(self,elem0,elem1):
        try:
            conns0 = elem0.ConnectorManager.Connectors
        except:
            conns0 = elem0.MEPModel.ConnectorManager.Connectors
        try:
            conns1 = elem1.ConnectorManager.Connectors
        except:
            conns1 = elem1.MEPModel.ConnectorManager.Connectors
        distance = 10**8
        co0 = None
        co1 = None
        for con in conns0:
            for con1 in conns1:
                if con.IsConnected == False and con1.IsConnected == False:
                    dis = con.Origin.DistanceTo(con1.Origin)
                    if dis < distance:
                        distance = dis
                        co0 = con
                        co1 = con1
        return co0,co1

    def ConnectorPruefen(self,conn_0,conn_1):
        if conn_0.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn_1.CoordinateSystem.BasisZ.Negate()):
            return 'Übergang'
        else:
            if conn_0.CoordinateSystem.BasisZ.AngleTo(conn_1.CoordinateSystem.BasisZ) >= math.pi/4:
                return 'Bogen'
            else:
                return None
    
    def ErstellenUebergang(self,doc,conn_0,conn_1):
        try:
            fi = doc.Create.NewTransitionFitting(conn_0, conn_1)
            doc.Regenerate()
            return fi
        except:
            try:
                fi = doc.Create.NewTransitionFitting(conn_1, conn_0)
                doc.Regenerate()
                return fi
            except Exception as e:
                TaskDialog.Show('Fehler',e.ToString())
    
    def ErstellenBogen(self,doc,conn_0,conn_1):
        doc.Regenerate()
        try:
            fi = doc.Create.NewElbowFitting(conn_0, conn_1)
            doc.Regenerate()
            return fi
        except:
            try:
                fi = doc.Create.NewElbowFitting(conn_1, conn_0)
                doc.Regenerate()
                return fi
            except Exception as e:
                TaskDialog.Show('Fehler',e.ToString())

    def ErmittelnLaenge(self,conn0,conn1):
        o = conn0.Origin-conn1.Origin
        ln = conn0.CoordinateSystem.BasisZ
        return abs(o.DotProduct(ln))
    
    def ConnectorPruefen_Versprung(self,conn_0,conn_1):
        P0 = conn_0.Origin
        P1 = conn_1.Origin
        liste = [(P0.X-P1.X) <= 0.1**8, (P0.Y-P1.Y) <= 0.1**8, (P0.Z-P1.Z) <= 0.1**8]
        anzahl = 0
        for el in liste:
            if el == True:
                anzahl += 1
        if anzahl == 0:
            return 4 # 4 Bogen
        if anzahl == 1:
            return 2 # 2 Bogen
        return 0 # Fehler       
    
    def Luftkanalerstellen(self,P0,P1,Kanal):
        duct = Kanal.Document.GetElement(DB.ElementTransformUtils.CopyElement(Kanal.Document,Kanal.Id,DB.XYZ(0,0,0))[0])
        duct.Location.Curve = DB.Line.CreateBound(P0,P1)
        return duct
    
    # def Luftkanalerstellen(self,conn0,conn1,Kanal):
    #     duct = DB.Mechanical.Duct.Create(Kanal.Document,Kanal.DuctType.Id,Kanal.ReferenceLevel.Id,conn0,conn1)
    #     # duct.Location.Curve = DB.Line.CreateBound(P0,P1)
    #     return duct
    
    
    def ErmittelnConnSet(self,_conn,duct):
        conns = duct.ConnectorManager.Connectors
        for conn in conns:
            if conn.Origin.IsAlmostEqualTo(_conn.Origin):
                return conn
    
    def ErmittelnRotateWinkel(self,conn0,conn1,einstellungswinkel):
        winkel = conn0.CoordinateSystem.BasisZ.AngleTo(conn1.CoordinateSystem.BasisZ)
        return einstellungswinkel - winkel
    
    def ErmittelnRotateAchse(self,conn0,conn1,conn2):
        vektor = conn0.CoordinateSystem.BasisZ.CrossProduct(conn1.CoordinateSystem.BasisZ)
        P = (conn0.Origin + conn2.Origin) / 2
        return DB.Line.CreateUnbound(P,vektor)      

    def VorbereitungBogen(self,conn0,conn1):
        def ConnBearbeiten(conn):
            owner = conn.Owner
            conn1 = None
            if owner.Category.Id.IntegerValue == -2008000:
                conns = owner.ConnectorManager.Connectors
                for con in conns:
                    if con.Id != conn.Id:
                        conn1 = con
                conn.Origin = conn1.Origin + conn.CoordinateSystem.BasisZ/304.8*10
        ConnBearbeiten(conn0)
        ConnBearbeiten(conn1)
    
    def Bogenversprungerstellen(self,uiapp):
        self.name = 'Versprung erstellen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        self.GUI.IsEnabled = False

        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ElementFilterFactory.CreateLuftkanalFilter(),'Wählt den ersten Kanal aus')
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ElementFilterFactory.CreateLuftkanalFilter(),'Wählt den zweiten Kanal aus')
                el0 = doc.GetElement(el0_ref.ElementId)
                el1 = doc.GetElement(el1_ref.ElementId)
                conn0,conn1 = self.ErmittelnConns(el0,el1)                
                if not (conn0 and conn1):
                    TaskDialog.Show('Info','Zwei offene Anschlüsse nötig!')
                    break
                art = self.ConnectorPruefen_Versprung(conn0,conn1)
                if art == 0:
                    TaskDialog.Show('Info','Erstellen eines Bogenversprungs nicht möglich!')
                    break
                if art == 2:
                    P_Center = (conn0.Origin + conn1.Origin) / 2
                    winkel = 45
                    # if self.GUI.Winkel.SelectedIndex:
                    #     winkel = self.GUI.Winkel.SelectedItem
                    # else:
                    #     winkel = self.GUI.Winkel.Text
                    einstellwinkel = (180 - float(winkel)) * math.pi / 180
                    t = DB.Transaction(doc)
                    failureHandlingOptions = t.GetFailureHandlingOptions()
                    failureHandler = FailureHandler()
                    failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
                    failureHandlingOptions.SetClearAfterRollback(True)
                    t.SetFailureHandlingOptions(failureHandlingOptions)
                    t.Start('Bogenversprung erstellen')
                    duct = self.Luftkanalerstellen(conn0.Origin,conn1.Origin,el0)

                    doc.Regenerate()
                    _conn0 = self.ErmittelnConnSet(conn0,duct)
                    _conn1 = self.ErmittelnConnSet(conn1,duct)

                    rotatewinkel = self.ErmittelnRotateWinkel(conn0,_conn0,einstellwinkel)
                    if rotatewinkel != 0:
                        achse = self.ErmittelnRotateAchse(conn0,_conn0,conn1)
                        duct.Location.Rotate(achse,rotatewinkel)
                        try:
                            bogen = self.ErstellenBogen(doc,conn0,_conn0)
                            doc.Delete(bogen.Id)
                            self.ErstellenBogen(doc,conn0,_conn0)
                        except Exception as e:
                            print(e)
                        
                        try:
                            bogen = self.ErstellenBogen(doc,conn1,_conn1)
                            doc.Delete(bogen.Id)
                            self.ErstellenBogen(doc,conn1,_conn1)
                        except Exception as e:
                            print(e)
                    
                    t.Commit()
                    t.Dispose()

            except Exception as e:
                if e.ToString() != 'The user aborted the pick operation.':
                    print(e.ToString())
                break

        self.GUI.IsEnabled = True

    def Formteilerstellen(self,uiapp):
        self.name = 'Übergang erstellen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        self.GUI.IsEnabled = False

        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ElementFilterFactory.CreateLuftteileFilter(),'Wählt den ersten Luftkanalteil aus')
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,ElementFilterFactory.CreateLuftteileFilter(),'Wählt den zweiten Luftkanalteil aus')
                el0 = doc.GetElement(el0_ref.ElementId)
                el1 = doc.GetElement(el1_ref.ElementId)
                conn0,conn1 = self.ErmittelnConns(el0,el1)                
                if not (conn0 and conn1):
                    TaskDialog.Show('Info','Zwei offene Anschlüsse nötig!')
                    break
                art = self.ConnectorPruefen(conn0,conn1)
                if not art:
                    TaskDialog.Show('Info','Erstellen eines Bogens oder Übergangs nicht möglich!')
                    break
                o0,o1 = conn0.Origin,conn1.Origin
                t = DB.Transaction(doc)
                failureHandlingOptions = t.GetFailureHandlingOptions()
                failureHandler = FailureHandler()
                failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
                failureHandlingOptions.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(failureHandlingOptions)
                t.Start('Übergang erstellen')
                if art == 'Bogen':
                    bogen = self.ErstellenBogen(doc,conn0,conn1)
                    if bogen:
                        doc.Delete(bogen.Id)
                        bogen = self.ErstellenBogen(doc,conn0,conn1)

                    if not bogen:
                        break
                else:
                    length = self.ErmittelnLaenge(conn0, conn1)
                    ueber = self.ErstellenUebergang(doc,conn0,conn1)
                    doc.Regenerate()
                t.Commit()

                
                if art == 'Übergang':
                    if ueber:
                        cs = ueber.MEPModel.ConnectorManager.Connectors
                    else:
                        continue
                else:
                    continue

                co_fi = ''
                for con in cs:
                    if not co_fi:co_fi = con
                    else:
                        if co_fi.Id > con.Id:co_fi = con

                l_Parameter = ueber.LookupParameter('L')
                if l_Parameter == None:l_Parameter = ueber.LookupParameter('l')
                if l_Parameter == None:l_Parameter = ueber.LookupParameter('DIN_l')
                if l_Parameter == None:
                    text = 'Das Skript funktionert nicht mit Familie '+ueber.Symbol.FamilyName+'. Übergangslänge kann nicht geändert werden. Bitte melden Sie sich bei Menghui.'
                    TaskDialog.Show('Fehler',text)
                    self.GUI.Close()
                    return

                t.Start('Länge ändern')

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
                
                if self.GUI.erst.IsChecked:
                    ueber.Location.Move(o0-conn0.Origin)
                else:
                    ueber.Location.Move(o1-conn1.Origin)
                        
                t.Commit()
                t.Dispose()

            except Exception as e:
                print(e)
                break

        self.GUI.IsEnabled = True


class TransitionFitting(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        self.GUI.create.IsEnabled = False
        self.GUI.update.IsEnabled = False
        self.GUI.erst.IsEnabled = False
        self.GUI.zweit.IsEnabled = False
        self.GUI.auto.IsEnabled = False
        self.GUI.manuell.IsEnabled = False
        self.GUI.button_1.IsEnabled = False
        

        while(True):
            try:
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
                    self.GUI.create.IsEnabled = True
                    self.GUI.update.IsEnabled = True
                    self.GUI.erst.IsEnabled = True
                    self.GUI.zweit.IsEnabled = True
                    self.GUI.auto.IsEnabled = True
                    self.GUI.manuell.IsEnabled = True
                    self.GUI.button_1.IsEnabled = True
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
                    self.GUI.create.IsEnabled = True
                    self.GUI.update.IsEnabled = True
                    self.GUI.erst.IsEnabled = True
                    self.GUI.zweit.IsEnabled = True
                    self.GUI.auto.IsEnabled = True
                    self.GUI.manuell.IsEnabled = True
                    self.GUI.button_1.IsEnabled = True
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
                        self.GUI.create.IsEnabled = True
                        self.GUI.update.IsEnabled = True
                        self.GUI.erst.IsEnabled = True
                        self.GUI.zweit.IsEnabled = True
                        self.GUI.auto.IsEnabled = True
                        self.GUI.manuell.IsEnabled = True
                        self.GUI.button_1.IsEnabled = True
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
                    self.GUI.create.IsEnabled = True
                    self.GUI.update.IsEnabled = True
                    self.GUI.erst.IsEnabled = True
                    self.GUI.zweit.IsEnabled = True
                    self.GUI.auto.IsEnabled = True
                    self.GUI.manuell.IsEnabled = True
                    self.GUI.button_1.IsEnabled = True
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
                    self.GUI.create.IsEnabled = True
                    self.GUI.update.IsEnabled = True
                    self.GUI.erst.IsEnabled = True
                    self.GUI.zweit.IsEnabled = True
                    self.GUI.auto.IsEnabled = True
                    self.GUI.manuell.IsEnabled = True
                    self.GUI.button_1.IsEnabled = True
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
            except:
                self.GUI.create.IsEnabled = True
                self.GUI.update.IsEnabled = True
                self.GUI.erst.IsEnabled = True
                self.GUI.zweit.IsEnabled = True
                self.GUI.auto.IsEnabled = True
                self.GUI.manuell.IsEnabled = True
                self.GUI.button_1.IsEnabled = True
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