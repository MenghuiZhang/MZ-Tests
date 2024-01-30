# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import math
from System.Collections.Generic import List
from IGF_ExternalEventHandler._ExternalEventHandler_NotReadOnly import ExternalEventHandlerFactory,ExternalEvent,TaskDialog,db,MyFailuresPreprocessor
from IGF_ExternalEventHandler._ExternalEventHandler_ReadOnly import ExternalEventHandlerFactory_READONLY
from IGF_ExternalEventHandler import ItemTemplateEventHandler
from IGF_Filters._selectionfilter import PickElementOptionFactory
from IGF_Klasse._ebene import EbeneFactory
from IGF_Klasse._trassen import TrassenFactory
from IGF_Klasse._familie_typ import FamilieTypFactory
from IGF_Klasse._ansicht import Grundriss
from IGF_Funktionen._Parameter import get_value
from IGF_Funktionen._Trassen import Volumenstrom_berechnen,Tempo_berechnen,Dimension_berechnen
from System import EventHandler, Uri
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs, ViewActivatingEventArgs
from Autodesk.Revit.DB.Events import DocumentOpenedEventArgs 
from IGF_Klasse._familyparameter import FamilyParameterFactory
from System.Windows.Forms import OpenFileDialog,DialogResult, SaveFileDialog
from IGF_Funktionen._Allgemein import RVTMainWindowActive

IS_Ebenen = EbeneFactory.GetAllEbenen()
IS_DuctType = TrassenFactory.GetAllLuftkanalType()
IS_PipeType = TrassenFactory.GetAllRohrType()
IS_FamilieType = FamilieTypFactory.GetAllModellFamilienType()

class ExternaleventListe(ItemTemplateEventHandler):
    def __init__(self):
        ItemTemplateEventHandler.__init__(self)
        self.failureHandler = MyFailuresPreprocessor()
        self.documentopenedeventhandler = EventHandler[DocumentOpenedEventArgs](self.documentopened)
        self.viewchangedeventhandler = EventHandler[ViewActivatingEventArgs](self.viewchanged)
        self.viewchangedstatus = False
        self.documentopenedstatus = True
        self.familydoc = ''
    
    def VisibilityEinstellen(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if doc.IsFamilyDocument is False:
            self.GUI.TI_Projekt.IsSelected = True
            self.GUI.TI_Familie.IsEnabled = False
            self.GUI.TI_Projekt.IsEnabled = True
        else:
            if self.familydoc != doc.Title:
                self.familydoc = doc.Title
                self.GUI.IsEnabled = False
                self._FamilyParameterAktualiesieren(uiapp)
                self.GUI.IsEnabled = True
            self.GUI.TI_Familie.IsSelected = True
            self.GUI.TI_Projekt.IsEnabled = False
            self.GUI.TI_Familie.IsEnabled = True
    
    def FamilyParameterAktualiesieren(self,uiapp):
        self.name = 'Parameter aktualisieren'
        if uiapp.ActiveUIDocument.Document.IsFamilyDocument is False:
            return
        if not uiapp.Application.SharedParametersFilename:
            TaskDialog.Show('Info','keine SharedParameterdatei gefunden')
            return
        self.familydoc = uiapp.ActiveUIDocument.Document.Title
        self.GUI.IsEnabled = False
        self._FamilyParameterAktualiesieren(uiapp)
        self.GUI.IsEnabled = True  
    
    def _FamilyParameterAktualiesieren(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _file = uiapp.Application.OpenSharedParameterFile()
        self.GUI.din_Params.ItemsSource = FamilyParameterFactory.GetDINParameter(_file)
        self.GUI.lin_Params.ItemsSource = FamilyParameterFactory.GetLINParameter(_file)
        self.GUI.mc_eklimax_Params.ItemsSource = FamilyParameterFactory.GetMCEKLIMAXParameter(_file)
        self.GUI.mc_Params.ItemsSource = FamilyParameterFactory.GetMCParameter(_file)
        self.GUI.igf_Params.ItemsSource = FamilyParameterFactory.GetIGFParameter(_file)
        self.GUI.igf_ga_Params.ItemsSource = FamilyParameterFactory.GetIGFGAParameter(_file)
        uidoc.Dispose()
        doc.Dispose()
    
    def FamilyParameterErstellen(self,uiapp):
        self.name = 'FamilyParameter Erstellen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if doc.IsFamilyDocument is False:
            return
        def FamilyParamErstellen(Liste):
            for el in Liste:
                if el.checked:el.AddParameter(doc)
        with db.Transaction(self.name,doc):
            FamilyParamErstellen(self.GUI.din_Params.Items)
            FamilyParamErstellen(self.GUI.lin_Params.Items)
            FamilyParamErstellen(self.GUI.mc_eklimax_Params.Items)
            FamilyParamErstellen(self.GUI.mc_Params.Items)
            FamilyParamErstellen(self.GUI.igf_Params.Items)
            FamilyParamErstellen(self.GUI.igf_ga_Params.Items)
        uidoc.Dispose()
        doc.Dispose()
        self.FamilyParameterAktualiesieren(uiapp)
    
    def documentopened(self,sender,e):
        doc = e.NewActiveView.Document
        if doc.IsFamilyDocument is False:
            return
        if not doc.Application.SharedParametersFilename:
            TaskDialog.Show('Info','keine SharedParameterdatei gefunden')
            return
        _file = doc.Application.OpenSharedParameterFile()
        self.GUI.din_Params.ItemsSource = FamilyParameterFactory.GetDINParameter(_file)
        # self.GUI.lin_Params.ItemsSource = FamilyParameterFactory.GetLINParameter(_file)
        # self.GUI.mc_eklimax_Params.ItemsSource = FamilyParameterFactory.GetMCEKLIMAXParameter(_file)
        # self.GUI.mc_Params.ItemsSource = FamilyParameterFactory.GetMCParameter(_file)
        # self.GUI.igf_Params.ItemsSource = FamilyParameterFactory.GetIGFParameter(_file)
        # self.GUI.igf_ga_Params.ItemsSource = FamilyParameterFactory.GetIGFGAParameter(_file)
        doc.Dispose()

        # self.ExcuteApp = self.FamilyParameterAktualiesieren
        # self.GUI.externalevent.Raise()
    
    def DocOpened(self,uiapp):
        uiapp.Application.DocumentOpened += self.documentopenedeventhandler
    
    def SharedParameterdurchsuchen(self,uiapp):
        self.name = 'durchsuchen'
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "Shared Parameter"
        dialog.Filter = "Txt Dateien|*.txt"
        if dialog.ShowDialog() == DialogResult.OK:
            self.GUI.sharedparameter.Text = dialog.FileName
            self.GUI.IsEnabled = False
            try:
                uiapp.Application.SharedParametersFilename = dialog.FileName
                _file = uiapp.Application.OpenSharedParameterFile()
                self.GUI.din_Params.ItemsSource = FamilyParameterFactory.GetDINParameter(_file)
                self.GUI.lin_Params.ItemsSource = FamilyParameterFactory.GetLINParameter(_file)
                self.GUI.mc_eklimax_Params.ItemsSource = FamilyParameterFactory.GetMCEKLIMAXParameter(_file)
                self.GUI.mc_Params.ItemsSource = FamilyParameterFactory.GetMCParameter(_file)
                self.GUI.igf_Params.ItemsSource = FamilyParameterFactory.GetIGFParameter(_file)
                self.GUI.igf_ga_Params.ItemsSource = FamilyParameterFactory.GetIGFGAParameter(_file)
            except:pass
            self.GUI.IsEnabled = True

    def viewchanged(self,sender,e):
        doc = e.NewActiveView.Document
        if doc.IsFamilyDocument is False:
            self.GUI.TI_Projekt.IsSelected = True
            self.GUI.TI_Familie.IsEnabled = False
            self.GUI.TI_Projekt.IsEnabled = True
            view = e.NewActiveView
            if view.ViewType.ToString() not in ['FloorPlan','CeilingPlan']:
                self.GUI.Expander_Workschnitte.IsEnabled = False
            else:
                self.GUI.Expander_Workschnitte.IsEnabled = True
                ausgewaehlt = [el.name for el in self.GUI.ListeSchnitte if el.checked]
                self.GUI.SchnitteGrundriss.elem = view
                self.GUI.ListeSchnitte = self.GUI.SchnitteGrundriss.getAllSchnitt()
                for el in self.GUI.ListeSchnitte:
                    if el.name in ausgewaehlt:
                        el.checked = True
                self.GUI.dg_Schnitte.ItemsSource = self.GUI.ListeSchnitte
        else:
            if self.familydoc != doc.Title:
                self.familydoc = doc.Title
                self.GUI.sharedparameter.Text = doc.Application.SharedParametersFilename
                self.GUI.IsEnabled = False
                try:
                    _file = doc.Application.OpenSharedParameterFile()
                    self.GUI.din_Params.ItemsSource = FamilyParameterFactory.GetDINParameter(_file)
                    self.GUI.lin_Params.ItemsSource = FamilyParameterFactory.GetLINParameter(_file)
                    self.GUI.mc_eklimax_Params.ItemsSource = FamilyParameterFactory.GetMCEKLIMAXParameter(_file)
                    self.GUI.mc_Params.ItemsSource = FamilyParameterFactory.GetMCParameter(_file)
                    self.GUI.igf_Params.ItemsSource = FamilyParameterFactory.GetIGFParameter(_file)
                    self.GUI.igf_ga_Params.ItemsSource = FamilyParameterFactory.GetIGFGAParameter(_file)
                except:pass
                self.GUI.IsEnabled = True
            self.GUI.TI_Familie.IsSelected = True
            self.GUI.TI_Projekt.IsEnabled = False
            self.GUI.TI_Familie.IsEnabled = True
    
    def AktiveViewActivated(self,uiapp):
        uiapp.ViewActivating += self.viewchangedeventhandler

    def ElemAnzeigen(self,uiapp):
        self.name = 'Element anzeigen'
        elemid = self.GUI.TB_ElementId.Text
        ExternalEventHandlerFactory_READONLY.ElementAnzeigenElementId(uiapp,elemid)
    
    def ElemAuswaehlen(self,uiapp):
        self.name = 'Element auswählen'
        elemid = self.GUI.TB_ElementId.Text
        ExternalEventHandlerFactory_READONLY.ElementAuswaehlenElementId(uiapp,elemid)
    
    def RaumAnzeigen(self,uiapp):
        self.name = 'Raum anzeigen'
        raumnr = self.GUI.TB_Raumnr.Text
        ExternalEventHandlerFactory_READONLY.RaumanzeigenRaumnummer(uiapp,raumnr)
    
    def RaumAuswaehlen(self,uiapp):
        self.name = 'Raum auswählen'
        raumnr = self.GUI.TB_Raumnr.Text
        ExternalEventHandlerFactory_READONLY.RaumAuswaehlenRaumnummer(uiapp,raumnr)
 
    def AbstandAnpassen(self,uiapp):
        self.name = 'Abstand anpassen'
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        while(True):
            try:
                ExternalEventHandlerFactory.AbstandEinstellen(uiapp, self.GUI.TextBox_Abstand)
            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True
    
    def UebergangAnpassen(self,uiapp):
        self.name = 'Übergang anpassen'
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        while(True):
            try:
                ExternalEventHandlerFactory.VerbindungselementAnpassen(uiapp, self.GUI.Radio_auto,self.GUI.Radio_manuell)
            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True
    
    def Verbindungerstellen(self,uiapp):
        self.name = 'Verbindungselement erstellen'
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        self.winkel_pruefen()
        while(True):
            try:
                try:elem0 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:x.Category.Id.IntegerValue in [-2008000,-2008016,-2008010,-2008013,-2001140] and (x.ConnectorManager.UnusedConnectors if x.Category.Id.IntegerValue in [-2008000] else x.MEPModel.ConnectorManager.UnusedConnectors).Size != 0,Text='Wählt den ersten Luftkanalteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                try:elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:x.Category.Id.IntegerValue in [-2008000,-2008016,-2008010,-2008013,-2001140] and (x.ConnectorManager.UnusedConnectors if x.Category.Id.IntegerValue in [-2008000] else x.MEPModel.ConnectorManager.UnusedConnectors).Size != 0 and x.Id.IntegerValue != elem0.Id.IntegerValue,Text='Wählt den zweiten Luftkanalteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                if self.GUI.uebergang.IsChecked:
                    self.Formteilerstellen(uiapp,elem0,elem1)
                else:
                    self.Bogenversprungerstellen(uiapp,elem0,elem1)

            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True
    
    def winkel_pruefen(self):
        if not self.GUI.uebergang.IsChecked:
            if self.GUI.ComboBox_winkel.SelectedIndex != -1:
                winkel = float(self.GUI.ComboBox_winkel.SelectedItem)
            else:
                winkel = float(self.GUI.ComboBox_winkel.Text)
            if winkel > 90:
                winkel = 90
                self.GUI.ComboBox_winkel.SelectedItem = '90'

    def ErmittelnConns(self,elem0,elem1):
        try:      conns0 = elem0.ConnectorManager.Connectors
        except:   conns0 = elem0.MEPModel.ConnectorManager.Connectors
        try:      conns1 = elem1.ConnectorManager.Connectors
        except:   conns1 = elem1.MEPModel.ConnectorManager.Connectors
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
        if conn_0.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn_1.CoordinateSystem.BasisZ.Negate()): return 'Übergang'
        else:
            if conn_0.CoordinateSystem.BasisZ.AngleTo(conn_1.CoordinateSystem.BasisZ) > 89.6/180.0*math.pi: return 'Bogen'
            else: return 'Fehler'
    
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
        return abs((conn0.Origin-conn1.Origin).DotProduct(conn0.CoordinateSystem.BasisZ))
    
    def ConnectorPruefen_Versprung(self,conn_0,conn_1):
        # if not conn_0.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn_1.CoordinateSystem.BasisZ.Negate()): return 0
        
        P0 = conn_0.Origin
        P1 = conn_1.Origin
        liste = [abs(P0.X-P1.X) <= 0.1**8, abs(P0.Y-P1.Y) <= 0.1**8, abs(P0.Z-P1.Z) <= 0.1**8]
        anzahl = 0
        for el in liste:
            if el == True:  anzahl += 1
        if anzahl == 0:   return 4 # 4 Bogen
        if anzahl == 1:  return 2 # 2 Bogen
        if anzahl == 2:  return 2 # 2 Bogen
        return 0 # Fehler       
    
    def Luftkanalerstellen(self,P0,P1,Kanal):
        duct = Kanal.Document.GetElement(DB.ElementTransformUtils.CopyElement(Kanal.Document,Kanal.Id,DB.XYZ(0,0,0))[0])
        duct.Location.Curve = DB.Line.CreateBound(P0,P1)
        return duct
        
    def Luftkanalerstellen_HorizontalVertikal(self,conn0,conn1,Kanal):
        """Horizontal: conn0 gehörrt zu elem0
           Vertikal: conn0 gehörrt zu elem1"""
        duct = Kanal.Document.GetElement(DB.ElementTransformUtils.CopyElement(Kanal.Document,Kanal.Id,DB.XYZ(0,0,0))[0])
        richtung = conn1.CoordinateSystem.BasisZ
        lange = self.ErmittelnLaenge(conn0, conn1)
        if lange < 100/304.8:
            lange = 100/304.8
        p0 = conn1.Origin + (lange-10/304.8) * richtung         
        p1 = conn1.Origin + 10/304.8 * richtung    
        duct.Location.Curve = DB.Line.CreateBound(DB.XYZ(p0.X,p0.Y,conn0.Origin.Z),DB.XYZ(p1.X,p1.Y,conn0.Origin.Z))
        return duct
    
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
    
    def GetNewlyCreatedElement(self,elem0,elem1,elem2):
        Liste = [elem2.Id]
        for conn in elem2.MEPModel.ConnectorManager.Connectors:
            if conn.IsConnected:
                for ref in conn.AllRefs:
                    if ref.Owner.Category.Id.ToString() in ['-2008000','-2008016','-2008010','-2008013','-2001140'] and ref.Owner.Id.ToString() not in [elem0.Id.ToString(),elem1.Id.ToString()]:
                        Liste.append(ref.Owner.Id)

        return Liste
    
    def Bogenversprungerstellen_Zwei(self,uiapp,elem0,elem1):
        self.name = 'Bogensversprung erstellen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        conn0,conn1 = self.ErmittelnConns(elem0,elem1)                
        if not (conn0 and conn1):
            TaskDialog.Show('Info','Zwei offene Anschlüsse nötig!')
            return
        art = self.ConnectorPruefen_Versprung(conn0,conn1)
        if art == 0:
            TaskDialog.Show('Info','Erstellen eines Bogenversprungs nicht möglich!')
            return
        if art == 2:
            P_Center = (conn0.Origin + conn1.Origin) / 2
            if self.GUI.ComboBox_winkel.SelectedIndex != -1:
                winkel = float(self.GUI.ComboBox_winkel.SelectedItem)
            else:
                winkel = float(self.GUI.ComboBox_winkel.Text)
            if winkel > 90:
                winkel = 90
                self.GUI.ComboBox_winkel.SelectedItem = '90'
            einstellwinkel = (180 - float(winkel)) * math.pi /180
            with db.TransactionGroup(self.name,True,doc) as tg:
                with db.Transaction('Luftkanal erstellen',doc) as t:
                    failureHandlingOptions = t.GetFailureHandlingOptions()
                    failureHandlingOptions.SetFailuresPreprocessor(self.failureHandler)
                    failureHandlingOptions.SetClearAfterRollback(True)
                    t.SetFailureHandlingOptions(failureHandlingOptions)
                    duct = self.Luftkanalerstellen(conn0.Origin,conn1.Origin,elem0)
                    doc.Regenerate()
                    _conn0 = self.ErmittelnConnSet(conn0,duct)
                    _conn1 = self.ErmittelnConnSet(conn1,duct)

                    rotatewinkel = self.ErmittelnRotateWinkel(conn0,_conn0,einstellwinkel)
                    if rotatewinkel != 0:
                        achse = self.ErmittelnRotateAchse(conn0,_conn0,conn1)
                        duct.Location.Rotate(achse,rotatewinkel)
                    doc.Regenerate()
                
                self.Formteilerstellen(uiapp, elem0, duct)
                self.Formteilerstellen(uiapp, duct, elem1)

        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
        
    def Bogenversprungerstellen(self,uiapp,elem0,elem1):
        self.name = 'Bogensversprung erstellen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        conn0,conn1 = self.ErmittelnConns(elem0,elem1)                
        if not (conn0 and conn1):
            TaskDialog.Show('Info','Zwei offene Anschlüsse nötig!')
            return
        art = self.ConnectorPruefen_Versprung(conn0,conn1)
        if art == 0:
            TaskDialog.Show('Info','Erstellen eines Bogenversprungs nicht möglich!')
            return
        if art == 2:
            self.Bogenversprungerstellen_Zwei(uiapp,elem0,elem1)
        
        elif art == 4:
            with db.TransactionGroup(self.name,True,doc) as tg:
                with db.Transaction('Luftkanal erstellen',doc) as t1:
                    failureHandlingOptions = t1.GetFailureHandlingOptions()
                    failureHandlingOptions.SetFailuresPreprocessor(self.failureHandler)
                    failureHandlingOptions.SetClearAfterRollback(True)
                    t1.SetFailureHandlingOptions(failureHandlingOptions)
                    if self.GUI.radio_horizontal.IsChecked:
                        duct = self.Luftkanalerstellen_HorizontalVertikal(conn0,conn1,elem1)
                    else:
                        duct = self.Luftkanalerstellen_HorizontalVertikal(conn1,conn0,elem0)
                    doc.Regenerate()

                self.Bogenversprungerstellen_Zwei(uiapp,elem0,duct)
                self.Bogenversprungerstellen_Zwei(uiapp,duct,elem1)
    
    def Formteilerstellen(self,uiapp,elem0,elem1):
        self.name = 'Bogen/Übergang'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        conn0,conn1 = self.ErmittelnConns(elem0,elem1)  
        if not (conn0 and conn1):
            TaskDialog.Show('Info','Zwei offene Anschlüsse nötig!')
            return
                
        art = self.ConnectorPruefen(conn0,conn1)
        if not art:
            TaskDialog.Show('Info','Erstellen eines Bogens oder Übergangs nicht möglich!')
            return
        o0,o1 = conn0.Origin,conn1.Origin
        with db.TransactionGroup(self.name,True,doc) as tg:
            with db.Transaction(self.name,doc) as t:
                failureHandlingOptions = t.GetFailureHandlingOptions()
                failureHandlingOptions.SetFailuresPreprocessor(self.failureHandler)
                failureHandlingOptions.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(failureHandlingOptions)
                if art == 'Bogen':
                    bogen = self.ErstellenBogen(doc,conn0,conn1)
                    if bogen:
                        doc.Delete(List[DB.ElementId](self.GetNewlyCreatedElement(elem0, elem1, bogen)))
                        bogen = self.ErstellenBogen(doc,conn0,conn1)

                    if not bogen:
                        return
                    
                else:
                    ueber = self.ErstellenUebergang(doc,conn0,conn1)
                    doc.Regenerate()
            
            if art == 'Bogen':
                with db.Transaction('Länge ändern',doc) as t:
                    failureHandlingOptions = t.GetFailureHandlingOptions()
                    failureHandlingOptions.SetFailuresPreprocessor(self.failureHandler)
                    failureHandlingOptions.SetClearAfterRollback(True)
                    t.SetFailureHandlingOptions(failureHandlingOptions)
                    bogen.LookupParameter('DIN_b').Set(bogen.LookupParameter('DIN_d').AsDouble())

            if art == 'Übergang':
                if ueber:
                    cs = ueber.MEPModel.ConnectorManager.Connectors
                else:
                    return
            else:
                return

            co_fi = ''
            for con in cs:
                if not co_fi:co_fi = con
                else:
                    if co_fi.Id > con.Id:co_fi = con

            l_Parameter = ueber.LookupParameter('DIN_l_soll')

            if l_Parameter == None or l_Parameter.IsReadOnly:
                text = 'Das Skript funktionert nicht mit Familie '+ueber.Symbol.FamilyName+'. Übergangslänge kann nicht geändert werden. Bitte melden Sie sich bei Menghui.'
                TaskDialog.Show('Fehler',text)       
                return
            
            with db.Transaction('Länge ändern',doc) as t:
                failureHandlingOptions = t.GetFailureHandlingOptions()
                failureHandlingOptions.SetFailuresPreprocessor(self.failureHandler)
                failureHandlingOptions.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(failureHandlingOptions)
                if self.GUI.auto.IsChecked:
                    try:
                        l_Parameter.Set(0)
                        doc.Regenerate()
                    except Exception as e:print(e)
                else:
                    try:
                        l_Parameter.Set(int(self.GUI.manuell.Content.Text)/304.8)
                        doc.Regenerate()
                    except:pass
                ueber.Location.Move(o0-conn0.Origin)

        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    def ChangeDimensionDuct(self,uiapp):
        self.name = 'dimensionieren'
        ExternalEventHandlerFactory.ChangeDimensionDuct(uiapp,self.GUI.TB_Breite.Text,self.GUI.TB_Hoehe.Text,self.GUI.TB_Durchmesser.Text,self.GUI.CB_InkKanalfitting.IsChecked)
        self.KanaldatenAusRevitNurEbenen(uiapp)

    def ChangeHeightDuct(self,uiapp):
        self.name = 'Höhe anpassen'
        if self.GUI.CB_Reference.SelectedIndex != -1:
            ReferenzEbene = self.GUI.CB_Reference.SelectedItem.elem.Id
            if self.GUI.RB_Oben.IsChecked:
                height = self.GUI.TB_Oben.Text
                Hoeheart = DB.BuiltInParameter.RBS_DUCT_TOP_ELEVATION
            elif self.GUI.RB_Mitte.IsChecked:
                height = self.GUI.TB_Mitte.Text
                Hoeheart = DB.BuiltInParameter.RBS_OFFSET_PARAM
            else:
                height = self.GUI.TB_Unten.Text
                Hoeheart = DB.BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION
            if height:
                ExternalEventHandlerFactory.ChangeHeightDuct(uiapp, ReferenzEbene, Hoeheart, height)
        self.KanaldatenAusRevitNurEbenen(uiapp)
    
    def KanaldatenAusRevit(self,uiapp):
        self.name = 'GUI aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = [doc.GetElement(elemid) for elemid in uidoc.Selection.GetElementIds() if doc.GetElement(elemid).Category.Id.ToString() in ['-2008000']]
        if len(elems) == 0:
            TaskDialog.Show('Info','Kein Luftkanal ausgewählt!')
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
        el = elems[0]
        try:
            b_p = el.get_Parameter(DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            h_p = el.get_Parameter(DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            d_p = el.get_Parameter(DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            mh_p = el.get_Parameter(DB.BuiltInParameter.RBS_OFFSET_PARAM)
            oh_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_TOP_ELEVATION)
            uh_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION)
            re_p = el.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
            v_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM)
            t_p = el.get_Parameter(DB.BuiltInParameter.RBS_VELOCITY)
            if b_p:self.GUI.TB_Breite.Text = str(int(round(get_value(b_p),1)))
            if h_p:self.GUI.TB_Hoehe.Text = str(int(round(get_value(h_p),1)))
            if d_p:self.GUI.TB_Durchmesser.Text = str(int(round(get_value(d_p),1)))
            if oh_p:self.GUI.TB_Oben.Text = str(int(round(get_value(oh_p),1)))
            if mh_p:self.GUI.TB_Mitte.Text = str(int(round(get_value(mh_p),1)))
            if uh_p:self.GUI.TB_Unten.Text = str(int(round(get_value(uh_p),1)))
            if v_p:self.GUI.TB_Volumenstrom.Text = str(int(round(get_value(v_p),1)))
            if t_p:self.GUI.TB_Tempo.Text = str(int(round(get_value(t_p),1)))
            if re_p:
                ebene = re_p.AsValueString()
                for el in self.GUI.IS_Ebenen:
                    if el.name == ebene:
                        break
                self.GUI.CB_Reference.SelectedItem = el
        except Exception as e:print(e)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
    def KanaldatenAusRevitNurEbenen(self,uiapp):
        self.name = 'GUI aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = [doc.GetElement(elemid) for elemid in uidoc.Selection.GetElementIds() if doc.GetElement(elemid).Category.Id.ToString() in ['-2008000']]
        if len(elems) == 0:
            TaskDialog.Show('Info','Kein Luftkanal ausgewählt!')
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
        el = elems[0]
        try:
            mh_p = el.get_Parameter(DB.BuiltInParameter.RBS_OFFSET_PARAM)
            oh_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_TOP_ELEVATION)
            uh_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION)
            re_p = el.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
            if oh_p:self.GUI.TB_Oben.Text = str(int(round(get_value(oh_p),1)))
            if mh_p:self.GUI.TB_Mitte.Text = str(int(round(get_value(mh_p),1)))
            if uh_p:self.GUI.TB_Unten.Text = str(int(round(get_value(uh_p),1)))
            if re_p:
                ebene = re_p.AsValueString()
                for el in self.GUI.IS_Ebenen:
                    if el.name == ebene:
                        break
                self.GUI.CB_Reference.SelectedItem = el
        except Exception as e:print(e)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    def Volumenstrom_berechnen(self,tempo = 0,b = None,h = None,d = None):
        return Volumenstrom_berechnen(tempo,b,h,d)
    
    def Tempo_berechnen(self,vol = 0,b = None,h = None,d = None):
        return Tempo_berechnen(vol,b,h,d)
    
    def Dimension_berechnen(self,vol = 0,tempo = 0, b = None,h = None,d = None):
        return Dimension_berechnen(vol,tempo,b,h,d)

    def get_all_Ebene(self):
        self.GUI.IS_Ebenen = EbeneFactory.GetAllEbenen()
        self.GUI.CB_Reference.ItemsSource = self.GUI.IS_Ebenen
    
    def get_all_FamilieTyp(self):
        self.GUI.IS_FamilieType = FamilieTypFactory.GetAllModellFamilienType()
        self.GUI.CB_FamilieType.ItemsSource = self.GUI.IS_FamilieType
    
    def get_all_TrassenTyp(self):
        self.GUI.IS_DuctType = TrassenFactory.GetAllLuftkanalType()
        self.GUI.CB_DuctType.ItemsSource = self.GUI.IS_DuctType
        self.GUI.IS_PipeType = TrassenFactory.GetAllRohrType()
        self.GUI.CB_PipeType.ItemsSource = self.GUI.IS_PipeType
    
    def set_up(self,uiapp):
        try:self.get_all_Ebene()
        except:pass
        try:self.get_all_TrassenTyp()
        except:pass
        try:self.get_all_FamilieTyp()
        except:pass
        try:
            view = uiapp.ActiveUIDocument.ActiveView
            if view.ViewType.ToString() not in ['FloorPlan','CeilingPlan']:
                self.GUI.Expander_Workschnitte.IsEnabled = False
            else:
                self.GUI.Expander_Workschnitte.IsEnabled = True
                ausgewaehlt = [el.name for el in self.GUI.ListeSchnitte if el.checked] if self.GUI.ListeSchnitte.Count > 0 else []
                self.GUI.SchnitteGrundriss.elem = view
                self.GUI.ListeSchnitte = self.GUI.SchnitteGrundriss.getAllSchnitt()
                for el in self.GUI.ListeSchnitte:
                    if el.name in ausgewaehlt:
                        el.checked = True
                self.GUI.dg_Schnitte.ItemsSource = self.GUI.ListeSchnitte
        except:pass
        if not self.viewchangedstatus:  
            self.AktiveViewActivated(uiapp)
            self.viewchangedstatus = True
        if not self.documentopenedstatus:  
            try:self.DocOpened(uiapp)
            except Exception as e:print(e)
            self.documentopenedstatus = True
    
    def ElementVerbinden(self,uiapp):
        self.name = 'Elemente verbinden'
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        while(True):
            try:
                try:elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:self.UnUsedConnectorsPruefen(x),Text='Wählt den ersten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                try:elem2 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:self.UnUsedConnectorsPruefen(x) and x.Id.ToString() != elem1.Id.ToString(),Text='Wählt den zweiten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                conn0,conn1 = self.GetparterConnector(elem1,elem2)
                if not self.connsrichtung_pruefen(conn0, conn1):
                    TaskDialog.Show('Fehler',"Die Anschlusssrichtung passt nicht")
                    continue
                with db.Transaction(self.name,uiapp.ActiveUIDocument.Document):
                    elem2.Location.Move(conn0.Origin-conn1.Origin)
                    uiapp.ActiveUIDocument.Document.Regenerate()
                    conn0.ConnectTo(conn1)

            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True

    def UnUsedConnectorsPruefen(self,elem):
        try:
            return elem.ConnectorManager.UnusedConnectors.Size > 0
        except:
            try:
                return elem.MEPModel.ConnectorManager.UnusedConnectors.Size > 0
            except:return False
        return False
    
    def GetparterConnector(self,elem1,elem2):
        try:conns0 = elem1.ConnectorManager.Connectors
        except:conns0 = elem1.MEPModel.ConnectorManager.Connectors
        try:conns1 = elem2.ConnectorManager.Connectors
        except:conns1 = elem2.MEPModel.ConnectorManager.Connectors

        Distance = 10000
        conn_0 = None
        conn_1 = None
        for conn1 in conns1:
            if not conn1.IsConnected:
                for conn0 in conns0:
                    if not conn0.IsConnected:
                        distance = conn0.Origin.DistanceTo(conn1.Origin)
                        if distance < Distance:
                            Distance = distance
                            conn_0 = conn0
                            conn_1 = conn1

        return conn_0,conn_1
    
    def connsrichtung_pruefen(self,conn0,conn1):
        if conn1.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn0.CoordinateSystem.BasisZ) or conn1.CoordinateSystem.BasisZ.IsAlmostEqualTo(conn0.CoordinateSystem.BasisZ.Negate()):
            return True
        else:
            return False

    def ElementTrennen(self,uiapp):
        self.name = 'Elemente trennen'
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        while(True):
            try:
                try:elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:self.UsedConnectorsPruefen(x),Text='Wählt den ersten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return 
                try:conns1 = elem1.ConnectorManager.Connectors
                except:conns1 = elem1.MEPModel.ConnectorManager.Connectors
                liste = []
                for conn in conns1:
                    if conn.IsConnected:
                        for ref in conn.AllRefs:
                            if ref.Owner.Category.Id.ToString() not in ['-2008123','-2008124','-2008122','-2008015','-2008043']:
                                liste.append(ref.Owner.Id.ToString())
                try:elem2 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:x.Id.ToString() in liste,Text='Wählt den zweiten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return 
                ConnectorSet = []
                for conn in conns1:
                    if conn.IsConnected:
                        for ref in conn.AllRefs:
                            if ref.Owner.Id.ToString() == elem2.Id.ToString():
                                ConnectorSet.append([conn,ref])
                
                if len(ConnectorSet) == 0:
                    TaskDialog.Show('Fehler',"Keine Verbindung gefunden")
                    continue
                with db.Transaction(self.name,uiapp.ActiveUIDocument.Document):
                    for conns in ConnectorSet:
                        conns[0].DisconnectFrom(conns[1])

            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True 

    def UsedConnectorsPruefen(self,elem):
        try:
            return elem.ConnectorManager.Connectors.Size > elem.ConnectorManager.UnusedConnectors.Size
        except:
            try:
                return elem.MEPModel.ConnectorManager.Connectors.Size > elem.MEPModel.ConnectorManager.UnusedConnectors.Size
            except:return False
        return False 
    
    def KanalstueckEinfuegen(self,uiapp):
        self.name = 'Kanal einfügen'
        if not (self.GUI.TB_KanalLaenge.Text and self.GUI.CB_DuctType.SelectedIndex != -1):
            TaskDialog.Show('Fehler','Länge nicht angegeben oder Typ nicht ausgewählt!')
            return
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        while(True):
            try:
                try:elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:self.UsedConnectorsPruefen(x) and x.Category.Id.ToString() in ['-2008016','-2008010','-2001140'] and len(self.VerbindenPruefen(x)) > 0,Text='Wählt den ersten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                liste = self.VerbindenPruefen(elem1)
                try:elem2 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:x.Id.ToString() in liste,Text='Wählt den zweiten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                levelid = elem1.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId()
                conns1 = elem1.MEPModel.ConnectorManager.Connectors
                ConnectorSet = []
                for conn in conns1:
                    if conn.IsConnected:
                        for ref in conn.AllRefs:
                            if ref.Owner.Id.ToString() == elem2.Id.ToString():
                                ConnectorSet.append([conn,ref])
                
                if len(ConnectorSet) == 0:
                    TaskDialog.Show('Fehler',"Keine Verbindung gefunden")
                    continue
                with db.Transaction(self.name,uiapp.ActiveUIDocument.Document):
                    for conns in ConnectorSet:
                        conn0 = conns[0]
                        conn1 = conns[1]
                        bearbeitungsbereich = elem1.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
                        uiapp.ActiveUIDocument.Document.GetWorksetTable().SetActiveWorksetId(DB.WorksetId(bearbeitungsbereich))
                        conn0.DisconnectFrom(conn1)
                        elem2.Location.Move(conn0.CoordinateSystem.BasisZ * float(self.GUI.TB_KanalLaenge.Text) / 304.8)
                        DB.Mechanical.Duct.Create(uiapp.ActiveUIDocument.Document,self.GUI.CB_DuctType.SelectedItem.elemid,levelid,conn0,conn1)

            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True 
    
    def RohrstueckEinfuegen(self,uiapp):
        self.name = 'Rohr einfügen'
        if not (self.GUI.TB_RohrLaenge.Text and self.GUI.CB_PipeType.SelectedIndex != -1):
            TaskDialog.Show('Fehler','Länge nicht angegeben oder Typ nicht ausgewählt!')
            return
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        while(True):
            try:
                try:elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:self.UsedConnectorsPruefen(x) and x.Category.Id.ToString() in ['-2008055','-2008049','-2001140'] and len(self.VerbindenPruefen(x)) > 0,Text='Wählt den ersten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                liste = self.VerbindenPruefen(elem1)
                try:elem2 = PickElementOptionFactory.CreateCurrentDocumentOption(uiapp.ActiveUIDocument, lambda x:x.Id.ToString() in liste,Text='Wählt den zweiten Bauteil aus')
                except:
                    self.GUI.IsEnabled = True
                    return
                levelid = elem1.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId()
                conns1 = elem1.MEPModel.ConnectorManager.Connectors
                ConnectorSet = []
                for conn in conns1:
                    if conn.IsConnected:
                        for ref in conn.AllRefs:
                            if ref.Owner.Id.ToString() == elem2.Id.ToString():
                                ConnectorSet.append([conn,ref])
                
                if len(ConnectorSet) == 0:
                    TaskDialog.Show('Fehler',"Keine Verbindung gefunden")
                    continue
                with db.Transaction(self.name,uiapp.ActiveUIDocument.Document):
                    for conns in ConnectorSet:
                        conn0 = conns[0]
                        conn1 = conns[1]
                        bearbeitungsbereich = elem1.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
                        uiapp.ActiveUIDocument.Document.GetWorksetTable().SetActiveWorksetId(DB.WorksetId(bearbeitungsbereich))
                        conn0.DisconnectFrom(conn1)
                        elem2.Location.Move(conn0.CoordinateSystem.BasisZ * float(self.GUI.TB_RohrLaenge.Text) / 304.8)
                        DB.Plumbing.Pipe.Create(uiapp.ActiveUIDocument.Document,self.GUI.CB_PipeType.SelectedItem.elemid,levelid,conn0,conn1)

            except:
                self.GUI.IsEnabled = True
                return
        self.GUI.IsEnabled = True 

    def VerbindenPruefen(self,elem):
        liste = []
        conns1 = elem1.MEPModel.ConnectorManager.Connectors
        for conn in conns1:
            if conn.IsConnected:
                for ref in conn.AllRefs:
                    if ref.Owner.Category.Id.ToString() not in ['-2008123','-2008124','-2008122','-2008015','-2008043','-2008000','-2008020','-2008050','-2008044']:
                        liste.append(ref.Owner.Id.ToString())
        return liste
    
    def SchnitteVerschieben(self,uiapp):
        self.name = 'Schnitte verschieben'
        schnitte = [el for el in self.GUI.ListeSchnitte if el.checked]
        if len(schnitte) == 0:
            TaskDialog.Show('Fehler','kein Schnitt ausgewählt!')
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        for elid in uidoc.Selection.GetElementIds():
            try:
                elem = doc.GetElement(elid)
                category = elem.Category
                if category.CategoryType.ToString() == 'Model':
                    elems.append(elem)
            except:pass
        if len(elems) == 0:
            TaskDialog.Show('Fehler','kein Element ausgewählt!')
            uidoc.Dispose()
            doc.Dispose()
            return
        
        with db.Transaction(self.name,doc):
            for schnitt in schnitte:
                if self.GUI.Schnittetiefe.IsChecked:
                    schnitt.verschieben_AnsichtsTiefe(elems)
                else:
                    schnitt.verschieben(elems)
        doc.Dispose()
        uidoc.Dispose()
            
    def ReferenzFamilieAktualisieren(self,uiapp):
        self.name = 'Referenztyp aktualisieren'

        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if uidoc.Selection.GetElementIds().Count == 1:
            elem = doc.GetElement(uidoc.Selection.GetElementIds()[0])
            if elem.Category.CategoryType.ToString() == 'Model' and isinstance(elem,DB.FamilyInstance):
                name = elem.Symbol.FamilyName + ': '+elem.Name
                for el in self.GUI.IS_FamilieType:
                    if el.familytypname == name:
                        self.GUI.CB_FamilieType.SelectedItem = el
                        uidoc.Dispose()
                        doc.Dispose()
                        return
        self.GUI.IsEnabled = False
        RVTMainWindowActive()
        try:
            elem = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:isinstance(x,DB.FamilyInstance),Text='Wählt den Famlietyp aus')
        except:
            self.GUI.IsEnabled = True
            return
        self.GUI.IsEnabled = True
        name = elem.Symbol.FamilyName + ': '+elem.Name
        for el in self.GUI.IS_FamilieType:
            if el.familytypname == name:
                self.GUI.CB_FamilieType.SelectedItem = el
                uidoc.Dispose()
                doc.Dispose()
                return
    
    def ChangeFamilietyp(self,uiapp):
        self.name = 'Familietyp ändern'
        uidoc = uiapp.ActiveUIDocument
        if uidoc.Selection.GetElementIds().Count == 0:
            return
        doc = uidoc.Document
        elemid = self.GUI.CB_FamilieType.SelectedItem.elem.Id
        elems = []
        for elid in uidoc.Selection.GetElementIds():
            elem = doc.GetElement(elid)
            if elem.Category.CategoryType.ToString() == 'Model' and isinstance(elem,DB.FamilyInstance):
                elems.append(elem)
        if len(elems) == 0:
            uidoc.Dispose()
            doc.Dispose()
            return
        
        message = 'Element'
        with db.Transaction(self.name,doc):
            for elem in elems:
                if elemid in elem.GetValidTypes():
                    try:elem.ChangeTypeId(elemid)
                    except:message += ' {},'.format(elem.Id.ToString())
                else:
                    message += ' {},'.format(elem.Id.ToString())
        uidoc.Dispose()
        doc.Dispose()
        if message != 'Element':
            message = message[:-1] + ' nicht angepasst'
            TaskDialog.Show('Fehler',message)
        


