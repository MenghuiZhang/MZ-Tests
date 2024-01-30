# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,ExternalEvent,TaskDialogCommonButtons,Selection
from rpw import db,DB
from IGF_Filters._selectionfilter import PickElementOptionFactory
from IGF_ExternalEventHandler._FailureHandler import MyFailuresPreprocessor
from IGF_Funktionen._Trassen import ErstellenBogen,ErstellenUebergang,IsHorizontalBogen
from IGF_Funktionen._Parameter import wert_schreibenbase
from IGF_Klasse._trassen import Trassen

class ExternalEventHandlerFactory:
    
    @staticmethod
    def AbstandEinstellen(uiapp,Abstand):
        """
        Abstand: TextBox
        """

        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elem0 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.IntegerValue in [-2008010,-2008016,-2001140,-2008049,-2008055],Text='Wählt das fixierte Luftkanalformteil/-zubehör/Rohrformteil/-zubehör/ HLS-Bauteile aus')

        if Abstand.Text == None or Abstand.Text == '':
            Abstand.Text = '10'
        try:
            abstand = float(Abstand.Text) / 304.8
        except:
            TaskDialog.Show('Fehler','ungültigen Abstand!')
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return

        if elem0.Category.Id.IntegerValue in [-2008010,-2008016]:
            elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.IntegerValue in [-2008010,-2008016],Text='Wählt das zu verschiebende Luftkanalformteil/-zubehör aus')
        elif elem0.Category.Id.IntegerValue in [-2008049,-2008055]:
            elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.IntegerValue in [-2008049,-2008055],Text='Wählt das zu verschiebende Roheformteil/-zubehör aus')
        else:
            elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.IntegerValue in [-2008010,-2008016,-2001140,-2008049,-2008055],Text='Wählt das fixierte Luftkanalformteil/-zubehör/Rohrformteil/-zubehör/ HLS-Bauteile aus')

        distance = None
        co0 = None
        co1 = None
        conns0 = list(elem0.MEPModel.ConnectorManager.Connectors)
        conns1 = list(elem1.MEPModel.ConnectorManager.Connectors)

        for con0 in conns0:
            for con1 in conns1:
                dis = con0.Origin.DistanceTo(con1.Origin)
                if not distance:
                    distance = dis
                    co0 = con0
                    co1 = con1
                elif dis < distance:
                    distance = dis
                    co0 = con0
                    co1 = con1
                    
        with db.Transaction('Abstand',doc):
            p_Neu = co0.Origin + DB.Line.CreateBound(co0.Origin,co1.Origin).Direction.Normalize() * abstand
            pinned = elem1.Pinned
            elem1.Pinned = False
            doc.Regenerate()
            elem1.Location.Move(p_Neu - co1.Origin)
            elem1.Pinned = pinned
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
    @staticmethod
    def VerbindungselementAnpassen(uiapp,Radiobutton_auto,Radiobutton_manuell):
        """
        Length: TextBox
        """
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if Radiobutton_manuell.Content.Text == None or Radiobutton_manuell.Content.Text == '':
            Radiobutton_manuell.Content.Text = '100'
        elem0 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.IntegerValue in [-2008000,-2008020,-2008016,-2008010,-2008055],Text='Wählt das fixierte Luftkanalformteil/-zubehör/Luftkanal/Flexkanal/HLS-Bauteile aus')
        try:
            conns0 = list(elem0.ConnectorManager.Connectors)
        except:
            conns0 = list(elem0.MEPModel.ConnectorManager.Connectors)
        Liste = []
        for conn in conns0:
            if conn.IsConnected:
                for ref in conn.AllRefs:
                    if ref.Owner.Category.Id.IntegerValue == -2008010:
                        if ref.Owner.MEPModel.PartType.ToString() == 'Transition':Liste.append(ref.Owner.Id.IntegerValue)
        if len(Liste) == 0:
            TaskDialog.Show('Kein Übergang damit verbunden!')
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return 
        elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Id.IntegerValue in Liste,Text='Wählt die Übergang aus')
        conns1 = list(elem1.MEPModel.ConnectorManager.Connectors)

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

        ln = DB.Line.CreateBound(co0.Origin,co0_.Origin).Direction.Normalize()

        o0 = co1.Origin
        
        l_Parameter = elem1.LookupParameter('DIN_l_soll')

        if l_Parameter == None or l_Parameter.IsReadOnly:
            text = 'Das Skript funktionert nicht mit Familie '+elem1.Symbol.FamilyName+'. Übergangslänge kann nicht geändert werden. Bitte melden Sie sich bei Menghui.'
            TaskDialog.Show('Fehler',text)       
            return

        l_origin = l_Parameter.AsDouble()

        with db.Transaction('Übergang anpassen',doc) as transaction:
            failureHandlingOptions = transaction.transaction.GetFailureHandlingOptions()
            failureHandler = MyFailuresPreprocessor()
            failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
            failureHandlingOptions.SetClearAfterRollback(True)
            transaction.SetFailureHandlingOptions(failureHandlingOptions)
            if co2.Id > co1.Id:
                try:
                    if Radiobutton_auto.IsChecked:
                        l_Parameter.Set(0)
                        doc.Regenerate()
                        DB.ElementTransformUtils.MoveElement(doc,elem1.Id,o0-co1.Origin)
                    else:

                        l_Parameter.Set(float(Radiobutton_manuell.Content.Text)/304.8)
                        doc.Regenerate()
                        DB.ElementTransformUtils.MoveElement(doc,elem1.Id,o0-co1.Origin)
                except:pass
            else:
                if elem0.Category.Id.ToString() == '-2008000':
                    if Radiobutton_auto.IsChecked:
                        l_Parameter.Set(0)
                        doc.Regenerate()
                        DB.ElementTransformUtils.MoveElement(doc,elem1.Id,o0-co1.Origin)
                    else:
                        l = elem0.LookupParameter('Länge').AsDouble()
                        if l_origin + l > float(Radiobutton_manuell.Content.Text)/304.8:
                            try:
                                l_Parameter.Set(float(Radiobutton_manuell.Content.Text)/304.8)
                                doc.Regenerate()
                                DB.ElementTransformUtils.MoveElement(doc,elem1.Id,o0-co1.Origin)
                            except:pass
                        else:
                            try:
                                elem1.Location.Move((float(Radiobutton_manuell.Content.Text)/304.8-l_origin)*ln)
                                l_Parameter.Set(float(Radiobutton_manuell.Content.Text)/304.8)
                                doc.Regenerate()
                                DB.ElementTransformUtils.MoveElement(doc,elem1.Id,o0-co1.Origin)
                            except:
                                TaskDialog.Show('Fehler','eingegebene Länge zu groß!')
              
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def ChangeHeightDuct(uiapp,ReferenzEbene,Hoeheart,Hoehe):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = [doc.GetElement(elemid) for elemid in uidoc.Selection.GetElementIds() if doc.GetElement(elemid).Category.Id.ToString() in ['-2008000']]
        if len(elems) != 1:
            TaskDialog.Show('Info','Bitte nur ein Luftkanal auswählen!')
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
        with db.Transaction('Höhe anpassen',doc):
            elems[0].get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM).Set(ReferenzEbene)
            elems[0].get_Parameter(Hoeheart).SetValueString(Hoehe)
            
        
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
    @staticmethod
    def ChangeDimensionDuct(uiapp,Breite = None,Hoehe= None,Durchmesser= None,InkFormteile = True):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if InkFormteile:
            elems = [doc.GetElement(elemid) for elemid in uidoc.Selection.GetElementIds() if doc.GetElement(elemid).Category.Id.ToString() in ['-2008000','-2008010']]
        else:
            elems = [doc.GetElement(elemid) for elemid in uidoc.Selection.GetElementIds() if doc.GetElement(elemid).Category.Id.ToString() in ['-2008000']]
        if len(elems) == 0:
            TaskDialog.Show('Info','Kein Luftkanal/Luftkanalformteile ausgewählt!')
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
        elemids = [e.Id.ToString() for e in elems]

        with db.Transaction('dimensionieren',doc):
            for el in elems:
                if el.Category.Id.ToString() == '-2008000':
                    try:
                        wert_schreibenbase(el.LookupParameter('Breite'),Breite)
                        wert_schreibenbase(el.LookupParameter('Höhe'),Hoehe)
                        wert_schreibenbase(el.LookupParameter('Durchmesser'),Durchmesser)
                                                    
                    except Exception as e:print(e)
                else:
                    conns = el.MEPModel.ConnectorManager.Connectors
                    isHorizontalBogen = IsHorizontalBogen(el)
                    for conn in conns:
                        refs = conn.AllRefs
                        for ref in refs:
                            if conns.Size > 2:
                                if ref.Owner.Id.ToString() not in elemids:continue
                            else:pass
                        mepinfo = conn.GetMEPConnectorInfo()
                        d = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_DIAMETER))
                        r = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_RADIUS))
                        h = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_HEIGHT))
                        b = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_WIDTH))
                        if isHorizontalBogen == True:
                            tmp = b
                            b = h
                            h = tmp
                            
                        if d != DB.ElementId.InvalidElementId:
                            wert_schreibenbase(el.get_Parameter(doc.GetElement(d).GetDefinition()), Durchmesser)

                        if r != DB.ElementId.InvalidElementId:
                            if Durchmesser:
                                try:wert_schreibenbase(el.get_Parameter(doc.GetElement(r).GetDefinition()), float(Durchmesser)/2)
                                except:pass
                        if h != DB.ElementId.InvalidElementId:
                            wert_schreibenbase(el.get_Parameter(doc.GetElement(h).GetDefinition()), Hoehe)
                        if b != DB.ElementId.InvalidElementId:
                            wert_schreibenbase(el.get_Parameter(doc.GetElement(b).GetDefinition()), Breite)


        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod 
    def CreateBeschriftungEinzeln(uiapp,Beschrifter,Beschrifter_RU = None,manuell = False,Linie = False):
        """
        Dict_Luftkanal: Luftkanal: 2 Beschrifter
        Beschrifter: Symbol Id
        """
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        
        BasisZ = None
        if uidoc.ActiveView.ViewType.ToString() in ['FloorPlan','CeilingPlan']:
            Range = uidoc.ActiveView.GetViewRange()
            try:
                l_u = doc.GetElement(Range.GetLevelId(DB.PlanViewPlane.CutPlane)).get_Parameter(DB.BuiltInParameter.LEVEL_ELEV).AsDouble()
                l_u_os = doc.GetElement(Range.GetOffset(DB.PlanViewPlane.CutPlane))
                BasisZ = l_u + l_u_os
            except:
                pass
            Range.Dispose()

        liste = []
        if manuell:
            elem = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.ToString() == '-2008000',Text='Wählt den Kanal aus')
            kanal = Trassen(elem)
        else:
            elem = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Category.Id.ToString() == '-2008000' and x.LookupParameter('Breite') != None,Text='Wählt den Kanal aus')
            kanal = Trassen(elem)
        if manuell:
            ref_2 = uidoc.Selection.PickPoint(Selection.ObjectSnapTypes.Points,'Wählt ein Einfügpunkt aus')
            if BasisZ:
                po = DB.XYZ(ref_2.X,ref_2.Y,BasisZ)
            else:
                po = ref_2
            with db.Transaction('Luftkanal beschriften',doc):
                if kanal.Shape == 'RE': kanal.CreatedBeschriftung_Point(Beschrifter, uidoc.ActiveView, po,Linie)
                else:kanal.CreatedBeschriftung_Point(Beschrifter_RU, uidoc.ActiveView, po,Linie)
        else:
            liste = []
            if kanal.StartLinkedElement:
                liste.append(kanal.StartLinkedElement.elemid)
            if kanal.EndLinkedElement:
                liste.append(kanal.EndLinkedElement.elemid)
            if len(liste) > 0:
                elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc, lambda x:x.Id.ToString() in liste,Text='Wählt den Luftkanalformteil/-zubehör aus')
                with db.Transaction('Luftkanal beschriften',doc):
                    if kanal.StartLinkedElement:
                        if elem1.Id.ToString() == kanal.StartLinkedElement.elemid:
                            kanal.CreatedBeschriftung_Start(Beschrifter, uidoc.ActiveView)
                        else:
                            kanal.CreatedBeschriftung_End(Beschrifter, uidoc.ActiveView)
                    else:
                        kanal.CreatedBeschriftung_End(Beschrifter, uidoc.ActiveView)
            else:
                with db.Transaction('Luftkanal beschriften',doc):
                    ref_2 = uidoc.Selection.PickPoint(Selection.ObjectSnapTypes.Points,'Wählt ein Einfügpunkt aus')
                    if kanal.StartConnector.Origin.DistanceTo(ref_2) <= kanal.EndConnector.Origin.DistanceTo(ref_2):
                        kanal.CreatedBeschriftung_Start(Beschrifter, uidoc.ActiveView)
                    else:
                        kanal.CreatedBeschriftung_End(Beschrifter, uidoc.ActiveView)
        
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
    @staticmethod
    def CreateBeschriftung(uiapp,beschrifter_re,beschrifter_ru,Dict_Luftkanal, Basisz = None):
        """
        Dict_Luftkanal: Luftkanal: 2 Beschrifter
        Beschrifter: Symbol Id
        """
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        
        with db.Transaction('Luftkanal beschriften',doc):
            Liste_Formteile = []
            Dict = {}
            for Luftkanal in Dict_Luftkanal.keys():
                Dict.setdefault(Dict_Luftkanal[Luftkanal],[])
                Dict[Dict_Luftkanal[Luftkanal]].append(Trassen(Luftkanal,Dict_Luftkanal[Luftkanal]))
            if 3 in Dict.keys():
                for Luftkanal in Dict[3]:
                    Luftkanal.CreatedBeschriftung_Start(beschrifter_re,uidoc.ActiveView,Basisz)
                    Luftkanal.CreatedBeschriftung_End(beschrifter_re,uidoc.ActiveView,Basisz)
                    Luftkanal.CreatedBeschriftung_Mitte(beschrifter_re,uidoc.ActiveView,Basisz)
                    if Luftkanal.StartLinkedElement:
                        Liste_Formteile.append(Luftkanal.StartLinkedElement.elemid)
                    if Luftkanal.EndLinkedElement:
                        Liste_Formteile.append(Luftkanal.EndLinkedElement.elemid)
            if 2 in Dict.keys():
                for Luftkanal in Dict[2]:
                    Luftkanal.CreatedBeschriftung_Start(beschrifter_re,uidoc.ActiveView,Basisz)
                    Luftkanal.CreatedBeschriftung_End(beschrifter_re,uidoc.ActiveView,Basisz)
                    if Luftkanal.StartLinkedElement:
                        Liste_Formteile.append(Luftkanal.StartLinkedElement.elemid)
                    if Luftkanal.EndLinkedElement:
                        Liste_Formteile.append(Luftkanal.EndLinkedElement.elemid)

            if 1 in Dict.keys():
                for Luftkanal in Dict[1]:
                    if Luftkanal.Shape == 'RU':
                        Luftkanal.CreatedBeschriftung_Mitte(beschrifter_ru,uidoc.ActiveView,Basisz)
                    else:
                        if not Luftkanal.StartLinkedElement:
                            Luftkanal.CreatedBeschriftung_Start(beschrifter_re,uidoc.ActiveView,Basisz)
                        else:
                            if Luftkanal.StartLinkedElement.Art == 'Formteil':
                                if Luftkanal.StartLinkedElement.IstBogen and Luftkanal.StartLinkedElement.IstAsymBogen == False:
                                    if Luftkanal.StartLinkedElement.elemid not in Liste_Formteile:
                                        Liste_Formteile.append(Luftkanal.StartLinkedElement.elemid)
                                        Luftkanal.CreatedBeschriftung_Start(beschrifter_re,uidoc.ActiveView,Basisz)
                                else:
                                    Luftkanal.CreatedBeschriftung_Start(beschrifter_re,uidoc.ActiveView,Basisz)
                            else:
                                Luftkanal.CreatedBeschriftung_Start(beschrifter_re,uidoc.ActiveView,Basisz)

                        if not Luftkanal.EndLinkedElement:
                            Luftkanal.CreatedBeschriftung_End(beschrifter_re,uidoc.ActiveView,Basisz)
                        else:
                            if Luftkanal.EndLinkedElement.Art == 'Formteil':
                                if Luftkanal.EndLinkedElement.IstBogen and Luftkanal.EndLinkedElement.IstAsymBogen == False:
                                    if Luftkanal.EndLinkedElement.elemid not in Liste_Formteile:
                                        Liste_Formteile.append(Luftkanal.EndLinkedElement.elemid)
                                        Luftkanal.CreatedBeschriftung_End(beschrifter_re,uidoc.ActiveView,Basisz)
                                else:
                                    Luftkanal.CreatedBeschriftung_End(beschrifter_re,uidoc.ActiveView,Basisz)
                            else:
                                Luftkanal.CreatedBeschriftung_End(beschrifter_re,uidoc.ActiveView,Basisz)

        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
