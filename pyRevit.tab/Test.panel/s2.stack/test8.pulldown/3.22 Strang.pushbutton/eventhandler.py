# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from pyrevit import script
from System.Windows.Media import Brushes
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Windows import Visibility
from System.Windows.Input import Key
import System


Visible = Visibility.Visible
Hidden = Visibility.Hidden

Schwarz = Brushes.Black
Rot = Brushes.Red
doc = __revit__.ActiveUIDocument.Document
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(number+name+'Sprinkler-Strang dimensionieren')

class FilterNeben(Selection.ISelectionFilter):
    def __init__(self,liste):
        self.Liste = liste

    def AllowElement(self,element):
        if element.Id.ToString() in self.Liste:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter(Selection.ISelectionFilter):

    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class AlleBauteile:
    def __init__(self,Liste,elem):
        self.elem = elem
        self.rohrliste = Liste
        self.liste = []
        self.liste1 = []
        self.rohr = []
        self.formteil = []
        self.t_stueck = None
        for el in self.rohrliste:
            self.liste.append(el)
            self.liste1.append(el)
            self.rohr.append(el)
        
        if self.elem.Category.Id.ToString() == '-2008049':self.formteil.append(self.elem.Id)       
        self.get_t_st(self.elem)
        self.get_t_st_1(self.elem)

    def get_t_st(self,elem):       
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:                    
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() not in self.liste:  
                                if owner.Category.Id.ToString() == '-2001140':
                                    return     
                                elif owner.Category.Id.ToString() == '-2008044':
                                    self.rohr.append(owner.Id.ToString())
                                elif owner.Category.Id.ToString() == '-2008049':
                                    self.formteil.append(owner.Id)               
                                self.get_t_st(owner)
    
    def get_t_st_1(self,elem):
        elemid = elem.Id.ToString()
        self.liste1.append(elemid)
        if self.t_stueck:return
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and cate == 'Rohrformteile':
                    self.t_stueck = elem.Id
                    return
                    
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste1:  
                            if owner.Category.Id.ToString() == '-2008044':
                                if self.t_stueck == None:
                                    self.rohrliste.append(owner.Id.ToString())                         
                            self.get_t_st_1(owner)

class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

class TemplateItem(TemplateItemBase):
    def __init__(self,dimension=None,von=None,bis=None):
        TemplateItemBase.__init__(self)
        self._von = von
        self._bis = bis
        self._dimension = dimension
        self._farbe_bis = Schwarz
        self._farbe_von = Schwarz
        self.durchmesser = ['20','25','32','40','50','65','80','100','125','150','200','250','300','350','400','450','500','600']
        # self.PropertyChanged = PropertyChangedEventHandler
    
    @property
    def von(self):
        return self._von
    @von.setter
    def von(self,value):
        if value != self._von:
            self._von = value
            self.RaisePropertyChanged('von')
    
    @property
    def bis(self):
        return self._bis
    @bis.setter
    def bis(self,value):
        if value != self._bis:
            self._bis = value
            self.RaisePropertyChanged('bis')
    
    @property
    def dimension(self):
        return self._dimension
    @dimension.setter
    def dimension(self,value):
        if value != self._dimension:
            self._dimension = value
            self.RaisePropertyChanged('dimension')
    
    @property
    def farbe_bis(self):
        return self._farbe_bis
    @farbe_bis.setter
    def farbe_bis(self,value):
        if value != self._farbe_bis:
            self._farbe_bis = value
            self.RaisePropertyChanged('farbe_bis')
    
    @property
    def farbe_von(self):
        return self._farbe_von
    @farbe_von.setter
    def farbe_von(self,value):
        if value != self._farbe_von:
            self._farbe_von = value
            self.RaisePropertyChanged('farbe_von')

class TPiece:
    def __init__(self,Liste_Rohre,elem,doc = None,logger = None,dict_dimension_Neu = None,List_dimension = None,grunddimension = None):
        self.doc = doc
        self.logger = logger
        self.dict_dimension_Neu = dict_dimension_Neu
        self.List_dimension = List_dimension
        self.angepasst = False
        self.liste = []
        self.Liste_Rohre_1 = []
        self.Liste_Rohre_2 = []
        self.t_dict = {}
        self.Anzahl_Endverbraucher = 0
        self.T_Piece_0 = None
        self.T_Piece_1 = None
        self.Endverbraucher_0 = None
        self.Endverbraucher_1 = None
        self.Dimension = 0
        self.grunddimension = grunddimension

        self.elem = elem
        self.elemid = self.elem.Id
        self.Liste_Pre_Rohre = Liste_Rohre
        
        self.get_Liste_T_Piece(self.elem)
        if self.Endverbraucher_0 not in [None,'Pass']:self.Anzahl_Endverbraucher += 1
        if self.Endverbraucher_1 not in [None,'Pass']:self.Anzahl_Endverbraucher += 1
        if self.Endverbraucher_1 and self.Endverbraucher_0:
            self.angepasst = True 
    
    def get_Liste_T_Piece(self,elem):
        if (self.T_Piece_0 or self.Endverbraucher_0) and (self.T_Piece_1 or self.Endverbraucher_1) :return
        
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and cate == 'Rohrformteile' and elemid != self.elemid.ToString():
                    if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                        self.T_Piece_0 = elem.Id.ToString()
                    else:
                        self.T_Piece_1 = elem.Id.ToString()
                    return

                for conn in conns:
                    if not conn.IsConnected:
                        if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                            self.Endverbraucher_0 = 'Pass'
                        else:self.Endverbraucher_1 = 'Pass'
                        continue
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste and owner.Id.ToString() not in self.Liste_Pre_Rohre:  
                            if owner.Category.Id.ToString() == '-2008044':
                                if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                                    self.Liste_Rohre_1.append(owner.Id.ToString())
                                else:
                                    self.Liste_Rohre_2.append(owner.Id.ToString())
                            elif owner.Category.Id.ToString() in ['-2008099','-2001160','-2001140']:
                                if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                                    self.Endverbraucher_0 = owner.Id.ToString()
                                else:self.Endverbraucher_1 = owner.Id.ToString()
                                return
                           
                            self.get_Liste_T_Piece(owner)

    def get_Liste_Verbraucher(self):
        if self.T_Piece_0 in self.t_dict.keys():
            if self.t_dict[self.T_Piece_0].angepasst == False:
                self.logger.error(self.t_dict[self.T_Piece_0].elemid.ToString())
            self.Anzahl_Endverbraucher += self.t_dict[self.T_Piece_0].Anzahl_Endverbraucher
        if self.T_Piece_1 in self.t_dict.keys():
            if self.t_dict[self.T_Piece_1].angepasst == False:
                self.logger.error(self.t_dict[self.T_Piece_1].elemid.ToString())
            self.Anzahl_Endverbraucher += self.t_dict[self.T_Piece_1].Anzahl_Endverbraucher
 
    def get_Dimension(self):
        for n in self.List_dimension:
            if self.Anzahl_Endverbraucher >= n:
                self.Dimension = self.dict_dimension_Neu[n]
                break

    def wert_schreiben(self):
        if self.Endverbraucher_0 != None:
            for rohr in self.Liste_Rohre_1:
                try:
                    self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.grunddimension))
                except Exception as e:
                    self.logger.error(e)
        if self.Endverbraucher_1 != None:
            for rohr in self.Liste_Rohre_2:
                try:
                    self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.grunddimension))
                except Exception as e:
                    self.logger.error(e)
        
        for rohr in self.Liste_Pre_Rohre:
            try:
                self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.Dimension))
            except Exception as e:
                self.logger.error(e)

class Rohr:
    def __init__(self,elemid,doc):
        self.elemid = DB.ElementId(int(elemid))
        self.doc = doc
        self.elem = self.doc.GetElement(self.elemid)
        self.durchmesser = self.elem.LookupParameter('IGF_X_SM_Durchmesser').AsValueString()
    def wert_schreiben(self):
        self.elem.LookupParameter('Durchmesser').SetValueString(self.durchmesser)

class Rohrformteil:
    def __init__(self,elem):
        self.doc = None
        self.elem = elem
        self.conns0 = ''
        self.conns1 = ''
        self.conns2 = ''
        self.liste = []
        self.art = ''
        self.get_conns()
    
    def get_conns(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        if conns.Size == 2:
            conns = list(conns)
            conn_sizes = [int(conn.Radius*304.8*2) for conn in conns]
            if conn_sizes[0] != conn_sizes[1]:
                self.art = 'Übergang'
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Category.Id.ToString() in ['-2008044','-2008055']:
                                try:
                                    conns_owner = owner.ConnectorManager.Connectors
                                except:
                                    conns_owner = owner.MEPModel.ConnectorManager.Connectors
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp):
                                        if self.conns0 == '':
                                            self.conns0 = conn_temp
                                        else:
                                            self.conns1 = conn_temp
            else:
                self.art = 'Bogen'
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            elemid = owner.Id.ToString()
                            if elemid in self.liste:
                                continue
                            self.liste.append(elemid)
                            if owner.Category.Id.ToString() == '-2008044':
                                conns_owner = owner.ConnectorManager.Connectors
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp):
                                        if self.conns0 == '':
                                            self.conns0 = conn_temp
                                        else:
                                            self.conns1 = conn_temp
                            elif owner.Category.Id.ToString() == '-2008049':
                                conns_owner = owner.MEPModel.ConnectorManager.Connectors
                                conn_ander = ''
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp) == False:
                                        conn_ander = conn_temp
                                        break 
                                allrefs2 = conn_ander.AllRefs
                                for ref2 in allrefs2:
                                    owner2 = ref2.Owner
                                    if owner2.Category.Id.ToString() == '-2008044':
                                        conns_owner2 = owner2.ConnectorManager.Connectors
                                        for conn_temp2 in conns_owner2:
                                            if conn_ander.IsConnectedTo(conn_temp2):
                                                if self.conns0 == '':
                                                    self.conns0 = conn_temp2
                                                else:
                                                    self.conns1 = conn_temp2
        else:
            self.art = 'T-Stück'
            for conn in conns:
                if conn.IsConnected:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        elemid = owner.Id.ToString()
                        if elemid in self.liste:
                            continue
                        self.liste.append(elemid)
                        if owner.Category.Id.ToString() == '-2008044':
                            conns_owner = owner.ConnectorManager.Connectors
                            for conn_temp in conns_owner:
                                if conn.IsConnectedTo(conn_temp):
                                    if self.conns0 == '':
                                        self.conns0 = conn_temp
                                    elif self.conns1 == '':
                                        self.conns1 = conn_temp
                                    else:
                                        self.conns2 = conn_temp
                        elif owner.Category.Id.ToString() == '-2008049':
                            conns_owner = owner.MEPModel.ConnectorManager.Connectors
                            conn_ander = ''
                            for conn_temp in conns_owner:
                                if conn.IsConnectedTo(conn_temp) == False:
                                    conn_ander = conn_temp
                                    break 
                            allrefs2 = conn_ander.AllRefs
                            for ref2 in allrefs2:
                                owner2 = ref2.Owner
                                if owner2.Category.Id.ToString() == '-2008044':
                                    conns_owner2 = owner2.ConnectorManager.Connectors
                                    for conn_temp2 in conns_owner2:
                                        if conn_ander.IsConnectedTo(conn_temp2):
                                            if self.conns0 == '':
                                                self.conns0 = conn_temp2
                                            elif self.conns1 == '':
                                                self.conns1 = conn_temp2
                                            else:
                                                self.conns2 = conn_temp2
        

    def createfittings(self):
        if self.art == 'Bogen':
            try:self.doc.Create.NewElbowFitting(self.conns0, self.conns1)
            except:self.doc.Create.NewElbowFitting(self.conns1, self.conns0)
        elif self.art == 'Übergang':
            try:self.doc.Create.NewTransitionFitting(self.conns0, self.conns1)
            except:self.doc.Create.NewTransitionFitting(self.conns1, self.conns0)
        elif self.art == 'T-Stück':
            try:self.doc.Create.NewTeeFitting(self.conns0, self.conns1,self.conns2)
            except:
                try:self.doc.Create.NewTeeFitting(self.conns2, self.conns0, self.conns1)
                except:self.doc.Create.NewTeeFitting(self.conns1, self.conns2, self.conns0)

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        self.GUI.write_config()
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter(),'Wählt ein Rohr aus')
        el0 = doc.GetElement(el0_ref)
        conns = el0.ConnectorManager.Connectors
        liste = []
        Liste_Rohr = []
        Liste_Rohr.append(el0.Id.ToString())
        for conn in conns:
            if conn.IsConnected:
                refs = conn.AllRefs
                for _ref in refs:
                    if _ref.Owner.Category.Name not in ['Rohr Systeme','Rohrdämmung'] and _ref.Owner.Id.ToString() != el0.Id.ToString():
                        liste.append(_ref.Owner.Id.ToString())

        if len(liste) ==  0:
            TaskDialog.Show('Info','kein Teil damit verbunden!')
        ref_1 = uidoc.Selection.PickObject(Selection.ObjectType.Element,FilterNeben(liste),'Wählt den nächsten Teil aus')
        self.GUI.allebauteile = AlleBauteile(Liste_Rohr,doc.GetElement(ref_1))
        self.GUI.button_dimension.IsEnabled = True
        self.GUI.button_change.IsEnabled = True
    
    def GetName(self):
        return "auswählen"

class DIMENSIONIEREN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        try:
            self.GUI.write_config()
            self.GUI.button_dimension.IsEnabled = False
            self.GUI.button_change.IsEnabled = False
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            logger = script.get_logger()
            dict_dimension = {int(el[0]):el[2] for el in self.GUI.config.Einstellungen}
            liste_dimension = dict_dimension.keys()[:]
            liste_dimension.sort(reverse=True)
            grunddimension = dict_dimension[liste_dimension[-1]]

            self.GUI.pb_t.Visibility = Visible
            self.GUI.pb_c.Visibility = Visible

            t_dict = {}

            if self.GUI.allebauteile.t_stueck == None:
                t = DB.Transaction(doc,'Dimension')
                t.Start()
                for elid in self.GUI.allebauteile.rohrliste:
                    try:
                        doc.GetElement(DB.ElementId(int(elid))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(grunddimension))
                    except Exception as e:
                        self.logger.error(e)
                t.Commit()
                t.Dispose()
                return

            elem = doc.GetElement(DB.ElementId(int(self.GUI.allebauteile.t_stueck.ToString())))
            b = TPiece(self.GUI.allebauteile.rohrliste,elem,doc,logger,dict_dimension,liste_dimension,grunddimension)
            t_dict[self.GUI.allebauteile.t_stueck.ToString()] = b
            Liste_T_Stuecke = []
            _Liste_T_Stuecke = []
            _Liste_T_Stuecke.append(self.GUI.allebauteile.t_stueck.ToString())
            if b.T_Piece_0 != None:Liste_T_Stuecke.append([b.T_Piece_0,b.Liste_Rohre_1])
            if b.T_Piece_1 != None:Liste_T_Stuecke.append([b.T_Piece_1,b.Liste_Rohre_2])
            while(len(Liste_T_Stuecke) > 0):
                t_temp = []
                for el in Liste_T_Stuecke:
                    elem = doc.GetElement(DB.ElementId(int(el[0])))
                    b = TPiece(el[1],elem,doc,logger,dict_dimension,liste_dimension,grunddimension)
                    if b.T_Piece_0 != None:t_temp.append([b.T_Piece_0,b.Liste_Rohre_1])
                    if b.T_Piece_1 != None:t_temp.append([b.T_Piece_1,b.Liste_Rohre_2])
                    t_dict[el[0]] = b
                    _Liste_T_Stuecke.insert(0,el[0])
                Liste_T_Stuecke = t_temp
            for el in _Liste_T_Stuecke:
                elem = t_dict[el]
                elem.t_dict = t_dict
                elem.get_Liste_Verbraucher()
                elem.get_Dimension()
                elem.angepasst = True

            self.GUI.maxvalue = len(t_dict.values())

            t = DB.Transaction(doc,'Dimension')
            t.Start()
            for n,el in enumerate(t_dict.values()):
                self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                System.Windows.Threading.DispatcherPriority.Background)
                self.GUI.value = n+1
                self.GUI.PB_text = str(n+1) + ' / '+ str(len(t_dict.values())) + ' T-Stück'
                el.wert_schreiben()
            t.Commit()
            t.Dispose()
            self.GUI.pb_t.Visibility = Hidden
            self.GUI.pb_c.Visibility = Hidden
            self.GUI.button_dimension.IsEnabled = True
            self.GUI.button_change.IsEnabled = True
        except Exception as e:
            self.GUI.pb_t.Visibility = Hidden
            self.GUI.pb_c.Visibility = Hidden
            self.GUI.button_dimension.IsEnabled = True
            self.GUI.button_change.IsEnabled = True
            logger.error(e)

    
    def GetName(self):
        return "dimensionieren"
    
    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text

class UEBERNEHMEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        self.GUI.write_config()
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        logger = script.get_logger()
        self.GUI.button_dimension.IsEnabled = False
        self.GUI.button_change.IsEnabled = False
        try:
            self.GUI.pb_t.Visibility = Visible
            self.GUI.pb_c.Visibility = Visible

            liste_formteil = []
            for elid in self.GUI.allebauteile.formteil:
                el = doc.GetElement(elid)
                formteil = Rohrformteil(el)
                if formteil.art:
                    if formteil.art == 'Übergang':
                        if formteil.conns0 and formteil.conns1:
                            liste_formteil.append(formteil)
                    else:
                        liste_formteil.append(formteil)

            t = DB.Transaction(doc,'Dimension übernehmen')
            t.Start()
            doc.Delete(List[DB.ElementId](self.GUI.allebauteile.formteil))
            doc.Regenerate()
            for el in self.GUI.allebauteile.rohr:
                try:
                    Rohr(el,doc).wert_schreiben()
                except Exception as e:
                    logger.error(e)
            doc.Regenerate()
            self.GUI.maxvalue = len(liste_formteil)

            for n,el in enumerate(liste_formteil):
                self.GUI.value = n+1
                self.GUI.PB_text = str(n+1) + ' / '+ str(len(liste_formteil)) + ' Formteile'
                self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                try:
                    el.doc = doc
                    el.createfittings()
                except:
                    try:
                        doc = app.ActiveUIDocument.Document
                        doc.Regenerate()
                        el.doc = doc
                        el.createfittings()
                    except Exception as e:
                        logger.error(e)

            t.Commit()
            t.Dispose()
            self.GUI.pb_t.Visibility = Hidden
            self.GUI.pb_c.Visibility = Hidden
        except Exception as e:
            self.GUI.pb_t.Visibility = Hidden
            self.GUI.pb_c.Visibility = Hidden
            logger.error(e)
    
    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text


    def GetName(self):
        return "Verbinden"