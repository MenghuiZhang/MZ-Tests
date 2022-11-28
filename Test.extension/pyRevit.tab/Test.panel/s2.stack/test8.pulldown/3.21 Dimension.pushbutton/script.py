# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
from rpw import revit,DB,UI
import os
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media import Brushes,VisualTreeHelper
from System.Windows.Input import Key
from System.Windows.Controls import GridViewRowPresenter,Border
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventHandler,PropertyChangedEventArgs

__title__ = "Sprinkler-Strang dimensionieren"
__doc__ = """

Parameter: IGF_X_SM_Durchmesser

[2022.08.16]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

logger = script.get_logger()
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(number+name+__title__)

Schwarz = Brushes.Black
Rot = Brushes.Red
uidoc = revit.uidoc

class FilterNeben(UI.Selection.ISelectionFilter):
    def __init__(self,liste):
        self.Liste = liste

    def AllowElement(self,element):
        if element.Id.ToString() in self.Liste:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter(UI.Selection.ISelectionFilter):

    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False


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

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.Liste = ObservableCollection[TemplateItem]()
        self.read_config()
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.lv.ItemsSource = self.Liste
        
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    def read_config(self):
        try:
            if config.Einstellungen:
                for el in config.Einstellungen:
                    self.Liste.Add(TemplateItem(el[2],el[0],el[1]))
        except:
            pass
        if self.Liste.Count == 0:
            self.Liste.Add(TemplateItem(None,1,None))
    
    def write_config(self):
        Liste = []
        try:
            for el in self.Liste:
                Liste.append([el.von,el.bis,el.dimension])
        except:
            pass
        config.Einstellungen = Liste
        script.save_config()
    
    def Setkey(self, sender, args):   
        if ((args.Key >= Key.D0 and args.Key <= Key.D9) or (args.Key >= Key.NumPad0 and args.Key <= Key.NumPad9) \
            or args.Key == Key.Delete or args.Key == Key.Back):
            args.Handled = False
        else:
            args.Handled = True
            
    def von_changed(self, sender, args):
        value = sender.Text
        item = sender.DataContext
        index = self.Liste.IndexOf(item)
        item.von = value
        if value != '' and value != None:
            if index > 0:
                for n in range(1,self.Liste.Count):
                    alt = self.Liste.Item[n-1]
                    try:
                        alt.bis = str(int(self.Liste.Item[n].von)-1)
                        if int(alt.bis) < int(alt.von):
                            alt.farbe_bis = Rot
                            self.Liste.Item[n].farbe_von = Rot
                        else:
                            alt.farbe_bis = Schwarz
                            self.Liste.Item[n].farbe_von = Schwarz
                    except Exception as e:print(e)   

    def Add(self, sender, args):
        self.Liste.Add(TemplateItem('20',None,None))
        self.lv.Items.Refresh()

    def dele(self, sender, args):
        self.Liste.Remove(self.lv.SelectedItem)
        for n in range(1,self.Liste.Count):
            alt = self.Liste.Item[n-1]
            try:
                alt.bis = str(int(self.Liste.Item[n].von)-1)
                if alt.bis < alt.von:
                    alt.farbe_bis = Rot
                    self.Liste.Item[n].farbe_von = Rot
                else:
                    alt.farbe_bis = Schwarz
                    self.Liste.Item[n].farbe_von = Schwarz
            except:pass   
        self.lv.Items.Refresh()


    def OK(self, sender, args):
        self.write_config()
        self.Close()
    def cancel(self, sender, args):
        self.Close()
        script.exit()
    
    def Selection_Changed(self,sender,args):
        if self.lv.SelectedIndex != -1:self.Remove.IsEnabled = True
        else:self.Remove.IsEnabled = False

       
einstellung = AktuelleBerechnung()
try:
    einstellung.ShowDialog()
except Exception as e:
    logger.error(e)
    einstellung.Close()
    script.exit()

dict_dimension = {int(el.von):el.dimension for el in einstellung.Liste}
List_dimension = dict_dimension.keys()[:]
List_dimension.sort(reverse=True)


t_dict = {}

class TPiece:
    def __init__(self,Liste_Rohre,elemid):
        self.dict_dimension_Neu = dict_dimension
        self.elemid = DB.ElementId(int(elemid))
        self.angepasst = False
        self.elem = doc.GetElement(self.elemid)
        self.Liste_Pre_Rohre = Liste_Rohre
        self.liste = []
        self.Liste_Rohre_1 = []
        self.Liste_Rohre_2 = []
        self.Anzahl_Endverbraucher = 0
        self.T_Piece_0 = None
        self.T_Piece_1 = None
        self.Endverbraucher_0 = None
        self.Endverbraucher_1 = None
        self.Dimension = 0
        self.grunddimension = self.dict_dimension_Neu[List_dimension[-1]]
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
        if self.T_Piece_0 in t_dict.keys():
            if t_dict[self.T_Piece_0].angepasst == False:
                logger.error(t_dict[self.T_Piece_0].elemid.ToString())
            self.Anzahl_Endverbraucher += t_dict[self.T_Piece_0].Anzahl_Endverbraucher
        if self.T_Piece_1 in t_dict.keys():
            if t_dict[self.T_Piece_1].angepasst == False:
                logger.error(t_dict[self.T_Piece_1].elemid.ToString())
            self.Anzahl_Endverbraucher += t_dict[self.T_Piece_1].Anzahl_Endverbraucher    
    def get_Dimension(self):
        for n in List_dimension:
            if self.Anzahl_Endverbraucher >= n:
                self.Dimension = self.dict_dimension_Neu[n]
                break

    def wert_schreiben(self):
        if self.Endverbraucher_0 != None:
            for rohr in self.Liste_Rohre_1:
                try:
                    doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.grunddimension))
                except Exception as e:
                    logger.error(e)
        if self.Endverbraucher_1 != None:
            for rohr in self.Liste_Rohre_2:
                try:
                    doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.grunddimension))
                except Exception as e:
                    logger.error(e)
        
        for rohr in self.Liste_Pre_Rohre:
            try:
                doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.Dimension))
            except Exception as e:
                logger.error(e)

ref_0 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,Filter(),'Wählt den Rohr aus')
el0 = doc.GetElement(ref_0)
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

class anfang:
    def __init__(self,rohr_liste,elem):
        self.elem = elem
        self.rohrliste = rohr_liste
        self.liste = []
        self.t_stueck = None
        for el in self.rohrliste:
            self.liste.append(el)
        self.get_t_st(self.elem)
        
    def get_t_st(self,elem):
        if self.t_stueck:return
        
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
                if conns.Size > 2 and cate == 'Rohrformteile':
                    self.t_stueck = elem.Id
                    return
                    
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste:  
                            if owner.Category.Id.ToString() == '-2008044':
                                if self.t_stueck == None:
                                    self.rohrliste.append(owner.Id.ToString())                         
                            self.get_t_st(owner)


_Liste_T_Stuecke = []
if len(liste) > 0:
    ref_1 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,FilterNeben(liste),'Wählt den nächsten Teil aus')
    a = anfang(Liste_Rohr,doc.GetElement(ref_1))
    b = TPiece(a.rohrliste,a.t_stueck.ToString())
    t_dict[a.t_stueck.ToString()] = b
    Liste_T_Stuecke = []
    _Liste_T_Stuecke.append(a.t_stueck.ToString())
    if b.T_Piece_0 != None:Liste_T_Stuecke.append([b.T_Piece_0,b.Liste_Rohre_1])
    if b.T_Piece_1 != None:Liste_T_Stuecke.append([b.T_Piece_1,b.Liste_Rohre_2])
    while(len(Liste_T_Stuecke) > 0):
        t_temp = []
        for el in Liste_T_Stuecke:
            b = TPiece(el[1],el[0])
            if b.T_Piece_0 != None:t_temp.append([b.T_Piece_0,b.Liste_Rohre_1])
            if b.T_Piece_1 != None:t_temp.append([b.T_Piece_1,b.Liste_Rohre_2])
            t_dict[el[0]] = b
            _Liste_T_Stuecke.insert(0,el[0])
        Liste_T_Stuecke = t_temp

    for el in _Liste_T_Stuecke:
        elem = t_dict[el]
        elem.get_Liste_Verbraucher()
        elem.get_Dimension()
        elem.angepasst = True

    t = DB.Transaction(doc,'Dimension')
    t.Start()
    for el in t_dict.values():
        el.wert_schreiben()
    t.Commit()
else:
    pass
