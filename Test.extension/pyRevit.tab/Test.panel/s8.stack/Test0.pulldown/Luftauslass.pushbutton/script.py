# coding: utf8
from Autodesk.Revit.UI import TaskDialog
from System.Collections.ObjectModel import ObservableCollection
from System.Text.RegularExpressions import Regex
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from clr import GetClrType
from rpw import revit,DB
from pyrevit import script,forms
from IGF_log import getlog
from IGF_lib import get_value

__title__ = "Herstellerdaten an Luftauslässe schreiben "
__doc__ = """
schreibt Vmin, Vmax, Hesteller und Reduziertfaktor an Luftauslässes
parameter:
IGF_L_SollVolumenstromMin
IGF_L_SollVolumenstromMax
Hersteller
IGF_L_SollVolumenstromFaktor

[2022.11.23]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:getlog(__title__)
except:pass

logger = script.get_logger()
doc = revit.doc

uidoc = revit.uidoc

DICT_Families = {}

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

class Familie(TemplateItemBase):
    def __init__(self,Familie,elems,Symbol):
        TemplateItemBase.__init__(self)
        self.Familiename = Familie
        self._hersteller = ''
        self._vmin = ''
        self._vmax = ''
        self._fredu = ''

        self.elems = elems
        self.symbol = Symbol
        
        try:self.get_daten()
        except:pass
        if len(self.elems) == 0:
            self.info = 'Typ nicht verwendet'
        else:
            self.info = 'Typ bereits verwendet'
    
    def get_daten(self):
        self.hersteller = get_value(self.symbol.LookupParameter('Hersteller'))
        self.vmin = get_value(self.symbol.LookupParameter('IGF_L_SollVolumenstromMin'))
        self.vmax = get_value(self.symbol.LookupParameter('IGF_L_SollVolumenstromMax'))
        self.fredu = get_value(self.symbol.LookupParameter('IGF_L_SollVolumenstromFaktor'))
    
    def werte_schreiben(self):
        try:self.wertschreiben(self.symbol.LookupParameter('Hersteller'),self.hersteller)
        except:pass
        try:self.wertschreiben(self.symbol.LookupParameter('IGF_L_SollVolumenstromMin'),self.vmin)
        except:pass
        try:self.wertschreiben(self.symbol.LookupParameter('IGF_L_SollVolumenstromMax'),self.vmax)
        except:pass
        try:self.wertschreiben(self.symbol.LookupParameter('IGF_L_SollVolumenstromFaktor'),self.fredu)
        except:pass
    
    def wertschreiben(self,param,wert):
        if wert is not None:
            if not param.IsReadOnly:
                if param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                else:
                    param.Set(wert)
    
    @property
    def hersteller(self):
        return self._hersteller
    @hersteller.setter
    def hersteller(self,value):
        if value != self._hersteller:
            self._hersteller = value
            self.RaisePropertyChanged('hersteller')

    @property
    def vmin(self):
        return self._vmin
    @vmin.setter
    def vmin(self,value):
        if value != self._vmin:
            self._vmin = value
            self.RaisePropertyChanged('vmin')
    
    @property
    def vmax(self):
        return self._vmax
    @vmax.setter
    def vmax(self,value):
        if value != self._vmax:
            self._vmax = value
            self.RaisePropertyChanged('vmax')

    @property
    def fredu(self):
        return self._fredu
    @fredu.setter
    def fredu(self,value):
        if value != self._fredu:
            self._fredu = value
            self.RaisePropertyChanged('fredu')

AUSWAHL_HEIZKOERPER_IS = ObservableCollection[Familie]()

def get_Heizkoeper_IS():
    Dict = {}
    
    HLSs = DB.FilteredElementCollector(doc).OfCategoryId(DB.ElementId(-2008013)).WhereElementIsNotElementType().ToElements()
    for el in HLSs:
        FamilyName = el.Symbol.FamilyName + ': ' + el.Name
   
        if FamilyName not in Dict.keys():
            Dict[FamilyName] = [el.Id.ToString()]            
        else:Dict[FamilyName].append(el.Id.ToString())
    
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Dict1 = {}
    for el in Families:
        if el.FamilyCategoryId.IntegerValue == -2008013:
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Dict1.keys():
                    Dict1[typname] = typ
    
    for fam in sorted(Dict1.keys()):

        if fam in Dict.keys():
            AUSWAHL_HEIZKOERPER_IS.Add(Familie(fam,Dict[fam],Dict1[fam]))
        else:
            AUSWAHL_HEIZKOERPER_IS.Add(Familie(fam,[],Dict1[fam]))

    
get_Heizkoeper_IS()

class Familienauswahl(forms.WPFWindow):
    def __init__(self):
        self.HLS_IS = AUSWAHL_HEIZKOERPER_IS
        self.regex1 = Regex("[^0-9,]+")
        self.regex2 = Regex("[^0-9]+")
        self.temp_coll = ObservableCollection[Familie]()
        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.lv_auslass.ItemsSource = self.HLS_IS
    
    def textinput1(self, sender, args):
        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True
    
    def textinput(self, sender, args):
        text = sender.Text
        if text:
            if text.find(',') != -1 and args.Text == ',':
                args.Handled = True
                return
        try:
            args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True
            
        
    def textchanged(self,sender,e):
        text = sender.Text
        self.temp_coll.Clear()
        if not text:
            self.lv_auslass.ItemsSource = self.HLS_IS
            return
       
        else:
            for el in self.HLS_IS:
                if el.Familiename.upper().find(text.upper()) != -1:
                    self.temp_coll.Add(el)
            self.lv_auslass.ItemsSource = self.temp_coll  

    def cancel(self,sender,args):
        self.Close()
            
    def OK(self,sender,args):
        t = DB.Transaction(doc,'Herstellerdaten')
        t.Start()
        for el in self.HLS_IS:
            el.werte_schreiben()
        t.Commit()
        t.Dispose()
        TaskDialog.Show('Info','Erledigt!')
    

    def herstellerchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_auslass.SelectedItem is not None:
            if Item in self.lv_auslass.SelectedItems:
                for item in self.lv_auslass.SelectedItems:
                    item.hersteller = text 

    def vminchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_auslass.SelectedItem is not None:
            if Item in self.lv_auslass.SelectedItems:
                for item in self.lv_auslass.SelectedItems:
                    item.vmin = text 
    
    def vmaxchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_auslass.SelectedItem is not None:
            if Item in self.lv_auslass.SelectedItems:
                for item in self.lv_auslass.SelectedItems:
                    item.vmax = text 

    def fvchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_auslass.SelectedItem is not None:
            if Item in self.lv_auslass.SelectedItems:
                for item in self.lv_auslass.SelectedItems:
                    item.fredu = text  
    

FamilienDialog = Familienauswahl()
try:FamilienDialog.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienDialog.Close()
    script.exit()