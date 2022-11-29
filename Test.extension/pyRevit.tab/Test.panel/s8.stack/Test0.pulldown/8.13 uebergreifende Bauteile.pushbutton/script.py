# coding: utf8
from Autodesk.Revit.UI import TaskDialog
from System.Collections.ObjectModel import ObservableCollection
from System.Text.RegularExpressions import Regex
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Collections.Generic import List
from clr import GetClrType
from rpw import revit,DB
from pyrevit import script,forms
from IGF_log import getlog
import os

__title__ = "8.13 Gewerkkürzel an Übergreifenden Bauteilen schreiben"
__doc__ = """

parameter:
IGF_X_Übergreifend
IGF_X_Gewerkkürzel_Exemplar

[2022.11.29]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:getlog(__title__)
except:pass

def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if not param:return ''
    if param.StorageType.ToString() == 'ElementId':
        return param.AsValueString()
    elif param.StorageType.ToString() == 'Integer':
        value = param.AsInteger()
    elif param.StorageType.ToString() == 'Double':
        value = param.AsDouble()
    elif param.StorageType.ToString() == 'String':
        value = param.AsString()
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value

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
    def __init__(self,category,Familie,elems,Symbol):
        TemplateItemBase.__init__(self)
        self.Familiename = Familie
        self._checked = False
        self.category = category

        self.elems = elems
        self.symbol = Symbol
        
        try:self.get_daten()
        except:pass
        if len(self.elems) == 0:
            self.info = 'Typ nicht verwendet'
        else:
            self.info = 'Typ bereits verwendet'
    
    def get_daten(self):
        typ = self.symbol.LookupParameter('IGF_X_Übergreifend').AsInteger()
        if typ:
            self.checked = True
        else:
            self.checked = False
    
    def werte_schreiben(self):
        
        try:
            if self.checked:
                self.symbol.LookupParameter('IGF_X_Übergreifend').Set(1)
            else:
                self.symbol.LookupParameter('IGF_X_Übergreifend').Set(0)
        except:pass
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

AUSWAHL_HEIZKOERPER_IS = ObservableCollection[Familie]()

def get_Heizkoeper_IS():
    Dict = {}
    filter_ = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_MechanicalEquipment,DB.BuiltInCategory.OST_GenericModel,DB.BuiltInCategory.OST_PlumbingFixtures,DB.BuiltInCategory.OST_MEPAnalyticalAirLoop]))
    HLSs = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter_).ToElements()
    for el in HLSs:
        category = el.Category.Name
        FamilyName = el.Symbol.FamilyName + ': ' + el.Name
        if category not in Dict.keys():
            Dict[category] = {}   
        if FamilyName not in Dict[category].keys():
            Dict[category][FamilyName] = [el]            
        else:Dict[category][FamilyName].append(el)
    
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Dict1 = {}
    for el in Families:
        if el.FamilyCategoryId.IntegerValue in [-2001140,-2000151,-2001160,-2001008]:
            category = el.FamilyCategory.Name
            if category not in Dict1.keys():
                Dict1[category] = {}
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Dict1[category].keys():
                    Dict1[category][typname] = typ
    
    for category in sorted(Dict1.keys()):
        for fam in sorted(Dict1[category].keys()):
            if category in Dict.keys():
                if fam in Dict[category].keys():
                    AUSWAHL_HEIZKOERPER_IS.Add(Familie(category,fam,Dict[category][fam],Dict1[category][fam]))
                else:
                    AUSWAHL_HEIZKOERPER_IS.Add(Familie(category,fam,[],Dict1[category][fam]))
            else:
                AUSWAHL_HEIZKOERPER_IS.Add(Familie(category,fam,[],Dict1[category][fam]))
    
get_Heizkoeper_IS()

class Familienauswahl(forms.WPFWindow):
    def __init__(self):
        self.HLS_IS = AUSWAHL_HEIZKOERPER_IS
        self.temp_coll = ObservableCollection[Familie]()
        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.lv.ItemsSource = self.HLS_IS
        self.result = False
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))
        

    def cancel(self,sender,args):
        self.Close()
            
    def schreiben(self,sender,args):
        self.Close()
        self.result = True
        t = DB.Transaction(doc,'Übergreifend')
        t.Start()
        for el in self.HLS_IS:
            el.werte_schreiben()
        t.Commit()
        t.Dispose()
    
    def checkedboxorsuchechanged(self):
        text = self.suche.Text
        checked = self.checkbox.IsChecked
        self.temp_coll.Clear()
        if not text:
            if checked:
                for el in self.HLS_IS:
                    if el.checked:
                        self.temp_coll.Add(el)
                self.lv.ItemsSource = self.temp_coll

            else:
                self.lv.ItemsSource = self.HLS_IS

            return 
        else:
            if checked:
                for el in self.HLS_IS:
                    if el.checked and (el.Familiename.upper().find(text.upper()) != -1 or el.category.upper().find(text.upper()) != -1):
                        self.temp_coll.Add(el)
                
            else:
                for el in self.HLS_IS:
                    if el.Familiename.upper().find(text.upper()) != -1 or el.category.upper().find(text.upper()) != -1:
                        self.temp_coll.Add(el)
            self.lv.ItemsSource = self.temp_coll

    def familiecheckedchanged(self,sender,e):
        self.checkedboxorsuchechanged()
        
    def textchanged(self,sender,e):
        self.checkedboxorsuchechanged()
    
    def checkedchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.checked = checked
        
FamilienDialog = Familienauswahl()
try:FamilienDialog.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienDialog.Close()
    script.exit()

class FamilieExamplar(TemplateItemBase):
    def __init__(self,Familie,elem):
        TemplateItemBase.__init__(self)
        self.Familiename = Familie.Familiename
        self._checked = False
        self._B = False
        self._G = False
        self._H = False
        self._K = False
        self._M = False
        self._R = False
        self._S = False
        self.category = Familie.category
        self.elem = elem
        
        try:self.get_daten()
        except:pass
    
    def get_daten(self):
        typ = self.elem.LookupParameter('IGF_X_Gewerkkürzel_Exemplar').AsString()
        if typ:
            liste = typ.split('_')
            for _typ in liste:
                if _typ == 'B':self.B = True
                elif _typ == 'G':self.G = True
                elif _typ == 'H':self.H = True
                elif _typ == 'K':self.K = True
                elif _typ == 'M':self.M = True
                elif _typ == 'R':self.R = True
                elif _typ == 'S':self.S = True
    
    def werte_schreiben(self):
        text = ''
        if self.B:text+='B_'
        if self.G:text+='G_'
        if self.H:text+='H_'
        if self.K:text+='K_'
        if self.M:text+='M_'
        if self.R:text+='R_'
        if self.S:text+='S_'
        if text is not None:
            if text:self.elem.LookupParameter('IGF_X_Gewerkkürzel_Exemplar').Set(text[:-1])
            else:self.elem.LookupParameter('IGF_X_Gewerkkürzel_Exemplar').Set('')
        
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    @property
    def S(self):
        return self._S
    @S.setter
    def S(self,value):
        if value != self._S:
            self._S = value
            self.RaisePropertyChanged('S')
    @property
    def B(self):
        return self._B
    @B.setter
    def B(self,value):
        if value != self._B:
            self._B = value
            self.RaisePropertyChanged('B')
    @property
    def G(self):
        return self._G
    @G.setter
    def G(self,value):
        if value != self._G:
            self._G = value
            self.RaisePropertyChanged('G')
    @property
    def H(self):
        return self._H
    @H.setter
    def H(self,value):
        if value != self._H:
            self._H = value
            self.RaisePropertyChanged('H')
    @property
    def K(self):
        return self._K
    @K.setter
    def K(self,value):
        if value != self._K:
            self._K = value
            self.RaisePropertyChanged('K')
    @property
    def M(self):
        return self._M
    @M.setter
    def M(self,value):
        if value != self._M:
            self._M = value
            self.RaisePropertyChanged('M')
    @property
    def R(self):
        return self._R
    @R.setter
    def R(self,value):
        if value != self._R:
            self._R = value
            self.RaisePropertyChanged('R')

class Uebergreifend(forms.WPFWindow):
    def __init__(self,DatenIS):
        self.HLS_IS = DatenIS
        self.temp_coll = ObservableCollection[FamilieExamplar]()
        forms.WPFWindow.__init__(self, 'mainwindow.xaml',handle_esc=False)
        self.lv.ItemsSource = self.HLS_IS
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))

    def cancel(self,sender,args):
        self.Close()
            
    def ok(self,sender,args):
        self.Close()
        t = DB.Transaction(doc,'Gewerkkürzel schreiben')
        t.Start()
        for el in self.HLS_IS:
            if el.checked:el.werte_schreiben()
        t.Commit()
        t.Dispose()
        
    def suchetextchanged(self,sender,e):
        text = sender.Text
        self.temp_coll.Clear()
        if not text:
            self.lv.ItemsSource = self.HLS_IS
        else:
            for el in self.HLS_IS:
                if el.Familiename.upper().find(text.upper()) != -1 or el.category.upper().find(text.upper()) != -1:
                    self.temp_coll.Add(el)
            self.lv.ItemsSource = self.temp_coll

    def checkedchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.checked = checked
    
    def bchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.B = checked
    def gchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.G = checked
    def hchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.H = checked
    def kchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.K = checked
    def mchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.M = checked
    def rchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.R = checked
    def schanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.S = checked

if FamilienDialog.result:
    Liste = [el for el in AUSWAHL_HEIZKOERPER_IS if el.checked]
    if len(Liste) > 0:
        IS_Daten = ObservableCollection[FamilieExamplar]()
        for el in Liste:
            for elem in el.elems:
                IS_Daten.Add(FamilieExamplar(el,elem))
        if IS_Daten.Count > 0:
            FamilienDialog1 = Uebergreifend(IS_Daten)
            try:FamilienDialog1.ShowDialog()
            except Exception as e:
                logger.error(e)
                FamilienDialog1.Close()
                script.exit()
