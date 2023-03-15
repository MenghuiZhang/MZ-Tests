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

__title__ = "Bauteil beschriften"
__doc__ = """

HLS-Bauteile/Luftkanal-/Rohrzubehör beschriftenParameter:
Typparameter: Beschreibung
Typparameter: Typenkommentare

[2022.11.22]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

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
        self._typ = ''
        self._beschreibung = ''
        self._size = ''
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
        self.beschreibung = self.symbol.LookupParameter('IGF_X_Beschreibung').AsString()
        self.typ = self.symbol.LookupParameter('IGF_X_Typ').AsString()
        self.size = self.symbol.LookupParameter('IGF_X_Größe').AsString()
        
    def werte_schreiben_beschrifen(self):
        try:self.symbol.LookupParameter('IGF_X_Beschreibung').Set(self.beschreibung)
        except:pass
        try:self.symbol.LookupParameter('IGF_X_Typ').Set(self.typ)
        except:pass
        try:self.symbol.LookupParameter('IGF_X_Größe').Set(self.size)
        except:pass
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

    @property
    def typ(self):
        return self._typ
    @typ.setter
    def typ(self,value):
        if value != self._typ:
            self._typ = value
            self.RaisePropertyChanged('typ')
    
    @property
    def size(self):
        return self._size
    @size.setter
    def size(self,value):
        if value != self._size:
            self._size = value
            self.RaisePropertyChanged('size')

    @property
    def beschreibung(self):
        return self._beschreibung
    @beschreibung.setter
    def beschreibung(self,value):
        if value != self._beschreibung:
            self._beschreibung = value
            self.RaisePropertyChanged('beschreibung')

AUSWAHL_HEIZKOERPER_IS = ObservableCollection[Familie]()

def get_Heizkoeper_IS():
    Dict = {}
    filter_ = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_MechanicalEquipment,
        DB.BuiltInCategory.OST_PipeAccessory,DB.BuiltInCategory.OST_DuctAccessory,
        DB.BuiltInCategory.OST_PlumbingFixtures,DB.BuiltInCategory.OST_Sprinklers,DB.BuiltInCategory.OST_DuctTerminal,]))
    HLSs = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter_).ToElements()
    for el in HLSs:
        category = el.Category.Name
        FamilyName = el.Symbol.FamilyName + ': ' + el.Name
        if category not in Dict.keys():
            Dict[category] = {}   
        if FamilyName not in Dict[category].keys():
            Dict[category][FamilyName] = [el.Id.ToString()]            
        else:Dict[category][FamilyName].append(el.Id.ToString())
    
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Dict1 = {}
    for el in Families:
        if el.FamilyCategoryId.IntegerValue in [-2001140,-2008016,-2008055,-2001160,-2008099,-2008013]:
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
        self.regex2 = Regex("[^,]+")
        self.temp_coll = ObservableCollection[Familie]()
        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.lv.ItemsSource = self.HLS_IS
    
    def textinput(self, sender, args):
        try:
            args.Handled = not self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True
        
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
    

    def cancel(self,sender,args):
        self.Close()
            
    def Beschriften(self,sender,args):
        t = DB.Transaction(doc,'Beschriften')
        t.Start()
        for el in self.HLS_IS:
            if el.checked:
                el.werte_schreiben_beschrifen()
        t.Commit()
        t.Dispose()
        TaskDialog.Show('Info','Erledigt!')


    def checkedchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv.SelectedIndex != -1:
            if item in self.lv.SelectedItems:
                for el in self.lv.SelectedItems:el.checked = checked
    

    def typchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv.SelectedItem is not None:
            if Item in self.lv.SelectedItems:
                for item in self.lv.SelectedItems:
                    item.typ = text 

    def sizechanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv.SelectedItem is not None:
            if Item in self.lv.SelectedItems:
                for item in self.lv.SelectedItems:
                    item.size = text  
    
    def beschreibungchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv.SelectedItem is not None:
            if Item in self.lv.SelectedItems:
                for item in self.lv.SelectedItems:
                    item.beschreibung = text  

FamilienDialog = Familienauswahl()
try:FamilienDialog.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienDialog.Close()
    script.exit()