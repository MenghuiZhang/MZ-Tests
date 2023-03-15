# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from IGF_log import getlog,getloglocal
from rpw import revit,DB,UI
from pyrevit import script, forms
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs


__title__ = "Ansichtsfilter anpassen"
__doc__ = """

Filter für aktuelle Ansicht anpassen

[2023.02.02]
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

uidoc = revit.uidoc
doc = revit.doc
logger = script.get_logger()
names = [v.Name for v in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).ToElements()]

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

class FilterKlasse(TemplateItemBase):
    def __init__(self,elemid,name,sichtbar = False,checked = True):
        TemplateItemBase.__init__(self)
        self.elemid = elemid
        self.name = name
        self._sichtbar = sichtbar
        self._checked = checked
        

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    @property
    def sichtbar(self):
        return self._sichtbar
    @sichtbar.setter
    def sichtbar(self,value):
        if value != self._sichtbar:
            self._sichtbar = value
            self.RaisePropertyChanged('sichtbar')


# Viewssource
activeview = uidoc.ActiveView

itemssource = ObservableCollection[FilterKlasse]()

def GetAllFilters():
    _dict = {}
    if activeview.ViewTemplateId.IntegerValue != -1:
        view = doc.GetElement(activeview.ViewTemplateId)
    else:
        view = activeview
    filters = view.GetFilters()
    for el in filters:
        name = doc.GetElement(el).Name
        sichtbarkeit = view.GetFilterVisibility(el)
        _dict[name] = FilterKlasse(el,name,sichtbarkeit)
    for name in sorted(_dict.keys()):
        itemssource.Add(_dict[name])

GetAllFilters()

# Texttyp

class WPFUI(forms.WPFWindow):
    def __init__(self):
        self.itemssource = itemssource
        self.leer_system = ObservableCollection[FilterKlasse]()    
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.DG_Filter.ItemsSource = self.itemssource

    def suchechanged(self, sender, args):
        self.leer_system.Clear()
        text_typ = sender.Text    

        if text_typ in ['',None]:
            self.DG_Filter.ItemsSource = self.itemssource
            return

        for item in self.itemssource:
            if item.name.upper().find(text_typ.upper()) != -1:
                self.leer_system.Add(item)

            self.DG_Filter.ItemsSource = self.leer_system
        self.DG_Filter.Items.Refresh()
    
    def filterchecked(self, sender, args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.DG_Filter.SelectedIndex != -1:
            if item in self.DG_Filter.SelectedItems:
                for el in self.DG_Filter.SelectedItems:
                    el.checked = checked
    
    def sichtbarkeitchecked(self, sender, args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.DG_Filter.SelectedIndex != -1:
            if item in self.DG_Filter.SelectedItems:
                for el in self.DG_Filter.SelectedItems:
                    el.sichtbar = checked

    def filteranpasseninanischt(self, sender, args):
        self.Close()
        t = DB.Transaction(doc,'Filter anpassen')
        t.Start()
        if activeview.ViewTemplateId.IntegerValue != -1:
            view = doc.GetElement(activeview.ViewTemplateId)
        else:
            view = activeview
        for el in self.itemssource:
            if el.checked:
                view.SetFilterVisibility(el.elemid,el.sichtbar)
            else:
                view.RemoveFilter(el.elemid)
        t.Commit()
    
    def filteranpasseninanischt1(self, sender, args):
        self.Close()
        t = DB.Transaction(doc,'Filter anpassen')
        t.Start()
        if activeview.ViewTemplateId.IntegerValue != -1:
            view = activeview.CreateViewTemplate()
            tempname = doc.GetElement(activeview.ViewTemplateId).Name
            name = tempname + ' Kopie 1'
            n = 1
            while (name in names):
                n += 1
                name = tempname + ' Kopie ' + str(n)
            view.Name = name
        else:
            view = activeview
        for el in self.itemssource:
            if el.checked:
                view.SetFilterVisibility(el.elemid,el.sichtbar)
            else:
                view.RemoveFilter(el.elemid)
        t.Commit()
 
        
       
wpfui = WPFUI()

try:wpfui.ShowDialog()
except Exception as e:
    wpfui.Close()
    logger.error(e)