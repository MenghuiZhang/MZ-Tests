# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from IGF_log import getlog,getloglocal
from rpw import revit,DB,UI
from pyrevit import script, forms
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from IGF_lib import ViewFilter_Erstellen,ViewFilterInViewHinzufügen

__title__ = "Ansichtsfilter erstellen"
__doc__ = """

Systemlegenden für ausgewählte Ansicht erstellen

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

class Itemtemplate(TemplateItemBase):
    def __init__(self,elem,name,checked = False):
        TemplateItemBase.__init__(self)
        self.elem = elem
        self.elemid = elem.Id
        
        self.name = name
        self._checked = checked
        

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

class Ansicht(Itemtemplate):
    def __init__(self,elem,name):
        Itemtemplate.__init__(self,elem,name)
        self.istTempalte = self.elem.IsTemplate
        self.viewtemplate = self.get_grundriss()
    
    def get_grundriss(self):
        if self.istTempalte:
            return
        else:
            if self.elem.ViewTemplateId.IntegerValue != -1:
                return self.elem.ViewTemplateId.IntegerValue
            else:return self.elemid.IntegerValue

class SystemClass(Itemtemplate):
    def __init__(self,elem,name):
        Itemtemplate.__init__(self,elem,name)

# Viewssource
VS_Views = ObservableCollection[Ansicht]()
VS_Grundrisse = ObservableCollection[Ansicht]()
VS_3d = ObservableCollection[Ansicht]()
VS_Schnitt = ObservableCollection[Ansicht]()
VS_System = ObservableCollection[SystemClass]()
VS_DuctSystem = ObservableCollection[SystemClass]()
VS_PipeSystem = ObservableCollection[SystemClass]()
activeview = uidoc.ActiveView

def GetAllViews():
    _dict = {}
    views = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).ToElements()
    for v in views:
        if v.IsTemplate:continue
        name = v.Name
        fam = v.ViewType.ToString()
        if fam in ['FloorPlan','ThreeD','Section']:
            _dict[name] = Ansicht(v,name)
            if name == activeview.Name:_dict[name].checked = True
    for name in sorted(_dict.keys()):
        ansicht = _dict[name]
        VS_Views.Add(ansicht)
        fam = ansicht.elem.ViewType.ToString()
        if fam == 'FloorPlan':
            VS_Grundrisse.Add(ansicht)
        elif fam == 'ThreeD':
            VS_3d.Add(ansicht)
        else:
            VS_Schnitt.Add(ansicht)
GetAllViews()

def GetAllSystemtype():
    _dict = {}
    systems = DB.FilteredElementCollector(doc).OfClass(DB.MEPSystemType).ToElements()
    for system in systems:
        name = system.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        _dict[name] = SystemClass(system,name)

    for name in sorted(_dict.keys()):
        if _dict[name].elem.GetType().Name == 'MechanicalSystemType':
            VS_DuctSystem.Add(_dict[name])
            VS_System.Add(_dict[name])
        elif _dict[name].elem.GetType().Name == 'PipingSystemType':
            VS_PipeSystem.Add(_dict[name])
            VS_System.Add(_dict[name])

GetAllSystemtype()

# Texttyp

class WPFUI(forms.WPFWindow):
    def __init__(self):
        self.VS_Views = VS_Views
        self.VS_Grundrisse = VS_Grundrisse
        self.VS_3d = VS_3d
        self.VS_Schnitt = VS_Schnitt
        self.VS_System = VS_System
        self.VS_DuctSystem = VS_DuctSystem
        self.VS_PipeSystem = VS_PipeSystem
        self.template_view = ObservableCollection[Ansicht]() 
        self.template_system = ObservableCollection[SystemClass]()   
        self.leer_view = ObservableCollection[Ansicht]() 
        self.temp_view1 = ObservableCollection[Ansicht]() 
        self.leer_system = ObservableCollection[SystemClass]()    
        self.alt_system = self.VS_System
        self.alt_View = self.VS_Views    
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.DG_System.ItemsSource = self.VS_System
        self.DG_Ansichts.ItemsSource = self.VS_Views

    def suchechanged(self, sender, args):
        self.template_view.Clear()
        text_typ = self.suche.Text    

        if text_typ in ['',None]:
            self.DG_Ansichts.ItemsSource = self.alt_View
            return

        for item in self.alt_View:
            if item.name.upper().find(text_typ.upper()) != -1:
                self.template_view.Add(item)

            self.DG_Ansichts.ItemsSource = self.template_view
        self.DG_Ansichts.Items.Refresh()
    
    def Systemtyp_suchechanged(self, sender, args):
        self.template_system.Clear()
        text_typ = self.suche_Systemtyp.Text    

        if text_typ in ['',None]:
            self.DG_System.ItemsSource = self.alt_system
            return

        for item in self.alt_system:
            if item.name.upper().find(text_typ.upper()) != -1:
                self.template_system.Add(item)

            self.DG_System.ItemsSource = self.template_system
        self.DG_System.Items.Refresh()

    def systemchecked(self, sender, args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.DG_System.SelectedIndex != -1:
            if item in self.DG_System.SelectedItems:
                for el in self.DG_System.SelectedItems:
                    el.checked = checked
    
    def ansichtschecked(self, sender, args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.DG_Ansichts.SelectedIndex != -1:
            if item in self.DG_Ansichts.SelectedItems:
                for el in self.DG_Ansichts.SelectedItems:
                    el.checked = checked

    def systemcheckedchanged(self,sender,e):
   
        if self.pipe_system.IsChecked and self.duct_system.IsChecked:
            self.alt_system = self.VS_System
        elif (not self.pipe_system.IsChecked) and self.duct_system.IsChecked:
            self.alt_system = self.VS_DuctSystem
            
        elif self.pipe_system.IsChecked and (not self.duct_system.IsChecked):
            self.alt_system = self.VS_PipeSystem
            
        elif (not self.pipe_system.IsChecked) and (not self.duct_system.IsChecked):
            self.alt_system = self.leer_system
        
        self.DG_System.ItemsSource = self.alt_system
    
    def viewcheckedchanged(self,sender,e):
        self.temp_view1.Clear()
        for el in self.VS_Views:
            if self._grundriss.IsChecked:
                if self.VS_Grundrisse.Contains(el):
                    self.temp_view1.Add(el)
            if self._3D.IsChecked:
                if self.VS_3d.Contains(el):
                    self.temp_view1.Add(el)
            if self._schnitte.IsChecked:
                if self.VS_Schnitt.Contains(el):
                    self.temp_view1.Add(el)
        
        self.alt_View = self.temp_view1
        self.DG_Ansichts.ItemsSource = self.alt_View

    def _getsystemtyp(self):
        systemtyp = [el for el in self.VS_System if el.checked]
        return systemtyp

    def _filtererstellen(self,systemtyp,t):
        if len(systemtyp) > 0:
            with forms.ProgressBar(title = "{value}/{max_value} Filter werden erstellt/angepasst.",cancellable=True,step=1) as pb:
                for n, system in enumerate(systemtyp):
                    if pb.cancelled:
                        if t.HasStarted: t.RollBack()
                        return
                    pb.update_progress(n + 1, len(systemtyp))
                    ViewFilter_Erstellen(system.elem)
    
    def _filterhinzufugen(self,ansichten,systemtyp,filters,t):
        if len(ansichten) > 0:
            with forms.ProgressBar(title = "{value}/{max_value} Ansichten werden angepasst.",cancellable=True,step=1) as pb:
                for n, ansichtid in enumerate(ansichten):
                    if pb.cancelled:
                        if t.HasStarted: t.RollBack()
                        return
                    pb.update_progress(n + 1, len(ansichten))
                    ansicht = doc.GetElement(DB.ElementId(ansichtid))
                    for system in systemtyp:
                        _filter = filters[system.name]
                        ViewFilterInViewHinzufügen(system.elem,_filter,ansicht)

    def filtererstellen(self, sender, args):
        self.Close()
        systemtyp = self._getsystemtyp()
        if len(systemtyp) == 0:
            return
        t = DB.Transaction(doc,'Filter erstellen')
        t.Start()
        self._filtererstellen(systemtyp,t)
        t.Commit()
        t.Dispose()
    
    def _get_all_filters(self):
        Filters_coll = DB.FilteredElementCollector(doc).OfClass(DB.FilterElement).ToElements()
        FILTERS = {i.Name :i for i in Filters_coll}
        return FILTERS

    def _get_all_ansichts(self):
        ansichts = []
        for el in self.VS_Views:
            if el.checked:
                if el.istTempalte:
                    if el.elemid.IntegerValue not in ansichts:
                        ansichts.append(el.elemid.IntegerValue)
                else:
                    if el.viewtemplate not in ansichts:
                        ansichts.append(el.viewtemplate)
        return ansichts

    def _get_all_ansichts1(self):
        ansichts1 = []
        ansichts2 = []
        ansichts3 = []
        for el in self.VS_Views:
            if el.checked:
                if el.istTempalte:
                    if el.elemid.IntegerValue not in ansichts1:
                        ansichts1.append(el.elemid.IntegerValue)
                else:
                    if el.elemid.IntegerValue != el.viewtemplate:
                        if el.viewtemplate not in ansichts1:
                            ansichts1.append(el.viewtemplate)
                            ansichts2.append(el.elemid.IntegerValue)
                    else:
                        ansichts3.append(el.viewtemplate)
        

        
        for elid in ansichts2:
            ansicht = doc.GetElement(DB.ElementId(elid))
            view = ansicht.CreateViewTemplate()
            tempname = doc.GetElement(ansicht.ViewTemplateId).Name
            name = tempname + ' Kopie 1'
            n = 1
            while (name in names):
                n += 1
                name = tempname + ' Kopie ' + str(n)
            view.Name = name
            doc.Regenerate()
            ansichts3.append(view.Id.IntegerValue)

        return ansichts3
    
    def filtererstelleninanischt(self, sender, args):
        self.Close()
        systemtyp = self._getsystemtyp()
        if len(systemtyp) == 0:
            return
        t = DB.Transaction(doc,'Filter erstellen')
        t.Start()
        self._filtererstellen(systemtyp,t)
        
        doc.Regenerate()
        ansichts = self._get_all_ansichts()
        filters = self._get_all_filters()
        self._filterhinzufugen(ansichts,systemtyp,filters,t)
        t.Commit()
        t.Dispose()
    
    def filtererstelleninanischt1(self, sender, args):
        self.Close()
        systemtyp = self._getsystemtyp()
        if len(systemtyp) == 0:
            return
        t = DB.Transaction(doc,'Filter erstellen')
        t.Start()
        self._filtererstellen(systemtyp,t)
        
        doc.Regenerate()
        ansichts = self._get_all_ansichts1()
        filters = self._get_all_filters()
        self._filterhinzufugen(ansichts,systemtyp,filters,t)
        t.Commit()
        t.Dispose()
 
        
       
wpfui = WPFUI()
try:wpfui.ShowDialog()
except Exception as e:
    wpfui.Close()
    logger.error(e)