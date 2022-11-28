# coding: utf8
from rpw import DB
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,ExternalEvent
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs

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

class Category(TemplateItemBase):
    def __init__(self,category,elems):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.category = category
        self.elems = elems
        self.Items = ObservableCollection[Bauteil]()
        self.get_Items()
    
    def get_Items(self):
        for family in sorted(self.elems.keys()):
            for typ in sorted(self.elems[family].keys()):
                item = Bauteil(self.category,family,typ,self.elems[family][typ],len(self.elems[family][typ]))
                self.Items.Add(item)
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

class Bauteil(TemplateItemBase):
    def __init__(self,category,familyname,typname,elems,anzahl):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.category = category
        self.familyname = familyname
        self.typname = typname
        self.elems = elems
        self.anzahl = anzahl
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

class ExternalEventListe(IExternalEventHandler):
    def __init__(self):
        self.name = ''
        self.GUI = None
        self.Executeapp = None
    def Execute(self,uiapp):
        try:
            self.Executeapp(uiapp)
        except Exception as e:
            TaskDialog.Show('Fehler',e.ToString())

    def GetName(self):
        return self.name
    
    def Reset(self,uiapp):
        self.name = 'ausgewählte Elemente aus Revit aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elemids = uidoc.Selection.GetElementIds()
        Dict = {}
        ItemsSource_Cate = ObservableCollection[Category]()
        for elid in elemids:
            elem = doc.GetElement(elid)
            typname = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            familyname = elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
            catname = elem.Category.Name
            if catname not in Dict.keys():
                Dict[catname] = {}
            if familyname not in Dict[catname].keys():
                Dict[catname][familyname] = {}
            if typname not in Dict[catname][familyname].keys():
                Dict[catname][familyname][typname] = []
            Dict[catname][familyname][typname].append(elid)
        for cate in sorted(Dict.keys()):
            Categoryinstance = Category(cate,Dict[cate])
            ItemsSource_Cate.Add(Categoryinstance)
        self.GUI.LV_category.ItemsSource = ItemsSource_Cate
        self.GUI.temp_coll.Clear()
    
    def PostSelect(self,uiapp):
        self.name = 'ausgewähle Typen in Revit markieren'
        uidoc = uiapp.ActiveUIDocument
        Liste = []
        for el in self.GUI.DG_Familie.Items:
            if el.checked:
                Liste.extend(el.elems)
        uidoc.Selection.SetElementIds(List[DB.ElementId](Liste))
        
