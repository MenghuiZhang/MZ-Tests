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

class Systemtyp(TemplateItemBase):
    def __init__(self,systemtyp):
        TemplateItemBase.__init__(self)
        self._checked = True
        self.systemtyp = systemtyp

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

class Category(TemplateItemBase):
    def __init__(self,category,elems_ohne,elem_mit_system):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.category = category
        self.elems_ohne = elems_ohne
        self.elem_mit_system = elem_mit_system
        self.Items = ObservableCollection[Bauteil]()
        self.Systems = ObservableCollection[Systemtyp]()
        self.dict_Systems = {}
        self.get_Items()
    
    def get_Items(self):
        for system in sorted(self.elem_mit_system.keys()):
            item = Systemtyp(system)
            self.Systems.Add(item)
            self.dict_Systems[system] = item

        
        for family in sorted(self.elems_ohne.keys()):
            for typ in sorted(self.elems_ohne[family].keys()):
                item = Bauteil(self.category,family,typ)
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
    def __init__(self,category,familyname,typname,elems=[]):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.category = category
        self.familyname = familyname
        self.typname = typname
        self.elems = elems
        self._anzahl = self.get_anzahl()
    
    def get_anzahl(self):
        return len(self.elems)

    @property
    def anzahl(self):
        return self._anzahl
    @anzahl.setter
    def anzahl(self,value):
        if value != self._anzahl:
            self._anzahl = value
            self.RaisePropertyChanged('anzahl')

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
        Dict1 = {}
        ItemsSource_Cate = ObservableCollection[Category]()
        for elid in elemids:
            elem = doc.GetElement(elid)
            typname = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            familyname = elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
            try:System = elem.LookupParameter('Systemtyp').AsValueString()
            except:System = ''
            catname = elem.Category.Name

            if catname not in Dict.keys():
                Dict[catname] = {}
                Dict1[catname] = {}

            if System not in Dict1[catname].keys():
                Dict1[catname][System] = {}
            if familyname not in Dict1[catname][System].keys():
                Dict1[catname][System][familyname] = {}
            if typname not in Dict1[catname][System][familyname].keys():
                Dict1[catname][System][familyname][typname] = []
            Dict1[catname][System][familyname][typname].append(elid)

            if familyname not in Dict[catname].keys():
                Dict[catname][familyname] = {}
            if typname not in Dict[catname][familyname].keys():
                Dict[catname][familyname][typname] = []
            Dict[catname][familyname][typname].append(elid)


        for cate in sorted(Dict.keys()):
            Categoryinstance = Category(cate,Dict[cate],Dict1[cate])
            ItemsSource_Cate.Add(Categoryinstance)
        self.GUI.LV_category.ItemsSource = ItemsSource_Cate
        self.GUI.temp_coll.Clear()
        self.GUI.temp_coll_Systemtyp.Clear()
    
    def PostSelect(self,uiapp):
        self.name = 'ausgewähle Typen in Revit markieren'
        uidoc = uiapp.ActiveUIDocument
        Liste = []
        for el in self.GUI.DG_Familie.Items:
            if el.checked:
                Liste.extend(el.elems)
        uidoc.Selection.SetElementIds(List[DB.ElementId](Liste))

