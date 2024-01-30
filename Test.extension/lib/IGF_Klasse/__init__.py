import clr
clr.AddReference('System.ComponentModel')
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List

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

class ItemTemplate(TemplateItemBase):
    def __init__(self):
        TemplateItemBase.__init__(self)
        self._checked = False
        self._visibility = True
    
    @property
    def visibility(self):
        return self._visibility
    @visibility.setter
    def visibility(self,value):
        if value != self._visibility:
            self._visibility = value
            self.RaisePropertyChanged('visibility')

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

class ItemTemplateMitName(ItemTemplate):
    def __init__(self,name):
        ItemTemplate.__init__(self)
        self.name = name

class ItemTemplateNurElem(ItemTemplate):
    def __init__(self,elem):
        ItemTemplate.__init__(self)
        self.elem = elem

class ItemTemplateMitElemUndName(ItemTemplate):
    def __init__(self,name,elem):
        ItemTemplate.__init__(self)
        self.name = name
        self.elem = elem

class ItemTemplateMitElem(ItemTemplate):
    def __init__(self,elem):
        ItemTemplate.__init__(self)
        self.elem = elem
        self.name = self.elem.Name

        self.famname = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsString()
        self.elemid = self.elem.Id

    @property
    def doc(self):
        return self.elem.Document

class TEMPCLASS(TemplateItemBase):
    def __init__(self,Zeile,Nummer,Relevant,Class):
        TemplateItemBase.__init__(self)
        self.Zeile = Zeile
        self.Nummer = Nummer
        self._Relevant = Relevant
        self.Class = Class
    
    @property
    def Relevant(self):
        return self._Relevant
    
    @Relevant.setter
    def Relevant(self,value):
        if self._Relevant != value:
            self._Relevant = value
            self.RaisePropertyChanged('Relevant')

