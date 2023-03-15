import clr
clr.AddReference('System.ComponentModel')
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
import Autodesk.Revit.DB as DB

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

class ItemTemplateMitElem(ItemTemplate):
    def __init__(self,elem):
        ItemTemplate.__init__(self)
        self.elem = elem
        self.name = self.elem.Name
        self.doc = self.elem.Document
        self.famname = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsString()
        self.elemid = self.elem.Id

