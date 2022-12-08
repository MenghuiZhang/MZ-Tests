# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,TaskDialogCommonButtons 
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
from pyrevit import script


doc = __revit__.ActiveUIDocument.Document

if doc.IsFamilyDocument == False:script.exit()

params = {param.Definition.Name:param for param in doc.FamilyManager.Parameters if param.Definition.ParameterType.ToString() == 'YesNo'}



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

class Parameter(TemplateItemBase):
    def __init__(self,param):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.param = param
        self.paramname = self.param.Definition.Name
        self._wert = None

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    
    @property
    def wert(self):
        return self._wert
    @wert.setter
    def wert(self,value):
        if value != self._wert:
            self._wert = value
            self.RaisePropertyChanged('wert')

class Familietyp(TemplateItemBase):
    def __init__(self,typ,params):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.typ = typ
        self.Familietyp = typ.Name
        self.dict_params = params
        self.dict_values = {}
        self.get_params_values()
    
    def get_params_values(self):
        for paramname in self.dict_params.keys():
            value = self.typ.AsInteger(self.dict_params[paramname])
            self.dict_values[paramname] = value

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
       
IS_Params = ObservableCollection[Parameter]()
IS_Types = ObservableCollection[Familietyp]()
for el in sorted(params.keys()):IS_Params.Add(Parameter(params[el]))
for typ in doc.FamilyManager.Types:IS_Types.Add(Familietyp(typ,params))

class ListeExternalEvent(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.Name = ''
        self.ExcuseApp = ''
        
    def Execute(self,uiapp):
        if self.ExcuseApp:
            try:
                self.ExcuseApp(uiapp)
            except Exception as e:
                TaskDialog.Show('Fehler',e.ToString())
    
    def GetName(self):
        return self.Name

    def Sichtbarkeiteinstellen(self,uiapp):
        self.Name = 'Sichtbarkeit einstellen'
        if TaskDialog.Show('Sichtbarkeit','Werte für ausgewälten Typen schreiben?', TaskDialogCommonButtons.Yes|TaskDialogCommonButtons.No ).ToString() != 'Yes':
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document  
        familymanager = doc.FamilyManager
        currecttype = familymanager.CurrentType 
       
        t = DB.Transaction(doc,'Sichtbarkeit einstellen')
        t.Start()
        for typ in self.GUI.IS_Types:
            if typ.checked:
                familymanager.CurrentType  = typ.typ
                doc.Regenerate()
                for param in self.GUI.IS_Params:
                    if param.checked:
                        if param.wert is not None:
                            
                            if param.wert == True:
                                typ.dict_values[param.paramname] = 1
                                familymanager.Set(param.param,1)
                            else:
                                typ.dict_values[param.paramname] = 0
                                familymanager.Set(param.param,0)
        familymanager.CurrentType  = currecttype
        t.Commit()
        t.Dispose()
        TaskDialog.Show('Sichtbarkeit','Erledigt!')
        