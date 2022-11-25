# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit,forms,script
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
import math
from System.Windows.Controls import CheckBox, GridViewColumn, TextBlock
from System.Windows.Data import Binding
from System.Windows import DataTemplate,FrameworkElementFactory,TextWrapping


doc = revit.doc

class PARAMETER(object):
    def __init__(self,name,checked = False):
        self.checked = checked
        self.name = name

class ITEMTEMPLATE(object):
    def __init__(self,name,Liste,ListeParameter,Elems,DictElementType):
        self.checked = False
        self.typ = name
        self.Elems = Elems
        self.Liste = Liste
        self.defikey = name + '_______' + Liste
        self.ListeParameter = ListeParameter
        self.DictElementType = DictElementType
        self.ListElementType = sorted(self.DictElementType.keys())
        self.elementtypeindex = -1
   

        for n,e in enumerate(self.ListeParameter):
            if int(self.Liste[n]) == 0:
                setattr(self,e,False)
            else:
                setattr(self,e,True)


    def changetype(self):
        if self.elementtypeindex != -1:
            for el in self.Elems:
                try:
                    el.ChangeTypeId(self.DictElementType[self.ListElementType[self.elementtypeindex]])
                except Exception as e:
                    print(e)

class FAMILIEN(object):
    def __init__(self,name,elems):
        self.name = name
        self.checked = False
        self.elems = elems
        self.IS_Params = ObservableCollection[PARAMETER]()
        self.Alleparams = []
        self.Params = []
        self.Values = {}
        self.dict_Families = {}
        self.verteilung = ObservableCollection[ITEMTEMPLATE]()
    
    def get_all_params(self):
        for el in self.elems:
            params = el.Parameters
            for param in params:
                defi = param.Definition
                try:
                    typ = defi.ParameterType.ToString()
                    if typ == 'YesNo':
                        self.Alleparams.append(defi.Name)
                except:
                    try:
                        typ = defi.GetDataType().TypeId
                        if typ == 'autodesk.spec:spec.bool-1.0.0':
                            self.Alleparams.append(defi.Name)
                    except:
                        pass
            break
    
    def get_IS_Params(self):
        self.get_all_params()
        for el in sorted(self.Alleparams):
            self.IS_Params.Add(PARAMETER(el))
    
    def get_families(self):
        for el in self.elems:
            types = el.GetValidTypes()
            for typ in types:
                name = doc.GetElement(typ).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                self.dict_Families[name] = typ
            break

liste_category = List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_DuctAccessory,DB.BuiltInCategory.OST_PipeAccessory])

Dict_Familien = {}
IS_Familien = ObservableCollection[FAMILIEN]()

def get_IS_Familien():
    Bauteile = DB.FilteredElementCollector(revit.doc).WherePasses(DB.ElementMulticategoryFilter(liste_category)).WhereElementIsNotElementType()
    for el in Bauteile:
        name = el.Symbol.FamilyName
        if name not in Dict_Familien.keys():
            Dict_Familien[name] = []
        Dict_Familien[name].append(el)
    for el in sorted(Dict_Familien.keys()):
        familie = FAMILIEN(el,Dict_Familien[el])
        familie.get_IS_Params()
        familie.get_families()
        IS_Familien.Add(familie)

get_IS_Familien()

class EINGABE(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self,'Parameter.xaml',handle_esc=False)
        self.IS = []

    def Ok(self,sender,arg):
        self.Close()

    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.Params.SelectedItem is not None:
            try:
                if sender.DataContext in self.Params.SelectedItems:
                    for item in self.Params.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.Params.Items.Refresh()
                else:
                    pass
            except:
                pass

# class PROFILE(forms.WPFWindow):
#     def __init__(self):
#         forms.WPFWindow.__init__(self,'Profile.xaml',handle_esc=False)
#         self.profil = None
#         self.Profiles = []

#     def Ok(self,sender,arg):
#         self.profil = self.Profile.SelectedItem.ToString()
#         self.Close()

#     def add(self, sender, args):
#         self.Profiles.append('Profile1')
#         self.Profile.ItemsSource = self.Profiles
#         self.Profile.SelectedIndex = len(self.Profiles) - 1
#         self.profil = self.Profile.SelectedItem.ToString()
#     def delete(self, sender, args):
#         self.Profiles.remove(self.Profile.SelectedItem.ToString())
#         self.Profile.ItemsSource = self.Profiles
#         self.Profile.SelectedIndex = -1
#         self.choose.Isenabled = False
#         self.dele.IsEnabled = False
#     def profilechanged(self, sender, args):
#         if self.Profile.SelectedIndex == -1:
#             self.choose.IsEnabled = False
#             self.dele.IsEnabled = False
#         else:
#             self.choose.IsEnabled = True
#             self.dele.IsEnabled = True

            

class FAMILIEBEARBEITEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
    def Execute(self,uiapp):
        self.GUI.write_config()
        # script.save_config()
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        
        t = DB.Transaction(doc,'Berechnungstype zuweisen')
        t.Start()
        for el in self.GUI.IS_Familien:
            if el.checked:
                for ele in el.verteilung:
                    if ele.checked:
                        ele.changetype()

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Berechnungstype zuweisen"