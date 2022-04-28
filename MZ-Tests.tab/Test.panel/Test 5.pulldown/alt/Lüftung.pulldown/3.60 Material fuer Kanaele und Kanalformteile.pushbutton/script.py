# coding: utf8
from System import Guid
from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow

__title__ = "3.60 Material für Kanäle und Luftkanalformteile"
__doc__ = """CAx Materialkz --->>> IGF_RLT_Material
CAx Materialkz: 7e758303-a7ae-470f-ab85-738065c2824e
IGF_RLT_Material: IGF Parameter """
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
active_view = uidoc.ActiveView

system_Luft = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
system_Luft_dict = {}

def coll2dict(coll,dict):
    for el in coll:
      #  name = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
        type = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        if type in dict.Keys:
            dict[type].append(el.Id)
        else:
            dict[type] = [el.Id]

coll2dict(system_Luft,system_Luft_dict)
system_Luft.Dispose()


class System(object):
    def __init__(self):
        self.checked = False
        self.SystemName = ''
        self.TypName = ''

    @property
    def TypName(self):
        return self._TypName
    @TypName.setter
    def TypName(self, value):
        self._TypName = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def ElementId(self):
        return self._ElementId
    @ElementId.setter
    def ElementId(self, value):
        self._ElementId = value

Liste_Luft = ObservableCollection[System]()

for key in system_Luft_dict.Keys:
    temp_system = System()
    temp_system.TypName = key
    temp_system.ElementId = system_Luft_dict[key]
    Liste_Luft.Add(temp_system)

# GUI Systemauswahl
class Systemauswahl(WPFWindow):
    def __init__(self, xaml_file_name,liste_Rohr):
        self.liste_Rohr = liste_Rohr
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[System]()
        self.altdatagrid = None

        try:
            self.dataGrid.ItemsSource = liste_Rohr
            self.altdatagrid = liste_Rohr
        except Exception as e:
            logger.error(e)

        self.suche.TextChanged += self.search_txt_changed

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.TypName.find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = True
        self.dataGrid.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = False
        self.dataGrid.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.checked
            item.checked = not value
        self.dataGrid.Items.Refresh()

    def auswahl(self,sender,args):
        self.Close()
    
Systemwindows = Systemauswahl("System.xaml",Liste_Luft)
Systemwindows.ShowDialog()

SystemListe = {}
for el in Liste_Luft:
    if el.checked == True:
        for it in el.ElementId:
            elem = doc.GetElement(it)
            sysname = elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
            systype = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            if not systype in SystemListe.Keys:
                SystemListe[systype] = [elem]
            else:
                SystemListe[systype].append(elem)

if forms.alert("Daten übernehmen?", ok=False, yes=True, no=True):
    t = Transaction(doc,'Material für Kanäle unf Formteile')
    t.Start()
    for systyp in SystemListe.Keys:
        sysliste = SystemListe[systyp]
        for sys_ele in sysliste:
            elements = sys_ele.DuctNetwork
            name = sys_ele.Name
            kanalundformteile = []
            for elem in elements:
                if elem.Category.Name in ['Luftkanäle','Luftkanalformteile']:
                    kanalundformteile.append(elem)

            title = '{value}/{max_value} Kanäle und Kanalformteile in System ' + systyp+': '+name
            with forms.ProgressBar(title=title,cancellable=True, step=10) as pb:
                for n_1,elem in enumerate(kanalundformteile):
                    if pb.cancelled:
                        t.RollBack()
                        script.exit()
                    pb.update_progress(n_1+1, len(kanalundformteile))
                    try:
                        typeid = elem.GetTypeId()
                        alt = doc.GetElement(typeid).get_Parameter(Guid('7e758303-a7ae-470f-ab85-738065c2824e')).AsString()
                        elem.LookupParameter('IGF_RLT_Material').Set(str(alt))
                    except Exception as e:
                        print(doc.GetElement(elem.GetTypeId()).FamilyName)
                        logger.error(e)
    t.Commit()