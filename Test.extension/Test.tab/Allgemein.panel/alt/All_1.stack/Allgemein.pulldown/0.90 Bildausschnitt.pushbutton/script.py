# coding: utf8
from pyrevit import revit, UI, DB
from pyrevit import script, forms
from System.Windows.Controls import *
from Autodesk.Revit.DB import *
from pyrevit.forms import WPFWindow
from System import Guid
from System.Collections.ObjectModel import *

__title__ = "0.90 schreibt Bildausschnitt in Bauteile ein"
__doc__ = """Bildausschnitt schreiben
Kategorien: Rohr-, Luftkanalzubehör, Luftkanaldurchlässe
Parameter:IGF_X_Bildausschnitt: 79005440-b93e-4571-be64-41c4073bad97"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)


# MEP Räume aus aktueller Projekt
views_collector = FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType()
views = views_collector.ToElementIds()
views_collector.Dispose()


class Ansicht(object):
    def __init__(self):
        self.Checked = False
    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self, value):
        self._Name = value
    @property
    def Checked(self):
        return self._Checked
    @Checked.setter
    def Checked(self, value):
        self._Checked = value
    @property
    def ElementId(self):
        return self._ElementId
    @ElementId.setter
    def ElementId(self, value):
        self._ElementId = value
    @property
    def Ebene0(self):
        return self._Ebene0
    @Ebene0.setter
    def Ebene0(self, value):
        self._Ebene0 = value
    @property
    def Ebene1(self):
        return self._Ebene1
    @Ebene1.setter
    def Ebene1(self, value):
        self._Ebene1 = value
    @property
    def Ebene2(self):
        return self._Ebene2
    @Ebene2.setter
    def Ebene2(self, value):
        self._Ebene2 = value

dict_Ebene = {}
Liste_Ansicht = ObservableCollection[Ansicht]()

for id in views:
    elem = doc.GetElement(id)
    if elem.ViewType.ToString() in ['FloorPlan','CeilingPlan','ThreeD','Section']:
        name = elem.Name
        Ebene0 = elem.LookupParameter('000_000_152_Ansichtsgruppe').AsString()
        if not Ebene0:
            Ebene0 = '???'
        Ebene1 = elem.LookupParameter('Unterdisziplin').AsString()
        if not Ebene1:
            Ebene1 = '???'
        Ebene2 = elem.get_Parameter(BuiltInParameter.VIEW_TYPE).AsString()
        if not Ebene2:
            continue
        temp = Ansicht()
        temp.Name = name
        temp.Ebene0 = Ebene0
        temp.Ebene1 = Ebene1
        temp.Ebene2 = Ebene2
        temp.ElementId = id.ToString()
        Liste_Ansicht.Add(temp)
        if not Ebene0 in dict_Ebene.Keys:
            dict_Ebene[Ebene0] = {}
        if not Ebene1 in dict_Ebene[Ebene0]:
            dict_Ebene[Ebene0][Ebene1] = [Ebene2]
        if not Ebene2 in dict_Ebene[Ebene0][Ebene1]:
            dict_Ebene[Ebene0][Ebene1].append(Ebene2)



keys1 = dict_Ebene.Keys[:]
keys1.append('keine')
keys1.sort()
keys2 = []
keys3 = []

# GUI Pläne
class ViewUI(WPFWindow):
    def __init__(self, xaml_file_name,liste_View):
        self.liste_View = liste_View
        WPFWindow.__init__(self, xaml_file_name)
        self.dataGrid.ItemsSource = liste_View
        self.tempcoll = ObservableCollection[Ansicht]()

        self.altdatagrid = liste_View
        self.group.ItemsSource = keys1
        self.unterdis.ItemsSource = keys2
        self.viewtype.ItemsSource = keys3
        
        self.suche.TextChanged += self.auswahl_txt_changed
        self.group.SelectionChanged += self.auswahl_txt_changed
        self.group.SelectionChanged += self.auswahl_group_changed
        
        self.unterdis.SelectionChanged += self.auswahl_txt_changed
        self.unterdis.SelectionChanged += self.auswahl_unterdis_changed

        self.viewtype.SelectionChanged += self.auswahl_txt_changed

    def auswahl_group_changed(self,sender,args):
        group = ''
        temp_key2 = []
        temp_key3 = []
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass

        try:
            temp_key2 = dict_Ebene[group].Keys[:]
            temp_key2.append('keine')
            temp_key2.sort()
            self.unterdis.ItemsSource = temp_key2
            self.unterdis.Text = 'keine'
            self.viewtype.ItemsSource = temp_key3
        except:
            self.unterdis.ItemsSource = temp_key2
            self.unterdis.Text = ''
            self.viewtype.ItemsSource = temp_key3

    def auswahl_unterdis_changed(self,sender,args):
        unterdis = ''
        keys3 = []
        try:
            unterdis = self.unterdis.SelectedItem.ToString()
        except:
            pass
        group = ''
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            keys3 = dict_Ebene[group][unterdis][:]
            keys3.append('keine')
            keys3.sort()
            self.viewtype.Text = 'keine'
            self.viewtype.ItemsSource = keys3
        except:
            self.viewtype.ItemsSource = keys3

    
        
    def auswahl_txt_changed(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.suche.Text
        group = ''
        viewtype = ''
        unterdis = ''
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            viewtype = self.viewtype.SelectedItem.ToString()
        except:
            pass
        try:
            unterdis = self.unterdis.SelectedItem.ToString()
        except:
            pass
        
        
        if text_typ in ['',None]:
            text_typ = self.suche.Text = ''

        for item in self.altdatagrid:
            if item.Name.find(text_typ) != -1:
                if group in ['','keine']:
                    self.tempcoll.Add(item)
                else:
                    if item.Ebene0 == group:
                        if unterdis in ['','keine']:
                            self.tempcoll.Add(item)
                        else:
                            if unterdis == item.Ebene1:
                                if viewtype in ['','keine']:
                                    self.tempcoll.Add(item)
                                else:
                                    if viewtype == item.Ebene2:
                                        self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.Checked = True
        self.dataGrid.Items.Refresh()


    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.Checked = False
        self.dataGrid.Items.Refresh()


    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.Checked
            item.Checked = not value
        self.dataGrid.Items.Refresh()
        
    def select(self,sender,args):
        self.Close()

VIEW = ViewUI('window.xaml',Liste_Ansicht)
VIEW.ShowDialog()

View_dict = {}
for el in Liste_Ansicht:
    if el.Checked:
        View_dict[el.Name] = el.ElementId

element_dict = {}
def BAermitteln(ids,Cate,Ansich,dict):
    title = "{value}/{max_value} Bauteile in Kategorie " + Cate+ ' in Ansicht ' + Ansich
    step = int(len(ids)/100)+1
    with forms.ProgressBar(title=title, cancellable=True, step=step) as pb:

        for n,item in enumerate(ids):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n, len(ids))
            if not item.ToString() in dict.Keys:
                dict[item.ToString()] = [Ansich]
            else:
                dict[item.ToString()].append(Ansich)

for ele in View_dict.Keys:
    view = ElementId(int(View_dict[ele]))
    RZ_coll = FilteredElementCollector(doc,view).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType()
    RZ = RZ_coll.ToElementIds()
    RZ_coll.Dispose()
    LZ_coll = FilteredElementCollector(doc,view).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType()
    LZ = LZ_coll.ToElementIds()
    LZ_coll.Dispose()
    LD_coll = FilteredElementCollector(doc,view).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
    LD = LD_coll.ToElementIds()
    LD_coll.Dispose()
    BAermitteln(RZ,'Rohrzubehör',ele,element_dict)
    BAermitteln(LZ,'Luftkanalzubehör',ele,element_dict)
    BAermitteln(LD,'Luftdurchlässe',ele,element_dict)

if forms.alert("Daten schreiben?", ok=False, yes=True, no=True):
    t = Transaction(doc,"Bildausschnitte")
    t.Start()
    elems = element_dict.Keys[:]
    title = "{value}/{max_value} Bauteile"
    step = int(len(elems)/100)+1
    with forms.ProgressBar(title=title, cancellable=True, step=step) as pb:
        n = 0
        for n,item in enumerate(elems):
            if pb.cancelled:
                t.RollBack()
                script.exit()
            pb.update_progress(n, len(elems))
            elem = doc.GetElement(ElementId(int(item)))
            Text_liste = element_dict[item]
            text_temp = ''
            for e in Text_liste:
                text_temp = text_temp + e + ','
            Text = text_temp[:-1]
            try:
                elem.get_Parameter(Guid('79005440-b93e-4571-be64-41c4073bad97')).Set(Text)
            except Exception as e:
                logger.error(e)

    t.Commit()