# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit import revit, UI, DB, script, forms
import clr
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from pyIGF_config import Server_config

__title__ = "Pläne erstellen"
__doc__ = """Pläne erstellen (nur für Grundrisse)"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config_user = script.get_config('Plaene')
server = Server_config()
config_server = server.get_config('Plaene')
config_temp = script.get_config('Plaene_Erstellung')

uidoc = revit.uidoc
doc = revit.doc

Config_PR = 'Server'
try:
    if config_user.getconfig:
        config = config_server  
    else:
        config = config_user
        Config_PR = 'User'
except:
    config = config_server 

try:
    getlog(__title__)
except:
    pass

# read config
try:
    selected_Hauptansicht = config.HA_erstellung 
except:
    selected_Hauptansicht = ''

try:
    selected_Legende = config.LG_erstellung 
except:
    selected_Legende = ''

try:
    selected_Plankopf = config.erstellung_plankopf 
except:
    selected_Plankopf = ''

try:
    gewerke_config = config.gewerke_erstellung
except:
    gewerke_config = ''


Gewerke_t = gewerke_config.split(',')
Gewerke = []
for n in range(len(Gewerke_t)):
    while(Gewerke_t[n][0] == ' '):
        item = Gewerke_t[n][1:]
        Gewerke_t[n] = item
    while(Gewerke_t[n][-1] == ' '):
        item = Gewerke_t[n][:-1]
        Gewerke_t[n] = item
    Gewerke.append(Gewerke_t[n])
if not '' in Gewerke:
   Gewerke.insert(0,'') 


Gewerke_IS = ObservableCollection[str]()
def listtocoll(liste,coll):
    for item in liste:
        coll.Add(item)
def colltolist(coll,liste):
    for item in coll:
        liste.append(item)
listtocoll(Gewerke,Gewerke_IS)
# class für Itemssource in WPF
class Itemtemplate(object):
    def __init__(self,_index,_name):
        self.Name = _name
        self.selectindex = _index
    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self, value):
        self._Name = value
    @property
    def selectindex(self):
        return self._selectindex
    @selectindex.setter
    def selectindex(self, value):
        self._selectindex = value

plankopf_liste = []
plankopfid_dict = {}

# Planköpfe aus aktueller Projekt
coll_Plankopf = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
titleblock_family = []
titleblock = []

for el in coll_Plankopf:
    if el.FamilyCategoryId.ToString() == '-2000280':
        titleblock_family.append(el)

for el in titleblock_family:
    symids = el.GetFamilySymbolIds()
    for id in symids:
        name = doc.GetElement(id).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
        plankopf_liste.append(name)
        plankopfid_dict[name] = id
coll_Plankopf.Dispose()
plankopf_liste.sort()
plankopf_liste.insert(0,'')

plankopf_IS = ObservableCollection[Itemtemplate]()
for n,el in enumerate(plankopf_liste):
    plankopf_IS.Add(Itemtemplate(n,el))

Selected_Plankopf_Index = 0
for e in plankopf_IS:
    if e.Name == selected_Plankopf:
        Selected_Plankopf_Index = e.selectindex

# Bildausschnitt aus aktueller Projekt
Bildaus_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType()
Bildausschnitt = Bildaus_collector.ToElementIds()
Bildaus_collector.Dispose()
ausschnitt_liste = []
ausschnitt_dict_Name = {}

for el in Bildausschnitt:
    elem = doc.GetElement(el)
    family = elem.Name
    ausschnitt_dict_Name[family] = el
    ausschnitt_liste.append(family)

ausschnitt_liste.sort()
ausschnitt_liste.insert(0,'Keine')
ausschnitt_dict_Name['Keine'] = None
ausschnitt_id_name = {}
ausschnitt_name_id = {}

ausschnitt_dict_ID = {}
ausschnitt_IS = ObservableCollection[Itemtemplate]()
ausschnitt_Leer = ObservableCollection[Itemtemplate]()
for n,name in enumerate(ausschnitt_liste):
    ausschnitt_IS.Add(Itemtemplate(n,name))
    ausschnitt_dict_ID[n] = ausschnitt_dict_Name[name]
    ausschnitt_id_name[name] = n
    ausschnitt_name_id[n] = name

# Views aus aktueller Projekt
views_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType()
views = views_collector.ToElementIds()
views_collector.Dispose()
view_liste = []
Legenden_liste = []
Legenden_dict_Name  = {}

for el in views:
    elem = doc.GetElement(el)
    typ = elem.ViewType.ToString()
    if elem.IsTemplate:
        continue
    if typ in ['FloorPlan','CeilingPlan']:
        view_liste.append(el)
    elif typ == 'Legend':
        Legenden_liste.append(elem.Name)
        Legenden_dict_Name[elem.Name] = el

Legenden_liste.sort()
Legenden_liste.insert(0,'Keine')
Legenden_dict_Name['Keine'] = None
Legenden_dict_ID = {}
Legenden_id_name = {}
# ItemsSource für Legenden
Legende_IS = ObservableCollection[Itemtemplate]()
for n,name in enumerate(Legenden_liste):
    Legende_IS.Add(Itemtemplate(n,name))
    Legenden_dict_ID[n] = Legenden_dict_Name[name]
    Legenden_id_name[name] = n

# Viewports aus aktueller Projekt
Viewport_liste = []
Viewport_dict_Name = {}
Viewport_id_name = {}
coll1 = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ElementType))
for el in coll1:
    if el.FamilyName == 'Ansichtsfenster':
        name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        Viewport_liste.append(name)
        Viewport_dict_Name[name] = el.Id
Viewport_liste.sort()

Viewport_IS = ObservableCollection[Itemtemplate]()
Viewport_dict_ID = {}
for i,name in enumerate(Viewport_liste):
    Viewport_IS.Add(Itemtemplate(i,name))
    Viewport_dict_ID[i] = Viewport_dict_Name[name]
    Viewport_id_name[name] = i


global selected_HA_ID
global selected_LG_ID
selected_HA_ID = 0
selected_LG_ID = 0
if selected_Hauptansicht:
    if selected_Hauptansicht in Viewport_id_name.keys():
        selected_HA_ID = Viewport_id_name[selected_Hauptansicht]
if selected_Legende:
    if selected_Legende in Viewport_id_name.keys():
        selected_LG_ID = Viewport_id_name[selected_Legende]


class Views(object):
    def __init__(self,elementid):
        self.viewport_IS = Viewport_IS
        self.bildausschnitt_IS = ausschnitt_IS
        self.legende_IS = Legende_IS

        self.ElementID = elementid
        self.checked = False
        self.ansichtname = ''
        self.Bildausschnitt_id = ''
        self.HA_id = ''
        self.legend_id = 0
        self.LG_id = ''
        self.gewerke = ''
        self.Ebene = False
        self.level = ''
        self.Freitext = ''
        self.planname = ''
        self.Ansicht = False
        self.Bild = False

        self.Group = ''
        self.Mcate = ''
        self.Subcate = ''
        try:
            self.level = doc.GetElement(self.ElementID).LookupParameter('Verknüpfte Ebene').AsString()
        except:
            pass

    @property
    def ElementID(self):
        return self._ElementID
    @ElementID.setter
    def ElementID(self, value):
        self._ElementID = value
    @property
    def Ansicht(self):
        return self._Ansicht
    @Ansicht.setter
    def Ansicht(self, value):
        self._Ansicht = value
    @property
    def Bild(self):
        return self._Bild
    @Bild.setter
    def Bild(self, value):
        self._Bild = value
    @property
    def gewerke(self):
        return self._gewerke
    @gewerke.setter
    def gewerke(self, value):
        self._gewerke = value
    @property
    def Ebene(self):
        return self._Ebene
    @Ebene.setter
    def Ebene(self, value):
        self._Ebene = value
    @property
    def Freitext(self):
        return self._Freitext
    @Freitext.setter
    def Freitext(self, value):
        self._Freitext = value
    @property
    def planname(self):
        return self._planname
    @planname.setter
    def planname(self, value):
        self._planname = value

    @property
    def Group(self):
        return self._Group
    @Group.setter
    def Group(self, value):
        self._Group = value
    @property
    def Mcate(self):
        return self._Mcate
    @Mcate.setter
    def Mcate(self, value):
        self._Mcate = value
    @property
    def Subcate(self):
        return self._Subcate
    @Subcate.setter
    def Subcate(self, value):
        self._Subcate = value

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def ansichtname(self):
        return self._ansichtname
    @ansichtname.setter
    def ansichtname(self, value):
        self._ansichtname = value
    @property
    def Bildausschnitt_id(self):
        return self._Bildausschnitt_id
    @Bildausschnitt_id.setter
    def Bildausschnitt_id(self, value):
        self._Bildausschnitt_id = value  
    @property
    def HA_id(self):
        return self._HA_id
    @HA_id.setter
    def HA_id(self, value):
        self._HA_id = value
    @property
    def legend_id(self):
        return self._legend_id
    @legend_id.setter
    def legend_id(self, value):
        self._legend_id = value
    @property
    def LG_id(self):
        return self._LG_id
    @LG_id.setter
    def LG_id(self, value):
        self._LG_id = value


Liste_Views = ObservableCollection[Views]()
Auswahldict = {}
planname_temp =  ObservableCollection[str]()

for viewid in view_liste:
    elem = doc.GetElement(viewid)
    tempclass = Views(viewid)
    tempclass.ansichtname = elem.Name

    tempclass.HA_id = selected_HA_ID
    tempclass.LG_id = selected_LG_ID

    param = elem.LookupParameter('Bildausschnitt')

    if param:
        try:
            tempclass.Bildausschnitt_id = ausschnitt_id_name[param.AsValueString()]
        except:
            tempclass.Bildausschnitt_id = 0
    else:
        tempclass.bildausschnitt_IS = ausschnitt_Leer

    try:
        group = elem.LookupParameter('000_000_152_Ansichtsgruppe').AsString()
        if not group:
            group = '???'
        tempclass.Group = group
        mcate = elem.LookupParameter('Unterdisziplin').AsString()
        if not mcate:
            mcate = '???'
        tempclass.Mcate = mcate
        scate = elem.get_Parameter(DB.BuiltInParameter.VIEW_TYPE).AsString()
        if not scate:
            scate = '???'
        tempclass.Subcate = scate
        if not group in Auswahldict.Keys:
            Auswahldict[group] = {}
        if not mcate in Auswahldict[group].Keys:
            Auswahldict[group][mcate] = []
        if not scate in Auswahldict[group][mcate]:
            Auswahldict[group][mcate].append(scate)
    except Exception as e:
        logger.error(e) 
    Liste_Views.Add(tempclass)

keys1 = Auswahldict.Keys[:]
keys1.sort()
keys1.insert(0,'Keine')

keys2 = []
keys3 = []

global create 
create = False

class PlaeneUI(WPFWindow):
    def __init__(self, xaml_file_name,liste_views):
        self.liste_views = liste_views
        WPFWindow.__init__(self, xaml_file_name)
        self.LV_view.ItemsSource = liste_views
        self.tempcoll = ObservableCollection[Views]()
        self.leercoll = ObservableCollection[Views]()

        self.Plankopf.ItemsSource = plankopf_IS
        self.altListview = liste_views
        self.group.ItemsSource = keys1
        self.main.ItemsSource = keys2
        self.sub.ItemsSource = keys3
        self.GW.ItemsSource = Gewerke_IS

        self.read_config()

        self.group.SelectionChanged += self.auswahl_txt_changed
        self.main.SelectionChanged += self.auswahl_txt_changed
        self.sub.SelectionChanged += self.auswahl_txt_changed


    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LV_view.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_view.SelectedItems:
                    for item in self.LV_view.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.LV_view.Items.Refresh()
                else:
                    pass
            except:
                pass

    def formchange(self):
        try:
            if self.LV_view.SelectedItem is not None:
                gewerke_temp = []
                ebene_temp = []
                ansicht_temp = []
                bildaus_temp = []
                Freitext_temp = []
                gewerke_pr = False
                ebene_pr = False
                Freitext_pr = False
                ansicht_pr = False
                bildaus_pr = False
                form_text_temp = ''
                Views.Ansicht
                for item in self.LV_view.SelectedItems:
                    Planname_temp = ''
                    if item.Ansicht:
                        Planname_temp = item.ansichtname + ' - '
                    else:
                        pass

                    if item.gewerke:
                        Planname_temp = Planname_temp + item.gewerke + ' - '
                    else:
                        pass
                    if item.Ebene:
                        try:
                            Planname_temp = Planname_temp + item.level + ' - '
                        except:
                            pass
                    else:
                        pass
                   
                    if item.Bild:
                        Planname_temp = Planname_temp + ausschnitt_name_id[item.Bildausschnitt_id] + ' - '
                    else:
                        pass


                    if item.Freitext:
                        Planname_temp = Planname_temp + item.Freitext + ' - '
                    else:
                        pass

                    planname = Planname_temp[:-3]                       
                    item.planname = planname 
                
                    gewerke_temp.append(item.gewerke)
                    ebene_temp.append(item.Ebene)
                    Freitext_temp.append(item.Freitext)
                    bildaus_temp.append(item.Bild)
                    ansicht_temp.append(item.Ansicht)
                
                if all(bildaus_temp) == True:
                    bildaus_pr = True
                    self.BA.IsChecked = True
                else:
                    self.BA.IsChecked = False
                
                if all(ansicht_temp) == True:
                    ansicht_pr = True
                    self.AN.IsChecked = True
                else:
                    self.AN.IsChecked = False

                if len(set(gewerke_temp)) == 1:
                    if gewerke_temp[0] != '':
                        gewerke_pr = True
                        self.GW.Text = gewerke_temp[0]
                    else:
                        self.GW.Text = ''
                else:
                    self.GW.Text = ''

                if all(ebene_temp) == True:
                    ebene_pr = True
                    self.EB.IsChecked = True
                else:
                    self.EB.IsChecked = False

                if len(set(Freitext_temp)) == 1:
                    if Freitext_temp[0] != '':
                        Freitext_pr = True
                        self.freitext.Text = Freitext_temp[0]
                    else:
                        self.freitext.Text = ''
                else:
                    self.freitext.Text = ''
                
                if ansicht_pr:
                    form_text_temp = 'Ansichtname - '
                else:
                    pass

                if gewerke_pr:
                    form_text_temp = form_text_temp + 'Gewerke - '
                else:
                    pass

                if ebene_pr:
                    form_text_temp = form_text_temp + 'Ebene - '
                else:
                    pass

                if bildaus_pr:
                    form_text_temp = form_text_temp + 'Bildauschnitt - '
                else:
                    pass
                
                if Freitext_pr:
                    form_text = form_text_temp + 'Freitext'
                else:
                    form_text = form_text_temp[:-3]
        
                self.form.Text = form_text

            else:
                self.form.Text = ''
        except:
            pass

    def Gechanged(self, sender, args):
        try:
            gewerke = sender.SelectedItem.ToString()
            if self.LV_view.SelectedItem is not None:
                try:
                    for item in self.LV_view.SelectedItems:
                        item.gewerke = gewerke   
                except:
                    pass
            else:
                sender.SelectedItem =  ''
                UI.TaskDialog.Show('Fehler','Kein Ansicht ausgewählt!')
            self.formchange()
            self.LV_view.Items.Refresh()
            
        except:
            pass

    def ebenecheckchanged(self, sender, args):
        try:
            Checked = sender.IsChecked
            if self.LV_view.SelectedItem is not None:
                try:
                    for item in self.LV_view.SelectedItems:
                        item.Ebene = Checked
                except:
                    pass
            else:
                sender.IsChecked = not Checked
                UI.TaskDialog.Show('Fehler','Kein Ansicht ausgewählt!')
            self.formchange()
            self.LV_view.Items.Refresh()
        except:
            pass
    def listviewslchanged(self, sender, args):
        self.formchange()
        pass
    def Freitextchanged(self, sender, args):
        try:
            text = sender.Text
            if self.LV_view.SelectedItem is not None:
                try:
                    for item in self.LV_view.SelectedItems:
                        item.Freitext = text
                except:
                    pass
            else:
              
                UI.TaskDialog.Show('Fehler','Kein Ansicht ausgewählt!')
            self.formchange()
            self.LV_view.Items.Refresh()
        except:
            pass

    def ansichtnamecheckchanged(self, sender, args):
        try:
            Checked = sender.IsChecked

            if self.LV_view.SelectedItem:
                try:
                    for item in self.LV_view.SelectedItems:
                        item.Ansicht = Checked
                except:
                    pass
            else:
                sender.IsChecked = not Checked
                UI.TaskDialog.Show('Fehler','Kein Ansicht ausgewählt!')
            self.formchange()
            self.LV_view.Items.Refresh()
        except:
            pass
    
    def bildausschnittcheckchanged(self, sender, args):
        try:
            Checked = sender.IsChecked
            if self.LV_view.SelectedItem is not None:
                try:
                    for item in self.LV_view.SelectedItems:
                        item.Bild = Checked
                except:
                    pass
            else:
                sender.IsChecked = not Checked
                UI.TaskDialog.Show('Fehler','Kein Ansicht ausgewählt!')
            self.formchange()
            self.LV_view.Items.Refresh()
        except:
            pass

    def additem(self, sender, args):
        temp_liste = []
        colltolist(self.GW.ItemsSource,temp_liste)
        if not self.GW.Text in temp_liste:
            Gewerke_IS.Add(self.GW.Text)
        try:
            gewerke = self.GW.Text
            if self.LV_view.SelectedItem is not None:
                try:
                    for item in self.LV_view.SelectedItems:
                        item.gewerke = gewerke   
                except:
                    pass
            self.formchange()
            self.LV_view.Items.Refresh()
        except:
            pass

    def BAS_changed(self, sender, args):
        SelectedIndex = sender.SelectedIndex
      ##  print(sender.DiaplayMemberPath)
        if self.LV_view.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_view.SelectedItems:
                    for item in self.LV_view.SelectedItems:
                        try:
                            item.Bildausschnitt_id = SelectedIndex
                        except:
                            pass
                    self.LV_view.Items.Refresh()
                else:
                    pass
            except:
                pass

    def HAF_changed(self, sender, args):
        SelectedIndex = sender.SelectedIndex
        if self.LV_view.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_view.SelectedItems:
                    for item in self.LV_view.SelectedItems:
                        try:
                            item.HA_id = SelectedIndex
                        except:
                            pass
                    self.LV_view.Items.Refresh()
                else:
                    pass
            except:
                pass

    def LG_changed(self, sender, args):
        SelectedIndex = sender.SelectedIndex
        if self.LV_view.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_view.SelectedItems:
                    for item in self.LV_view.SelectedItems:
                        try:
                            item.legend_id = SelectedIndex
                        except:
                            pass
                    self.LV_view.Items.Refresh()
                else:
                    pass
            except:
                pass

    def LGF_changed(self, sender, args):
        SelectedIndex = sender.SelectedIndex
        if self.LV_view.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_view.SelectedItems:
                    for item in self.LV_view.SelectedItems:
                        try:
                            item.LG_id = SelectedIndex
                        except:
                            pass
                    self.LV_view.Items.Refresh()
                else:
                    pass
            except:
                pass

    def auswahl_group_changed(self,sender,args):
        group = ''
        temp_key2 = []
        temp_key3 = []
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            temp_key2 = Auswahldict[group].Keys[:]
            temp_key2.sort()
            temp_key2.insert(0,'Keine')
    
            self.main.ItemsSource = temp_key2
            self.main.Text = 'Keine'
            self.sub.ItemsSource = temp_key3
        except:
            self.sub.ItemsSource = temp_key3

    def auswahl_main_changed(self,sender,args):
        main = ''
        try:
            main = self.main.SelectedItem.ToString()
        except:
            pass
        group = ''
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            keys3 = Auswahldict[group][main][:]
            keys3.sort()
            keys3.insert(0,'Keine')
    
            self.sub.Text = 'Keine'
            self.sub.ItemsSource = keys3
        except:
            self.sub.ItemsSource = []
            self.sub.Text = ''

    def read_config(self):
        try:
            self.bz_l_erstellung.Text = config.bz_l_erstellung
        except:
            self.bz_l_erstellung.Text = config.bz_l_erstellung = '10'
        try:
            self.bz_r_erstellung.Text = config.bz_r_erstellung
        except:
            self.bz_r_erstellung.Text = config.bz_r_erstellung = '10'
        try:
            self.bz_o_erstellung.Text = config.bz_o_erstellung
        except:
            self.bz_o_erstellung.Text = config.bz_o_erstellung = '10'
        try:
            self.bz_u_erstellung.Text = config.bz_u_erstellung
        except:
            self.bz_u_erstellung.Text = config.bz_u_erstellung = '10'

        try:
            self.pk_l_erstellung.Text = config.pk_l_erstellung
        except:
            self.pk_l_erstellung.Text = config.pk_l_erstellung = '20'
        try:
            self.pk_r_erstellung.Text = config.pk_r_erstellung
        except:
            self.pk_r_erstellung.Text = config.pk_r_erstellung = '5'
        try:
            self.pk_o_erstellung.Text = config.pk_o_erstellung
        except:
            self.pk_o_erstellung.Text = config.pk_o_erstellung = '5'
        try:
            self.pk_u_erstellung.Text = config.pk_u_erstellung
        except:
            self.pk_u_erstellung.Text = config.pk_u_erstellung = '5'

        try:
            if config.erstellung_plankopf in plankopf_liste:
                self.Plankopf.SelectedIndex = Selected_Plankopf_Index
            else:
                self.Plankopf.Text = config.erstellung_plankopf = ''
        except:
            self.Plankopf.Text = config.erstellung_plankopf = ''

    def write_config(self):
        try:
            config_temp.bz_l_erstellung = self.bz_l_erstellung.Text
        except:
            pass

        try:
            config_temp.bz_r_erstellung = self.bz_r_erstellung.Text
        except:
            pass

        try:
            config_temp.bz_o_erstellung = self.bz_o_erstellung.Text
        except:
            pass

        try:
            config_temp.bz_u_erstellung = self.bz_u_erstellung.Text
        except:
            pass

        try:
            config_temp.pk_u_erstellung = self.pk_u_erstellung.Text
        except:
            pass

        try:
            config_temp.pk_o_erstellung = self.pk_o_erstellung.Text
        except:
            pass

        try:
            config_temp.pk_l_erstellung = self.pk_l_erstellung.Text
        except:
            pass

        try:
            config_temp.pk_r_erstellung = self.pk_r_erstellung.Text
        except:
            pass

        try:
            config_temp.erstellung_plankopf = self.Plankopf.SelectedItem.Name
        except:
            config_temp.erstellung_plankopf = ''

        script.save_config()


    def auswahl_txt_changed(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.ansichtsuche.Text.upper()
        group = ''
        sub = ''
        main = ''
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            sub = self.sub.SelectedItem.ToString()
        except:
            pass
        try:
            main = self.main.SelectedItem.ToString()
        except:
            pass

        if not group:
            group = self.group.Text = ''
        if not sub:
            sub = self.sub.Text = ''
        if not main:
            main = self.main.Text = ''

        if text_typ in ['',None]:
            self.LV_view.ItemsSource = self.altListview
            text_typ = self.ansichtsuche.Text = ''

        for item in self.altListview:
            if item.ansichtname.upper().find(text_typ) != -1:
                if group in ['','Keine']:
                    self.tempcoll.Add(item)
                else:
                    if item.Group == group:
                        if main in ['','Keine']:
                            self.tempcoll.Add(item)
                        else:
                            if main == item.Mcate:
                                if sub in ['','Keine']:
                                    self.tempcoll.Add(item)
                                else:
                                    if sub == item.Subcate:
                                        self.tempcoll.Add(item)
            self.LV_view.ItemsSource = self.tempcoll
        self.LV_view.Items.Refresh()

    def check(self,sender,args):
        for item in self.LV_view.Items:
            item.checked = True
        self.LV_view.Items.Refresh()


    def uncheck(self,sender,args):
        for item in self.LV_view.Items:
            item.checked = False
        self.LV_view.Items.Refresh()


    def toggle(self,sender,args):
        for item in self.LV_view.Items:
            value = item.checked
            item.checked = not value
        self.LV_view.Items.Refresh()

    def aktu(self,sender,args):

        global create 
        create = True
        try:
            plankopf = self.Plankopf.SelectedItem.Name
        except:
            plankopf = ''

        if not plankopf:
            UI.TaskDialog.Show('Fehler','Kein Plankopf ausgewählt')
        else:
            self.write_config()
            self.Close()

    def close(self,sender,args):
        self.Close()

Planfenster = PlaeneUI("window.xaml",Liste_Views)
try:
    Planfenster.ShowDialog()
except Exception as e:
    Planfenster.Close()
    logger.error(e)
    script.exit()

if not create:
    script.exit()

Liste_Plaene = []
for ele in Liste_Views:
    if ele.checked:
        Liste_Plaene.append(ele)

if any(Liste_Plaene):
    t_raster = DB.Transaction(doc,'Ansicht anpassen')
    t_raster.Start()
    with forms.ProgressBar(title='{value}/{max_value} Ansichten ausgewählt',cancellable=True, step=1) as pb:
        for n, elem in enumerate(Liste_Plaene):
            if pb.cancelled:
                t_raster.RollBack()
                script.exit()
            pb.update_progress(n+1, len(Liste_Plaene))
            view = doc.GetElement(elem.ElementID)

            # Bildausschnitt
            if elem.Bildausschnitt_id > 0:
                param = view.LookupParameter('Bildausschnitt')
                if param:
                    try:
                        param.Set(ausschnitt_dict_ID[elem.Bildausschnitt_id])
                    except:
                        logger.error('Fehler beim Ändern des Bildausschnittes der Ansicht {}.'.format(elem.ansichtname))

            # Zuschneidbereich
            try:
                cropbox = view.GetCropRegionShapeManager()
                if config_temp.bz_l_erstellung:
                    try:
                        cropbox.LeftAnnotationCropOffset = float(config_temp.bz_l_erstellung) / 304.8
                    except:
                        pass
                if config_temp.bz_r_erstellung:
                    try:
                        cropbox.RightAnnotationCropOffset = float(config_temp.bz_r_erstellung) / 304.8
                    except:
                        pass
                if config_temp.bz_o_erstellung:
                    try:
                        cropbox.TopAnnotationCropOffset = float(config_temp.bz_o_erstellung) / 304.8
                    except:
                        pass
                if config_temp.bz_u_erstellung:
                    try:
                        cropbox.BottomAnnotationCropOffset = float(config_temp.bz_u_erstellung) / 304.8
                    except:
                        pass
            except:
                pass

            doc.Regenerate()

            # Raster anpassen
            rasters_collector = DB.FilteredElementCollector(doc,elem.ElementID).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType()
            rasters = rasters_collector.ToElementIds()
            rasters_collector.Dispose()
            box = view.get_BoundingBox(view)
            max_X = box.Max.X
            max_Y = box.Max.Y
            min_X = box.Min.X
            min_Y = box.Min.Y
            for rasid in rasters:
                raster = doc.GetElement(rasid)
                raster.Pinned = False
                gridCurves = raster.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)
                if not gridCurves:
                    continue
                for gridCurve in gridCurves:
                    start = gridCurve.GetEndPoint( 0 )
                    end = gridCurve.GetEndPoint( 1 )
                    X1 = start.X
                    Y1 = start.Y
                    Z1 = start.Z
                    X2 = end.X
                    Y2 = end.Y
                    Z2 = end.Z
                    newStart = None
                    newEnd = None
                    newLine = None
                    if abs(X1-X2) > 1:
                        newStart = DB.XYZ(max_X,Y1,Z1)
                        newEnd = DB.XYZ(min_X,Y2,Z2)
                    if abs(Y1-Y2) > 1:
                        newStart = DB.XYZ(X1,max_Y,Z1)
                        newEnd = DB.XYZ(X2,min_Y,Z2)
                    if all([newStart,newEnd]):
                        newLine = DB.Line.CreateBound( newStart, newEnd )
                    if newLine:
                        raster.SetCurveInView(DB.DatumExtentType.ViewSpecific, view, newLine )
                raster.Pinned = True
    doc.Regenerate()
    t_raster.Commit()

    t_Plan = DB.Transaction(doc,'Pläne erstellen')
    t_Plan.Start()
    with forms.ProgressBar(title='{value}/{max_value} Pläne erstellt',cancellable=True, step=1) as pb1:
        try:
            plankopfid = plankopfid_dict[config_temp.erstellung_plankopf]
        except:
            t_Plan.RollBack()
            logger.error('Plankopf kann nicht gefunden werden')
        for n1, elem1 in enumerate(Liste_Plaene):
            if pb1.cancelled:
                t_Plan.RollBack()
                script.exit()
            pb1.update_progress(n1+1, len(Liste_Plaene))
            viewsheet = DB.ViewSheet.Create(doc,plankopfid)
            #sheetname ändern
            if elem1.planname:
                viewsheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).Set(elem1.planname)

            location = DB.XYZ((viewsheet.Outline.Max.U+viewsheet.Outline.Min.U)/2,(viewsheet.Outline.Max.V+viewsheet.Outline.Min.V)/2,0)
            if elem1.legend_id > 0:
                try:
                    viewport1 = DB.Viewport.Create(doc,viewsheet.Id,Legenden_dict_ID[elem1.legend_id],location)
                    try:
                        typeid = Viewport_dict_ID[elem1.LG_id]
                        viewport1.ChangeTypeId(typeid)
                    except:
                        logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Legende von Ansicht {}.'.format(elem1.ansichtname))
                        print(30*'-')
                    try:
                        x1_move = viewsheet.Outline.Max.U - viewport1.get_BoundingBox(viewsheet).Max.X - float(config_temp.pk_r_erstellung) / 304.8 
                        y1_move = viewsheet.Outline.Max.V - viewport1.get_BoundingBox(viewsheet).Max.Y - float(config_temp.pk_o_erstellung) / 304.8
                        xyz1_move = DB.XYZ(x1_move,y1_move,0)
                        viewport1.Location.Move(xyz1_move)
                    except:
                        logger.error('Fehler beim Verschieben der Legende von Ansicht {}.'.format(elem1.ansichtname))
                        print(30*'-')
                except Exception as e:
                    logger.error('Fehler beim Erstellen der Legende von Ansicht {}.'.format(elem1.ansichtname))  
                    print(30*'-')

            try:
                viewport = DB.Viewport.Create(doc,viewsheet.Id,elem1.ElementID,location)
                try:
                    typeid = Viewport_dict_ID[elem1.HA_id]
                    viewport.ChangeTypeId(typeid)
                except:
                    logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Hauptansicht {}'.format(elem1.ansichtname))
                    print(30*'-')
                try:
                    x_move = viewport.get_BoundingBox(viewsheet).Max.X + float(config_temp.pk_l_erstellung) / 304.8 - viewsheet.Outline.Max.U
                    y_move = viewsheet.Outline.Max.V - viewport.get_BoundingBox(viewsheet).Max.Y - float(config_temp.pk_o_erstellung) / 304.8
                    xyz_move = DB.XYZ(x_move,y_move,0)
                    viewport.Location.Move(xyz_move)
                except:
                    logger.error('Fehler beim Verschieben der Hauptansicht {}'.format(elem1.ansichtname))
                    print(30*'-')
            except:
                logger.error('Fehler beim Erstellen der Hauptansicht {}. Pläne wird gelöscht.'.format(elem1.ansichtname))
                print(30*'-')
                doc.Delete(viewsheet.Id)
    t_Plan.Commit()