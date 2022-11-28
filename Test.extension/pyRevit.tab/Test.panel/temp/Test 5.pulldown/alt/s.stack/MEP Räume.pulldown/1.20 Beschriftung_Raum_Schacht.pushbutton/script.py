# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms
from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection

__title__ = "1.20 Beschriftung Raum & Schacht"
__doc__ = """
Beschriftung für Raum und Schacht

imput Parameter:
MEP-Räme:
IGF_A_InstallationsGewerke: Schachtsart
L: Lüftungsschacht
S: Sanitärschacht
K: Kälteschacht
H: Heizungsschacht
GA: Automationsschacht
SP: Feuerlöschschacht
TG: Technische Gase Schacht
E: Elektroschacht

Koordinierte Schacht: Bsp. H, L: Heizungs- und Lüftungsschacht, default: nicht ausgewählt

Sontige Schächte: IGF_A_InstallationsGewerke = null, default: ausgewählt

TGA_RLT_InstallationsSchacht: Ja/Nein

[2021.11.23]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number

config = script.get_config(name+number+'Beschriftung Raum/Schacht')


beschriftung_collector = DB.FilteredElementCollector(doc).OfClass(DB.Family)
Beschriftungstype_dic = {' ': None}

for item in beschriftung_collector:
    category = item.FamilyCategory.Name
    if category == "MEP-Raumbeschriftungen":
        Tpy = item.GetFamilySymbolIds()
        for i in Tpy:
            famliyundtype = doc.GetElement(i).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
            Beschriftungstype_dic[famliyundtype] = i

Beschriftungstype_liste = Beschriftungstype_dic.keys()[:]
Beschriftungstype_liste.sort()
beschriftung_collector.Dispose()

views_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType()
views = views_collector.ToElementIds()
views_collector.Dispose()
view_dict = {}

for el in views:
    elem = doc.GetElement(el)
    typ = elem.ViewType.ToString()
    if elem.IsTemplate:
        continue
    if typ in ['FloorPlan','CeilingPlan']:
        name = elem.Name
        view_dict[name] = el

sorted(view_dict)

class Views(object):
    def __init__(self,elementid,name):
        self.ElementID = elementid
        self.Name = name
        self.checked = False
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value

Liste_Views = ObservableCollection[Views]()
for name in view_dict.keys():
    Liste_Views.Add(Views(view_dict[name],name))

class Beschriftungen(WPFWindow):
    def __init__(self, xaml_file_name,liste_Views):
        self.liste_Views = liste_Views
        
        self.tempcoll = ObservableCollection[Views]()
        WPFWindow.__init__(self, xaml_file_name)
        self.LV_view.ItemsSource = liste_Views

        self.set_comboboxes()
        self.read_config()
        self.ansichtsuche.TextChanged += self.such_txt_changed

    def read_config(self):
        try:
            if config.raum not in Beschriftungstype_dic.keys():
                config.raum = ""
            else:
                try:
                    self.raum.Text = config.raum
                except:
                    self.raum.Text = config.raum = ""
        except:
            pass

        try:
            if config.luft not in Beschriftungstype_dic.keys():
                config.luft = ""
            else:
                try:
                    self.luft.Text = config.luft
                except:
                    self.luft.Text = config.luft = ""
        except:
            pass
        
        try:
            if config.sani not in Beschriftungstype_dic.keys():
                config.sani = ""
            else:
                try:
                    self.sani.Text = config.sani
                except:
                    self.sani.Text = config.sani = ""
        except:
            pass

        try:
            if config.heiz not in Beschriftungstype_dic.keys():
                config.heiz = ""
            else:
                try:
                    self.heiz.Text = config.heiz
                except:
                    self.heiz.Text = config.heiz = ""
        except:
            pass

        try:
            if config.kalt not in Beschriftungstype_dic.keys():
                config.kalt = ""
            else:
                try:
                    self.kalt.Text = config.kalt
                except:
                    self.kalt.Text = config.kalt = ""
        except:
            pass

        try:
            if config.gase not in Beschriftungstype_dic.keys():
                config.gase = ""
            else:
                try:
                    self.gase.Text = config.gase
                except:
                    self.gase.Text = config.gase = ""
        except:
            pass

        try:
            if config.Auto not in Beschriftungstype_dic.keys():
                config.Auto = ""
            else:
                try:
                    self.Auto.Text = config.Auto
                except:
                    self.Auto.Text = config.Auto = ""
        except:
            pass

        try:
            if config.Feuer not in Beschriftungstype_dic.keys():
                config.Feuer = ""
            else:
                try:
                    self.Feuer.Text = config.Feuer
                except:
                    self.Feuer.Text = config.Feuer = ""
        except:
            pass

        try:
            if config.elek not in Beschriftungstype_dic.keys():
                config.elek = ""
            else:
                try:
                    self.elek.Text = config.elek
                except:
                    self.elek.Text = config.elek = ""
        except:
            pass

        try:
            if config.Koor not in Beschriftungstype_dic.keys():
                config.Koor = ""
            else:
                try:
                    self.Koor.Text = config.Koor
                except:
                    self.Koor.Text = config.Koor = ""
        except:
            pass

        try:
            if config.sonstige not in Beschriftungstype_dic.keys():
                config.sonstige = ""
            else:
                try:
                    self.sonstige.Text = config.sonstige
                except:
                    self.sonstige.Text = config.sonstige = ""
        except:
            pass

    def write_config(self):
        config.raum = self.raum.Text
        config.luft = self.luft.Text
        config.sani = self.sani.Text
        config.heiz = self.heiz.Text
        config.kalt = self.kalt.Text
        config.gase = self.gase.Text
        config.Auto = self.Auto.Text
        config.Feuer = self.Feuer.Text
        config.elek = self.elek.Text
        config.Koor = self.Koor.Text
        config.sonstige = self.sonstige.Text
        script.save_config()

    def set_comboboxes(self):
        self.raum.ItemsSource = Beschriftungstype_liste
        self.luft.ItemsSource = Beschriftungstype_liste
        self.sani.ItemsSource = Beschriftungstype_liste
        self.heiz.ItemsSource = Beschriftungstype_liste
        self.kalt.ItemsSource = Beschriftungstype_liste
        self.gase.ItemsSource = Beschriftungstype_liste
        self.Auto.ItemsSource = Beschriftungstype_liste
        self.Feuer.ItemsSource = Beschriftungstype_liste
        self.elek.ItemsSource = Beschriftungstype_liste
        self.Koor.ItemsSource = Beschriftungstype_liste
        self.sonstige.ItemsSource = Beschriftungstype_liste

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
    
    def such_txt_changed(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.ansichtsuche.Text.upper()

        if text_typ in ['',None]:
            self.LV_view.ItemsSource = self.liste_Views
            text_typ = self.ansichtsuche.Text = ''

        for item in self.liste_Views:
            if item.Name.upper().find(text_typ) != -1:
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

    def ok(self, sender, args):
        self.Close()
        self.write_config()

    def abbrechen(self, sender, args):
        self.Close()
        script.exit()

Beschriftung = Beschriftungen('window.xaml',Liste_Views)
try:
    Beschriftung.ShowDialog()
except Exception as e:
    Beschriftung.Close()
    logger.error(e)
    script.exit()

views_liste = []
for el in Liste_Views:
    if el.checked:
        
        views_liste.append(el)

if len(views_liste) == 0:
    logger.info('Keine Ansicht ausgewählt.')
    script.exit()

Schacht_Dict = {}
if Beschriftung.r.IsChecked:
    try:
        Schacht_Dict['R'] = Beschriftungstype_dic[Beschriftung.raum.Text]
    except:
        Schacht_Dict['R'] = None
        
if Beschriftung.l.IsChecked:
    try:
        Schacht_Dict['L'] = Beschriftungstype_dic[Beschriftung.luft.Text]
    except:
        Schacht_Dict['L'] = None

if Beschriftung.s.IsChecked:
    try:
        Schacht_Dict['S'] = Beschriftungstype_dic[Beschriftung.sani.Text]
    except:
        Schacht_Dict['S'] = None

if Beschriftung.h.IsChecked:
    try:
        Schacht_Dict['H'] = Beschriftungstype_dic[Beschriftung.heiz.Text]
    except:
        Schacht_Dict['H'] = None

if Beschriftung.k.IsChecked:
    try:
        Schacht_Dict['K'] = Beschriftungstype_dic[Beschriftung.kalt.Text]
    except:
        Schacht_Dict['K'] = None

if Beschriftung.g.IsChecked:
    try:
        Schacht_Dict['TG'] = Beschriftungstype_dic[Beschriftung.gase.Text]
    except:
        Schacht_Dict['TG'] = None

if Beschriftung.at.IsChecked:
    try:
        Schacht_Dict['GA'] = Beschriftungstype_dic[Beschriftung.Auto.Text]
    except:
        Schacht_Dict['GA'] = None

if Beschriftung.sp.IsChecked:
    try:
        Schacht_Dict['SP'] = Beschriftungstype_dic[Beschriftung.Feuer.Text]
    except:
        Schacht_Dict['SP'] = None

if Beschriftung.e.IsChecked:
    try:
        Schacht_Dict['E'] = Beschriftungstype_dic[Beschriftung.elek.Text]
    except:
        Schacht_Dict['E'] = None

if Beschriftung.ko.IsChecked:
    try:
        Schacht_Dict['KO'] = Beschriftungstype_dic[Beschriftung.Koor.Text]
    except:
        Schacht_Dict['KO'] = None

if Beschriftung.so.IsChecked:
    try:
        Schacht_Dict['SO'] = Beschriftungstype_dic[Beschriftung.sonstige.Text]
    except:
        Schacht_Dict['SO'] = None


class MEPRaum:
    def __init__(self,elemid,tag,view):
        self.elemid = elemid
        self.view = view
        self.Tag = tag
        self.elem  =doc.GetElement(self.elemid)
        self.IsSchacht = False
        self.Schachtart = ''
        self.Tagtyp = None
        self.Tagtypid = None
        self.Koor = False
        self.get_Raumtyp()
        self.get_Tagtyp()
        self.get_Tagtypid()
        self.CreateTag()
        self.changetagtyp()
    
    def get_Raumtyp(self):
        isschacht_Param = self.elem.LookupParameter("TGA_RLT_InstallationsSchacht")
        if isschacht_Param:
            isschacht = isschacht_Param.AsInteger()
            if isschacht:
                self.IsSchacht = True
                param_typ = self.elem.LookupParameter("IGF_A_InstallationsGewerke")
                if param_typ:
                    self.Schachtart = param_typ.AsString()
                else:
                    self.Schachtart = 'SO'
            else:
                self.IsSchacht = False
    
    def get_Tagtyp(self):
        if not self.IsSchacht:
            self.Tagtyp = 'R'
        else:
            if self.Schachtart == 'L':
                self.Tagtyp = 'L'
            elif self.Schachtart == 'S':
                self.Tagtyp = 'S'
            elif self.Schachtart == 'K':
                self.Tagtyp = 'K'
            elif self.Schachtart == 'H':
                self.Tagtyp = 'H'
            elif self.Schachtart == 'GA':
                self.Tagtyp = 'GA'
            elif self.Schachtart == 'SP':
                self.Tagtyp = 'SP'
            elif self.Schachtart == 'TG':
                self.Tagtyp = 'TG'
            elif self.Schachtart == 'E':
                self.Tagtyp = 'E'
            elif self.Schachtart == 'SO':
                self.Tagtyp = 'SO'
            else:
                self.Tagtyp = self.Schachtart
                self.Koor = True
    
    def get_Tagtypid(self):
        if self.IsSchacht:
            if self.Koor:
                if 'KO' in Schacht_Dict.keys():
                    self.Tagtypid = Schacht_Dict['KO']
                else:
                    for el in Schacht_Dict.keys():
                        if self.Schachtart.find('el') != -1:
                            self.Tagtypid = Schacht_Dict[el]
                            break

            else:
                try:
                    self.Tagtypid = Schacht_Dict[self.Tagtyp]
                except:
                    pass
        else:
            try:
                self.Tagtypid = Schacht_Dict[self.Tagtyp]
            except:
                pass

    
    def CreateTag(self):
        if self.Tag:
            return
        uv = DB.UV(self.elem.Location.Point.X,self.elem.Location.Point.Y)
        newtag = doc.Create.NewSpaceTag(self.elem,uv,self.view)
        self.Tag = newtag 
    
    def changetagtyp(self):
        if self.Tagtypid:
            self.Tag.ChangeTypeId(self.Tagtypid)


trans = DB.Transaction(doc,"MEP-Raum-Beschriftung")
trans.Start()
with forms.ProgressBar(cancellable=True, step=10) as pb:
    for n, view in enumerate(views_liste):
        mepraum_coll = DB.FilteredElementCollector(doc, view.ElementID).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
        mepraumids = mepraum_coll.ToElementIds()
        mepraum_coll.Dispose()
        tag_coll = DB.FilteredElementCollector(doc, view.ElementID).OfCategory(DB.BuiltInCategory.OST_MEPSpaceTags).WhereElementIsNotElementType()
        tag_dict = {}
        View = doc.GetElement(view.ElementID)
        for el in tag_coll:
            tag_dict[el.Space.Id.ToString()] = el
        pb.title='{value}/{max_value} MEP-Räume in Anschit ' +  view.Name + ' --- ' + str(n+1) + '/' + str(len(views_liste)) + 'Ansichts' 
        for n1, spaceid in enumerate(mepraumids):
            if pb.cancelled:
                trans.RollBack()
                script.exit()
            pb.update_progress(n1 + 1, len(mepraumids))
            try:
                tag = tag_dict[spaceid.ToString()]
            except:
                tag = None
            MEPRaum(spaceid,tag,View)
            
trans.Commit()
trans.Dispose()