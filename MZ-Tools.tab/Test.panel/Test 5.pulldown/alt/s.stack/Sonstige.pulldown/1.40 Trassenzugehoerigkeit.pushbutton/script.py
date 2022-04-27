# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection


__title__ = "1.50 Geschosszugehörigkeit"
__doc__ = """
schreibt Ebenename in Bauteile.

Parameter:
IGF_Trassenzugehörigkeit

Kategorie:
HLS-Bauteile,
Luftkanäle,
Rohre,
Kabeltrassen,
Leerrohr,
Luftdurchlässe,
Luftkanalzubehör,
Rohrzubehör


[2021.11.25]
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

config = script.get_config(name+number+'Geschosszugehörigkeit')

HLS_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
hlsids = HLS_collector.ToElementIds()
HLS_collector.Dispose()

Rohre_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()
Rohre = Rohre_collector.ToElementIds()
Rohre_collector.Dispose()

Luft_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType()
Luft = Luft_collector.ToElementIds()
Luft_collector.Dispose()

Terminal_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
Terminal = Terminal_collector.ToElementIds()
Terminal_collector.Dispose()

Kabel_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_CableTray).WhereElementIsNotElementType()
Kabel = Kabel_collector.ToElementIds()
Kabel_collector.Dispose()

Leer_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Conduit).WhereElementIsNotElementType()
Leer = Leer_collector.ToElementIds()
Leer_collector.Dispose()

Kanalzu_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType()
Kanalzu = Kanalzu_collector.ToElementIds()
Kanalzu_collector.Dispose()

Rohrzu_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType()
Rohrzu = Rohrzu_collector.ToElementIds()
Rohrzu_collector.Dispose()

Levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
levels = Levels_collector.ToElementIds()
Levels_collector.Dispose()

class Ebenen(object):
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.name = self.elem.Name
        self.name_neu = ''

    @property
    def name(self):
        return self._name
    @name.setter
    def name(self, value):
        self._name = value

    @property
    def name_neu(self):
        return self._name_neu
    @name_neu.setter
    def name_neu(self, value):
        self._name_neu = value

dict_ebene = {}
for elid in levels:
    name = doc.GetElement(elid).Name
    dict_ebene[name] = elid

liste_ebene = dict_ebene.keys()[:]
liste_ebene.sort()

Liste = ObservableCollection[Ebenen]()
for name in liste_ebene:
    Liste.Add(Ebenen(dict_ebene[name]))
    
class Ebeneumbenennen(WPFWindow):
    def __init__(self, xaml_file_name,liste):
        self.Liste = liste
        WPFWindow.__init__(self, xaml_file_name)
        self.LV_Ebene.ItemsSource = self.Liste
        self.read_config()
    
    def read_config(self):
        _dict = {}
        try:
            for item in config.liste:
                _dict[item[0]] = item[1]
            for item1 in self.Liste:
                if item1.name in _dict.keys():
                    item1.name_neu = _dict[item1.name]
            self.LV_Ebene.Items.Refresh()
        except:
            config.liste = []
        
    def write_config(self):
        liste = []
        for el in self.Liste:
            liste.append([el.name,el.name_neu])
        config.liste = liste
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        self.Close()
    def abbrechen(self,sender,args):
        self.Close()
        script.exit()

ebeneumbenennen = Ebeneumbenennen('window.xaml',Liste)
try:
    ebeneumbenennen.ShowDialog()
except Exception as e:
    ebeneumbenennen.Close()
    logger.error(e)
    script.exit()

ebene_dict_mit_name = {}
for el in Liste:
    ebene_dict_mit_name[el.name] = el.name_neu

class Leitung:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.ebene = self.elem.ReferenceLevel
        if self.ebene:
            self.ebenename = self.ebene.Name
        else:
            self.ebenename = ''
            logger.error('Element {} hat kein Ebene zugewiesen.'.format(self.elemid))
        if self.ebenename:
            try:
                self.ebenename_neu = ebene_dict_mit_name[self.ebenename]
            except:
                self.ebenename_neu = self.ebenename
        else:
            self.ebenename_neu = ''
    def wert_schreiben(self):
        param = self.elem.LookupParameter('IGF_Trassenzugehörigkeit')
        if param:
            param.Set(self.ebenename_neu)

class Bauteile:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.ebene = doc.GetElement(self.elem.LevelId)
        if self.ebene:
            self.ebenename = self.ebene.Name
        else:
            self.ebenename = ''
            logger.error('Element {} hat kein Ebene zugewiesen.'.format(self.elemid))
        if self.ebenename:
            try:
                self.ebenename_neu = ebene_dict_mit_name[self.ebenename]
            except:
                self.ebenename_neu = self.ebenename
        else:
            self.ebenename_neu = ''
    def wert_schreiben(self):
        param = self.elem.LookupParameter('IGF_Trassenzugehörigkeit')
        if param:
            param.Set(self.ebenename_neu)

class Auslass:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.ebene = doc.GetElement(self.elem.LevelId)
        
        if not self.ebene:
            self.ebene = self.elem.Host
        if self.ebene:
            self.ebenename = self.ebene.Name
        else:
            self.ebenename = ''
            logger.error('Element {} hat kein Ebene zugewiesen.'.format(self.elemid))
        if self.ebenename:
            try:
                self.ebenename_neu = ebene_dict_mit_name[self.ebenename]
            except:
                self.ebenename_neu = self.ebenename
        else:
            self.ebenename_neu = ''
    def wert_schreiben(self):
        param = self.elem.LookupParameter('IGF_Trassenzugehörigkeit')
        if param:
            param.Set(self.ebenename_neu)


kategorien_dict = {}
if ebeneumbenennen.hls.IsChecked:
    kategorien_dict["HLS-Bauteile"] = hlsids
if ebeneumbenennen.rohr.IsChecked:
    kategorien_dict["Rohre"] = Rohre
if ebeneumbenennen.luft.IsChecked:
    kategorien_dict["Luftkanäle"] = Luft
if ebeneumbenennen.kabel.IsChecked:
    kategorien_dict["Kabeltrassen"] = Kabel
if ebeneumbenennen.leer.IsChecked:
    kategorien_dict["Leerrohr"] = Leer
if ebeneumbenennen.auslass.IsChecked:
    kategorien_dict["Luftdurchlässe"] = Terminal
if ebeneumbenennen.ductaccessory.IsChecked:
    kategorien_dict["Luftkanalzubehör"] = Kanalzu
if ebeneumbenennen.pipeaccessory.IsChecked:
    kategorien_dict["Rohrzubehör"] = Rohrzu

if len(kategorien_dict.keys()) > 0:
    for elem in kategorien_dict.keys():
        abfrage = 'Geschosszugehörigkeit in {} schreiben?'.format(elem)
        if forms.alert(abfrage, ok=False, yes=True, no=True):
            t = DB.Transaction(doc,'Trassenzugehörigkeit -- '+ elem)
            t.Start()
            elemids = kategorien_dict[elem]
            with forms.ProgressBar(title='{value}/{max_value} '+elem, cancellable=True, step=int(len(elemids)/1000)+10) as pb:
                for n,elemid in enumerate(elemids):
                    if pb.cancelled:
                        t.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(elemids))
                    if elem in ['HLS-Bauteile','Rohrzubehör','Luftkanalzubehör']:
                        elemclass = Bauteile(elemid)
                    elif elem == 'Luftdurchlässe':
                        elemclass = Auslass(elemid)
                    else:
                        elemclass = Leitung(elemid)
                    elemclass.wert_schreiben()
            t.Commit()
            t.Dispose()