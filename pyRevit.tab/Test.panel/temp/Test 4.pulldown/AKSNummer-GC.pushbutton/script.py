# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from System.Collections.ObjectModel import ObservableCollection
from IGF_lib import Muster_Pruefen


__title__ = "8.30 AKS Nummer(für GC)"
__doc__ = """
AKS-Nummer ins Modell schreiben

Parameter: 
input:
IGF_X_KG_Exemplar
IGF_X_SystemKürzel_Exemplar
IGF_X_AnlagenNr_Exemplar
IGF_X_Bauteilnummerierung

output:
IGF_X_Bauteilnummerierung


[2022.04.05]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+' - AKS Nummer')

try:
    getlog(__title__)
except:
    pass

def etage_nummer(etage):
    if etage.upper().find('E01') != -1:return '01'
    if etage.upper().find('E02') != -1:return '02'
    if etage.upper().find('E03') != -1:return '03'
    if etage.upper().find('E04') != -1:return '04'
    if etage.upper().find('E05') != -1:return '05'
    if etage.upper().find('E06') != -1:return '06'
    if etage.upper().find('E1') != -1:return '1'
    if etage.upper().find('E2') != -1:return '2'
    if etage.upper().find('E3') != -1:return '3'
    if etage.upper().find('E4') != -1:return '4'
    if etage.upper().find('E5') != -1:return '5'
    if etage.upper().find('E6') != -1:return '6'
    if etage.upper().find('E7') != -1:return '7'
    if etage.upper().find('E8') != -1:return '8'
    if etage.upper().find('E9') != -1:return '9'
    if etage.find('Bodenplatte') != -1:return '06'
    if etage.find('Dachaufsicht') != -1:return '9'
    return ''

class Komponent(object):
    def __init__(self,name,items):
        self.checked = False
        self.familie = name
        self.abk = ''
        self.items = items

class bauteil(object):
    def __init__(self,elem,etage,raum,altnum):
        self.elem = elem
        self.etage = etage
        self.raum = raum
        self.altnum = altnum

DICT_Bauteile_LUFT = {}
DICT_Bauteile_ROHR = {}
Liste_Luft = ObservableCollection[Komponent]()
Liste_Rohr = ObservableCollection[Komponent]()

def get_daten():
    luftzubehoers = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
    rohrzubehoers = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
    for el in luftzubehoers:
        if Muster_Pruefen(el):continue
        fam = el.Symbol.FamilyName
        if fam in DICT_Bauteile_LUFT.keys():DICT_Bauteile_LUFT[fam].append(el)
        else:DICT_Bauteile_LUFT[fam] = [el]
    for el in rohrzubehoers:
        if Muster_Pruefen(el):continue
        fam = el.Symbol.FamilyName
        if fam in DICT_Bauteile_ROHR.keys():DICT_Bauteile_ROHR[fam].append(el)
        else:DICT_Bauteile_ROHR[fam] = [el]

get_daten()

def get_IS():
    for el in sorted(DICT_Bauteile_LUFT.keys()):Liste_Luft.Add(Komponent(el,DICT_Bauteile_LUFT[el]))
    for el in sorted(DICT_Bauteile_ROHR.keys()):Liste_Rohr.Add(Komponent(el,DICT_Bauteile_ROHR[el]))

get_IS()


class Window_AKS(forms.WPFWindow):
    def __init__(self,xamldatei,IS_luft,Is_Rohr):
        self.IS_luft = IS_luft
        self.Is_Rohr = Is_Rohr
        forms.WPFWindow.__init__(self, xamldatei)
        self.read_config()
        self.listview.ItemsSource = self.IS_luft
        self.altdatagrid = self.IS_luft
        self.tempcoll = ObservableCollection[Komponent]()
    
    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.listview.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.familie.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.listview.ItemsSource = self.tempcoll
        self.listview.Items.Refresh()

    def rohr(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.Is_Rohr
        self.altdatagrid = self.Is_Rohr

    def luft(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.IS_luft
        self.altdatagrid = self.IS_luft

    def read_config(self):
        try:
            for el in self.IS_luft :
                if el.familie in config.AKS_Dict.keys():
                    el.abk = config.AKS_Dict[el.familie]
            for el in self.Is_Rohr :
                if el.familie in config.AKS_Dict.keys():
                    el.abk = config.AKS_Dict[el.familie]
        except Exception as e:logger.error(e)
            
        try:self.building.Text = config.building
        except:self.building.Text = config.building = ""
    
    def write_config(self):
        try:config.building=self.building.Text
        except:pass
        try:
            dict_test = {}
            for el in self.IS_luft:
                dict_test[el.familie]=el.abk
            for el in self.Is_Rohr:
                dict_test[el.familie]=el.abk
            config.AKS_Dict = dict_test 
        except Exception as e:logger.error(e)
        script.save_config()
    
    def start(self, sender, args):
        self.write_config()
        dict_items = {}
        for el in self.Is_Rohr:
            if el.checked:
                for item in el.items:
                    kg = item.LookupParameter('IGF_X_KG_Exemplar').AsValueString()
                    if kg is None:kg = '400'
                    system = item.LookupParameter('IGF_X_SystemKürzel_Exemplar').AsString()
                    if system is None:system = 'XXX'
                    annr = item.LookupParameter('IGF_X_AnlagenNr_Exemplar').AsValueString()
                    if annr is None:annr = 'XX'
                    try:
                        etage = item.LookupParameter('Ebene').AsValueString()
                        etage = etage_nummer(etage)
                    except:etage = ''
                    try:
                        mep = item.Space[doc.GetElement(item.get_Parameter(DB.BuiltInParameter.PHASE_CREATED).AsElementId())].Number
                        liste = mep.split('.')
                        mep = liste[len(liste)-1]
                        try:
                            etage = item.Space[doc.GetElement(item.get_Parameter(DB.BuiltInParameter.PHASE_CREATED).AsElementId())].LookupParameter('Ebene').AsValueString()
                            etage = etage_nummer(etage)
                        except:etage = ''
                        
                    except:mep = ''

                    try:
                        altnum = item.LookupParameter('IGF_X_Bauteilnummerierung').AsString()
                        liste0 = altnum.split('-')
                        liste1 = liste0[2].split('.')
                        liste2 = liste1[1].split('_')
                        altnum = int(liste2[0])
                    except:altnum = ''

                    if kg not in dict_items.keys():
                        dict_items[kg] = {}
                    if system not in dict_items[kg].keys():
                        dict_items[kg][system] = {}
                    if annr not in dict_items[kg][system].keys():
                        dict_items[kg][system][annr] = {}
                    if el.abk not in dict_items[kg][system][annr].keys():
                        dict_items[kg][system][annr][el.abk] = [bauteil(item,etage,mep,altnum)]
                    else:
                        dict_items[kg][system][annr][el.abk].append(bauteil(item,etage,mep,altnum))
        
        for el in self.IS_luft:
            if el.checked:
                for item in el.items:
                    kg = item.LookupParameter('IGF_X_KG_Exemplar').AsValueString()
                    if kg is None:kg = '400'
                    system = item.LookupParameter('IGF_X_SystemKürzel_Exemplar').AsString()
                    if system is None:system = 'XXX'
                    annr = item.LookupParameter('IGF_X_AnlagenNr_Exemplar').AsValueString()
                    if annr is None:annr = 'XX'
                    try:
                        etage = item.LookupParameter('Ebene').AsValueString()
                        etage = etage_nummer(etage)
                    except:etage = ''
                    try:
                        mep = item.Space[doc.GetElement(item.get_Parameter(DB.BuiltInParameter.PHASE_CREATED).AsElementId())].Number
                        liste = mep.split('.')
                        mep = liste[len(liste)-1]
                        try:
                            etage = item.Space[doc.GetElement(item.get_Parameter(DB.BuiltInParameter.PHASE_CREATED).AsElementId())].LookupParameter('Ebene').AsValueString()
                            etage = etage_nummer(etage)
                        except:etage = ''
                        
                    except:mep = ''
                    try:
                        altnum = item.LookupParameter('IGF_X_Bauteilnummerierung').AsString()
                        liste0 = altnum.split('-')
                        liste1 = liste0[2].split('.')
                        liste2 = liste1[1].split('_')
                        altnum = int(liste2[0])
                    except:altnum = ''
                    
                    if kg not in dict_items.keys():
                        dict_items[kg] = {}
                    if system not in dict_items[kg].keys():
                        dict_items[kg][system] = {}
                    if annr not in dict_items[kg][system].keys():
                        dict_items[kg][system][annr] = {}
                    if el.abk not in dict_items[kg][system][annr].keys():
                        dict_items[kg][system][annr][el.abk] = [bauteil(item,etage,mep,altnum)]
                    else:
                        dict_items[kg][system][annr][el.abk].append(bauteil(item,etage,mep,altnum))

        
        t = DB.Transaction(doc,'AKS Nummer')
        t.Start()
        if self.ergaenzen.IsChecked:
            try:
                for kg in dict_items.keys():
                    for system in dict_items[kg].keys():
                        for annr in dict_items[kg][system].keys():
                            for abk in dict_items[kg][system][annr].keys():
                                max_num = 0
                                for item in dict_items[kg][system][annr][abk]:
                                    if item.altnum:
                                        try:
                                            if item.altnum > max_num:max_num = item.altnum
                                        except:pass
                                for item in dict_items[kg][system][annr][abk]:
                                    if not item.altnum:
                                        try:
                                            max_num += 1
                                            kz = kg + '-'+system+'.'+annr+'-'+abk+'.'+str(max_num)+'_'+self.building.Text+'.'+item.etage+'.'+item.raum
                                            item.elem.LookupParameter('IGF_X_Bauteilnummerierung').Set(kz)
                                        except Exception as e:logger.error(e)
            except:pass
        else:
            try:
                for kg in dict_items.keys():
                    for system in dict_items[kg].keys():
                        for annr in dict_items[kg][system].keys():
                            for abk in dict_items[kg][system][annr].keys():
                                max_num = 0
                                for item in dict_items[kg][system][annr][abk]:
                                    try:
                                        max_num += 1
                                        kz = kg + '-'+system+'.'+annr+'-'+abk+'.'+str(max_num)+'_'+self.building.Text+'.'+item.etage+'.'+item.raum
                                        item.elem.LookupParameter('IGF_X_Bauteilnummerierung').Set(kz)
                                    except Exception as e:logger.error(e)
            except:pass


        t.Commit()
        self.Close()
    def close(self, sender, args):
        self.Close()

    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.listview.SelectedItem is not None:
            try:
                if sender.DataContext in self.listview.SelectedItems:
                    for item in self.listview.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass                 
                    
                    self.listview.Items.Refresh()                       
                else:
                    pass
            except:
                pass

    def abkchanged(self, sender, args):
        abk = sender.Text
        if self.listview.SelectedItem is not None:
            try:
                
                for item in self.listview.SelectedItems:
                    try:
                        item.abk = abk
                    except:
                        pass                 
                self.listview.Items.Refresh()                       

            except:
                pass    

window_AKS = Window_AKS("aksnummer.xaml",Liste_Luft,Liste_Rohr)
try:
    window_AKS.ShowDialog()
except Exception as e:
    logger.error(e)
    window_AKS.Close()
    script.exit()
