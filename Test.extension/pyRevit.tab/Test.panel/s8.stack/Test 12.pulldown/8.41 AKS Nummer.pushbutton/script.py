# coding: utf8
from Autodesk.Revit.UI import TaskDialog
from System.Collections.ObjectModel import ObservableCollection
from System.Text.RegularExpressions import Regex
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Collections.Generic import List
from clr import GetClrType
from rpw import revit,DB
from pyrevit import script,forms
from IGF_log import getlog
import os
from IGF_lib import Muster_Pruefen,get_value
from System import Guid 


__title__ = "8.30 AKS Nummer(für NA)"
__doc__ = """
AKS-Nummer ins Modell schreiben.
RaumNummer in Beiteile schreiben.
Schema:720081CB-DA99-40DC-9415-E53F280AA1F1 in Familietyp
Form: KG + '-'+ SystemABK+ '.' + AnlagenNr.+'-'+BauteilABK+'.'+Fortlaufende Nr.+'_'+Gebäudename+'.'+Ebene+'.'+Raumnr.

Parameter:
IGF_X_KG_Exemplar
IGF_X_SystemKürzel_Exemplar
IGF_X_AnlagenNr_Exemplar
IGF_X_Bauteilnummerierung
IGF_X_Einbauort


[2022.06.29]
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

def etage_nummer(etage):
    if etage.upper().find('E01') != -1:return '01'
    if etage.upper().find('E02') != -1:return '02'
    if etage.upper().find('E03') != -1:return '03'
    if etage.upper().find('E04') != -1:return '04'
    if etage.upper().find('E05') != -1:return '05'
    if etage.upper().find('E06') != -1:return '06'
    if etage.upper().find('E1') != -1:return 'E1'
    if etage.upper().find('E2') != -1:return 'E2'
    if etage.upper().find('E3') != -1:return 'E3'
    if etage.upper().find('E4') != -1:return 'E4'
    if etage.upper().find('E5') != -1:return 'E5'
    if etage.upper().find('E6') != -1:return 'E6'
    if etage.upper().find('E7') != -1:return 'E7'
    if etage.upper().find('E8') != -1:return 'E8'
    if etage.upper().find('E9') != -1:return 'E9'
    if etage.find('Bodenplatte') != -1:return '06'
    if etage.find('Dachaufsicht') != -1:return 'E9'
    return ''

class Komponent(TemplateItemBase):
    def __init__(self,name,items):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.familie = name
        self._abk = ''
        self.items = items
        self.get_abk()
    
    def get_abk(self):
        for el in self.items:
            if el.bauteilkuerzel:
                self.abk = el.bauteilkuerzel
                return
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

    @property
    def abk(self):
        return self._abk
    @abk.setter
    def abk(self,value):
        if value != self._abk:
            self._abk = value
            self.RaisePropertyChanged('abk')

class bauteil:
    def __init__(self,elem):
        self.building = ''
        self.elem = elem
        self.bsk_nr = ''
        self.kg = ''
        self.system = ''
        self.Anlagenr = ''
        self.ebene = ''
        self.raum = ''
        self.zaehler = ''
        self.bauteilkuerzel = ''
        self.raumnr = ''
        self.get_alle_info()
        self.daten_korrigieren()
    
    def get_value(self,param):
        return get_value(param)
    
    def get_alle_info(self):
        self.bsk_nr = self.get_value(self.elem.LookupParameter('IGF_X_Bauteilnummerierung'))
        self.kg = self.get_value(self.elem.LookupParameter('IGF_X_KG_Exemplar'))
        self.system = self.get_value(self.elem.LookupParameter('IGF_X_SystemKürzel_Exemplar'))
        self.Anlagenr = self.get_value(self.elem.LookupParameter('IGF_X_AnlagenNr_Exemplar'))
        self.ebene = self.get_value(self.elem.LookupParameter('Ebene'))
        self.raum = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)]
        self.zaehler = self.get_value(self.elem.LookupParameter('IGF_X_Bauteil_Zähler'))
        self.bauteilkuerzel = self.get_value(self.elem.LookupParameter('IGF_X_Bauteil_Kürzel'))
    
    def daten_korrigieren(self):
        if self.raum:
            self.raumnr = self.raum.Number
            if self.raumnr.find('-') != -1:
                try:self.raumnr = self.raumnr[self.raumnr.find('-')+1:]
                except:logger.error('Raumnummer {}, Elementid {}'.format(self.raumnr,self.elem.Id.ToString()))
        else:
            self.raumnr = 'XXX'
        
        if not self.kg:self.kg = 400
        if not self.system:self.system = 'XXX'
        if not self.Anlagenr:self.Anlagenr = 'XX'
        if self.ebene:
            try:
                self.ebene = etage_nummer(self.ebene)
            except Exception as e:
                logger.error(e)
                self.ebene = 'XX'
        if not self.ebene:self.ebene = 'XX'
        if not self.zaehler:
            if self.bsk_nr:
                try:
                    liste0 = self.bsk_nr.split('-')
                    liste1 = liste0[2].split('.')
                    liste2 = liste1[1].split('_')
                    self.zaehler = int(liste2[0])
                except:
                    self.zaehler = None
        else:
            try:
                self.zaehler = int(self.zaehler)
            except:
                self.zaehler = None
                logger.error('Zahler, Elementid {}'.format(self.elem.Id.ToString()))
    
    def wert_schreiben(self):
        aks_nr =  self.kg + '-'+self.system+'.'+self.Anlagenr+'-'+self.bauteilkuerzel+'.'+str(self.zaehler)+'_'+self.building+'.'+self.ebene+'.'+self.raumnr
        try:self.elem.LookupParameter('IGF_X_Bauteilnummerierung').Set(aks_nr)
        except:logger.error('AKS_Nr, elemid {}'.format(self.elem.Id.ToString()))
        try:
            if self.raum:
                self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raum.Number)
        except:logger.error('Raumnr, elemid {}'.format(self.elem.Id.ToString()))
        try:self.elem.LookupParameter('IGF_X_Bauteil_Kürzel').Set(str(self.bauteilkuerzel))
        except:logger.error('Bauteil_Kürzel, elemid {}'.format(self.elem.Id.ToString()))
        try:self.elem.LookupParameter('IGF_X_Bauteil_Zähler').Set(int(self.zaehler))
        except:logger.error('Bauteil_Zähler, elemid {}'.format(self.elem.Id.ToString()))

DICT_Bauteile_LUFT = {}
DICT_Bauteile_ROHR = {}
DICT_Bauteile_HLS = {}
Liste_Luft = ObservableCollection[Komponent]()
Liste_Rohr = ObservableCollection[Komponent]()
Liste_HLS = ObservableCollection[Komponent]()

def get_daten():
    luftzubehoers = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
    rohrzubehoers = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
    hlsbauteile = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
    for el in luftzubehoers:
        if Muster_Pruefen(el):continue
        fam = el.Symbol.FamilyName
        if fam in DICT_Bauteile_LUFT.keys():DICT_Bauteile_LUFT[fam].append(bauteil(el))
        else:DICT_Bauteile_LUFT[fam] = [bauteil(el)]

    for el in rohrzubehoers:
        if Muster_Pruefen(el):continue
        fam = el.Symbol.FamilyName
        if fam in DICT_Bauteile_ROHR.keys():DICT_Bauteile_ROHR[fam].append(bauteil(el))
        else:DICT_Bauteile_ROHR[fam] = [bauteil(el)]

    for el in hlsbauteile:
        if Muster_Pruefen(el):continue
        fam = el.Symbol.FamilyName
        if fam in DICT_Bauteile_HLS.keys():DICT_Bauteile_HLS[fam].append(bauteil(el))
        else:DICT_Bauteile_HLS[fam] = [bauteil(el)]

get_daten()

def get_IS():
    for el in sorted(DICT_Bauteile_LUFT.keys()):
        Liste_Luft.Add(Komponent(el,DICT_Bauteile_LUFT[el]))
    for el in sorted(DICT_Bauteile_ROHR.keys()):
        Liste_Rohr.Add(Komponent(el,DICT_Bauteile_ROHR[el]))
    for el in sorted(DICT_Bauteile_HLS.keys()):
        Liste_HLS.Add(Komponent(el,DICT_Bauteile_HLS[el]))

get_IS()

class Window_AKS(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, "aksnummer.xaml",handle_esc=False)
        self.IS_luft = Liste_Luft
        self.IS_Rohr = Liste_Rohr
        self.IS_HLS = Liste_HLS
        self.dict_daten = {}
        
        self.read_config()

        self.listview.ItemsSource = self.IS_luft
        self.altdatagrid = self.IS_luft
        self.tempcoll = ObservableCollection[Komponent]()
        
    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = sender.Text.upper()
        if text_typ in ['',None]:
            self.listview.ItemsSource = self.altdatagrid

        else:
            for item in self.altdatagrid:
                if item.familie.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.listview.ItemsSource = self.tempcoll

    def rohr(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.IS_Rohr
        self.altdatagrid = self.IS_Rohr

    def luft(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.IS_luft
        self.altdatagrid = self.IS_luft

    def hls(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.IS_HLS
        self.altdatagrid = self.IS_HLS

    def read_config(self):
        try:
            for el in self.IS_Rohr:
                if el.familie in config.AKS_Dict.keys():
                    if config.AKS_Dict[el.familie] != '': el.abk = config.AKS_Dict[el.familie] 
            for el in self.IS_luft:
                if el.familie in config.AKS_Dict.keys():
                    if config.AKS_Dict[el.familie] != '': 
                        el.abk = config.AKS_Dict[el.familie] 
            for el in self.IS_HLS:
                if el.familie in config.AKS_Dict.keys():
                    if config.AKS_Dict[el.familie] != '': el.abk = config.AKS_Dict[el.familie] 
        except Exception as e:logger.error(e)

        try:self.building.Text = config.building
        except:self.building.Text = config.building = ""
    
    def write_config(self):
        try:config.building=self.building.Text
        except:pass
        try:
            dict_test = {}
            for el in self.IS_HLS:
                dict_test[el.familie]=el.abk
            for el in self.IS_Rohr:
                dict_test[el.familie]=el.abk
            for el in self.IS_luft:
                dict_test[el.familie]=el.abk
            config.AKS_Dict = dict_test
        except Exception as e:logger.error(e)
        script.save_config()

    def daten_sammeln(self,Liste):
        for familie in Liste:
            if familie.checked:
                for elem in familie.items:
                    elem.building = self.building.Text
                    elem.bauteilkuerzel = familie.abk
                    if elem.kg not in self.dict_daten.keys():
                        self.dict_daten[elem.kg] = {}
                    if elem.system not in self.dict_daten[elem.kg].keys():
                        self.dict_daten[elem.kg][elem.system] = {}
                    if elem.Anlagenr not in self.dict_daten[elem.kg][elem.system].keys():
                        self.dict_daten[elem.kg][elem.system][elem.Anlagenr] = {}
                    if familie.abk not in self.dict_daten[elem.kg][elem.system][elem.Anlagenr].keys():
                        self.dict_daten[elem.kg][elem.system][elem.Anlagenr][familie.abk] = []
                    self.dict_daten[elem.kg][elem.system][elem.Anlagenr][familie.abk].append(elem)            

    def start(self, sender, args):
        self.write_config()
        self.Close()
        self.daten_sammeln(self.IS_Rohr)
        self.daten_sammeln(self.IS_luft)
        self.daten_sammeln(self.IS_HLS)

        t = DB.Transaction(doc,'AKS Nummer')
        t.Start()
        
        if self.ergaenzen.IsChecked:
            try:
                for kg in self.dict_daten.keys():
                    for system in self.dict_daten[kg].keys():
                        for annr in self.dict_daten[kg][system].keys():
                            for abk in self.dict_daten[kg][system][annr].keys():
                                items = self.dict_daten[kg][system][annr][abk]
                                Liste = []
                                for item in items:
                                    if item.zaehler:
                                        if item.zaehler not in Liste:
                                            Liste.append(item.zaehler)
                                        else:
                                            logger('Bauteil {} besitzt ein doppelete fortlaufende Nummer und wird erneuet zählt.'.format(item.elem.Id.ToString()))
                                            item.zaehler = None

                  
                                nummer = [n for n in range(1,len(items)+1) if n not in Liste]
                                for item in items:
                                    if item.zaehler:item.wert_schreiben()
                                    else:
                                        item.zaehler = nummer[0]
                                        nummer.remove(nummer[0])
                                        item.wert_schreiben()
            except Exception as e:logger.error(e)
        else:
            try:
                for kg in self.dict_daten.keys():
                    for system in self.dict_daten[kg].keys():
                        for annr in self.dict_daten[kg][system].keys():
                            for abk in self.dict_daten[kg][system][annr].keys():
                                items = self.dict_daten[kg][system][annr][abk]
                                for n,item in enumerate(items):
                                    item.zahler = n+1
                                    item.wert_schreiben()
            except Exception as e:logger.error(e)


        t.Commit()

    def close(self, sender, args):
        self.Close()

    def checkedchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.listview.SelectedIndex != -1:
            if item in self.listview.SelectedItems:
                for el in self.listview.SelectedItems:el.checked = checked

    def abkchanged(self, sender, args):
        item = sender.DataContext
        abk = sender.Text
        if self.listview.SelectedIndex != -1:
            if item in self.listview.SelectedItems:
                for el in self.listview.SelectedItems:el.abk = abk


window_AKS = Window_AKS()
try:
    window_AKS.ShowDialog()
except Exception as e:
    logger.error(e)
    window_AKS.Close()
    script.exit()
