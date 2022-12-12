# coding: utf8
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
from IGF_lib import get_value
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Text.RegularExpressions import Regex

__title__ = "2.20 Kühllast vs Kühlleistung"
__doc__ = """
Kühlleistung summieren
Kühllast und instalierte Kühlleistung abgleichen

Parameter:
IGF_RLT_ZuluftminRaum: Zuluftmengen
IGF_RLT_ZuluftTemperatur: Zulufttemperatur
LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE: Zulufttemperatur falls IGF_RLT_ZuluftTemperatur nicht eingegeben wird
LIN_BA_DESIGN_COOLING_TEMPERATURE: Raumtemperatur
IGF_K_KühllastLaborRaum: Kühllast Labor Raum
IGF_S_KühllastLaborPWK: Kühllast für Laboreinrichtung über PKW
LIN_BA_CALCULATED_COOLING_LOAD: Kühllast Gebäude
IGF_K_DeS_Leistung: Kühlleistung DeS
IGF_K_ULK_Leistung: Kühlleistung ULK
IGF_K_KA_Leistung: sonstige Kühlleistung
IGF_RLT_ZuluftKühlleistung: Kühlleistung Zuluft, Zuluftfaktor * (Vol_zu * 1000 * 1.2 * 1.006 * (Temp_Raum - Temp_Zu) / 3600)
IGF_K_KühlleistungRaum: Summe von Zuluft- & DeS- & ULK- & Kältekühlleistung
IGF_K_KühllastGesamt: Summe von Kühllast Gebäude und Kühllast Labor Raum
IGF_K_KühlleistungBilanz: gesamte Kühlleistung - gesamte Kühllast
IGF_K_KühlBilanzProzent: gesamte Kühlleistung / gesamte Kühllast


[Version: 2.0]
[2022.12.09]
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'Kuehlen-Familie')

try:getlog(__title__)
except:pass

bauteile = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElementIds()

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

class MechanicalEquipment(TemplateItemBase):
    def __init__(self,name,elems,params):
        TemplateItemBase.__init__(self)
        self.Name = name
        self._Selectedindex = -1
        self._Selectedart = -1
        self.Paras = params
        self.liste_art = ['Segel','Umluftkühler','Sonstige']
        self.Arts = sorted(self.liste_art)
        self._checked = False
        self.elems = elems
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    
    @property
    def Selectedindex(self):
        return self._Selectedindex
    @Selectedindex.setter
    def Selectedindex(self,value):
        if value != self._Selectedindex:
            self._Selectedindex = value
            self.RaisePropertyChanged('Selectedindex')
    
    @property
    def Selectedart(self):
        return self._Selectedart
    @Selectedart.setter
    def Selectedart(self,value):
        if value != self._Selectedart:
            self._Selectedart = value
            self.RaisePropertyChanged('Selectedart')


HLS_Dict = {}
HLS_Para_Dict = {}
for elid in bauteile:
    elem = doc.GetElement(elid)
    Family = elem.Symbol.FamilyName + ': ' + elem.Name
    if not Family in HLS_Dict.keys():
        HLS_Dict[Family] = [elem]
        Paraliste = []
        for para in elem.Parameters:
            if not para.Definition.Name in Paraliste:
                try:Paraliste.append(para.Definition.Name)
                except:pass
        HLS_Para_Dict[Family]=Paraliste
    else:
        HLS_Dict[Family].append(elem)

Liste_Datagrid = ObservableCollection[MechanicalEquipment]()

for name in sorted(HLS_Dict.keys()):Liste_Datagrid.Add(MechanicalEquipment(name,HLS_Dict[name],sorted(HLS_Para_Dict[name])))

class FamilieEinstellen(forms.WPFWindow):
    def __init__(self):
        self.Liste_Datagrid = Liste_Datagrid
        forms.WPFWindow.__init__(self, "window.xaml",handle_esc=False)
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.dataGrid.ItemsSource = self.Liste_Datagrid
        self.read_config()
        self.start = False
        self.zuluftleistungfaktor = 0.8
        self.regex1 = Regex("[^0-9,]+")
        
    def read_config(self):
        def read_config_intern(_dict,_Liste):
            for item in _Liste:
                if item.Name in _dict.keys():
                    try:item.checked = _dict[item.Name][0]
                    except:item.checked = False
                    try:item.Selectedart = item.Arts.index(_dict[item.Name][1])
                    except:item.Selectedart = -1
                    try:item.Selectedindex = item.Paras.index(_dict[item.Name][2])
                    except:item.Selectedindex = -1
                   
        try:read_config_intern(config.HeizFamilien,self.Liste_Datagrid)
        except:pass
        try:
            if config.faktor:
                try:self.faktor.Text = config.faktor
                except:pass
        except:pass

 
    def write_config(self):
        def write_conifg_intern(_Liste):
            _dict = {}
            for item in _Liste:
                if item.Selectedindex == -1:param = ''
                else:param = item.Paras[item.Selectedindex]
                if item.Selectedart == -1:art = ''
                else:art = item.Arts[item.Selectedart]
                _dict[item.Name] = [item.checked,art,param]
            return _dict
        try:config.HeizFamilien = write_conifg_intern(self.Liste_Datagrid)
        except:pass
        try:
            config.faktor = self.faktor.Text
        except:pass
        
        script.save_config()

    def _checked_changed(self,sender,Datagird):
        item = sender.DataContext
        checked = sender.IsChecked
        if Datagird.SelectedIndex != -1:
            if item in Datagird.SelectedItems:
                for el in Datagird.SelectedItems:el.checked = checked

    def _Param_changed(self,sender,Datagird):
        item = sender.DataContext
        paramindex = sender.SelectedIndex
        if paramindex != -1:
            param = item.Paras[paramindex]
            if Datagird.SelectedIndex != -1:
                if item in Datagird.SelectedItems:
                    for el in Datagird.SelectedItems:
                        if param in el.Paras:
                            el.Selectedindex = el.Paras.index(param)
    
    def _suche_textchanged(self,sender,Datagird,temp,default):
        text = sender.Text
        if not text:
            Datagird.ItemsSource = default
            return
        temp.Clear()
        for item in default:
            if item.Name.upper().find(text.upper()) != -1:
                temp.Add(item)
        Datagird.ItemsSource = temp 

    def suche_changed(self,sender,e):
        self._suche_textchanged(sender,self.dataGrid,self.tempcoll,self.Liste_Datagrid)
    
    def checkedchanged(self,sender,e):
        self._checked_changed(sender,self.dataGrid)

    def param_select_changed(self,sender,e):
        self._Param_changed(sender,self.dataGrid)
    
    def art_select_changed(self,sender,e):
        item = sender.DataContext
        paramindex = sender.SelectedIndex
        if paramindex != -1:
            param = item.Arts[paramindex]
            if self.dataGrid.SelectedIndex != -1:
                if item in self.dataGrid.SelectedItems:
                    for el in self.dataGrid.SelectedItems:
                        if param in el.Arts:
                            el.Selectedart = el.Arts.index(param)

    def pruefen(self,item):
        if item.checked:
            if item.Selectedindex == -1:
                return "{}: Heizleistung-Parameter nicht definiert".format(item.Name)
            elif item.Selectedart == -1:
                return "{}: Bauteilart nicht definiert".format(item.Name)
            else:
                return None

    def ok(self,sender,args):
        for el in self.Liste_Datagrid:
            text = self.pruefen(el)
            if text:
                UI.TaskDialog.Show('Fehler',text)
                return
        try:
            self.zuluftleistungfaktor = float(self.faktor.Text.replace(',','.'))
        except:
            UI.TaskDialog.Show('Fehler','Bitte ein gültige Zuluftleistungsfaktor eingeben!')
            return
        
        
        self.write_config()
        self.start = True
        self.Close()
    
    def textinput(self, sender, args):
        try:
            if sender.Text in ['',None]:
                args.Handled = self.regex1.IsMatch(args.Text)
            elif sender.Text.find(',') != -1 and args.Text == ',':
                args.Handled = True
            else:
                args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True


    def close(self,sender,args):
        self.Close()
    
Familie_Auswahl = FamilieEinstellen()
try:
    Familie_Auswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    Familie_Auswahl.Close()
    script.exit()

if Familie_Auswahl.start == False:
    script.exit()

class Bauteil:
    def __init__(self, elem, param, art):
        self.elem = elem
        self.paramname = param
        self.art = art
        if self.Muster_Pruefen():
            return
        
        self.space = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)]
        if self.space is None:
            logger.error('Kein MEP Raum für Bauteil {} gefunden!'.format(self.elem.Id.ToString()))
            self.spaceid = None
        else:
            self.spaceid = self.space.Id.IntegerValue

        self.param = self.elem.LookupParameter(self.paramname)
        if self.param:
            self.leistung = float(self.get_value(self.param))
        else:
            self.leistung = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.paramname))

    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
    def get_value(self,param):
        return get_value(param)
    
class MEPRaum:
    def __init__(self, elemid, Liste_Bauteile):
        self.elemid = elemid
        self.zuluftleistungsfaktor = Familie_Auswahl.zuluftleistungfaktor
        self.elem = doc.GetElement(self.elemid)
        self.name = self.get_value(self.elem.LookupParameter('Name'))
        self.nummer = self.elem.Number
        self.Liste_Bauteile = Liste_Bauteile
        self.Leistung_ULK = 0
        self.Leistung_DES = 0
        self.Leistung_SON = 0
        self.Leistung_ZUL = 0
        self.Leistung_Gesamt = 0
        self.kuehllast_Gesamt = 0
        self.Bilanz = 0
        self.Prozent = 1
        
        self.kuehllast_Gebaeude = self.get_value(self.elem.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD'))
        self.Vol_zu = self.get_value(self.elem.LookupParameter('IGF_RLT_ZuluftminRaum'))
        self.T_raum = self.get_value(self.elem.LookupParameter('LIN_BA_DESIGN_COOLING_TEMPERATURE'))
        self.Kuehllast_Labor_Raum = self.get_value(self.elem.LookupParameter('IGF_K_KühllastLaborRaum'))
        self.Kuehllast_Labor_PKW = self.get_value(self.elem.LookupParameter('IGF_S_KühllastLaborPWK'))
        self.raumtyp = self.elem.LookupParameter('Bedingungstyp').AsValueString()

        try:
            self.T_zu = self.get_value(self.elem.LookupParameter('IGF_RLT_ZuluftTemperatur'))
            if self.T_zu == '0' or self.T_zu == None:
                try:
                    self.T_zu = self.get_value(self.elem.LookupParameter('LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE'))
                except:
                    self.T_zu = -273.15
                    logger.error('kein Zulufttemperatur eingegeben in MEP-Raum {}-{}'.format(self.nummer,self.name))
            else:
                pass
        except:
            try:
                self.T_zu = self.get_value(self.elem.LookupParameter('LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE'))
            except:
                self.T_zu = -273.15
                logger.error('kein Zulufttemperatur eingegeben in MEP-Raum {}-{}'.format(self.nummer,self.name))

        self.Leistung_berechnen()

    def get_value(self,param):
        return get_value(param)

    def Leistung_berechnen(self):
        if len(self.Liste_Bauteile) > 0:
            for item in self.Liste_Bauteile:
                if item.art == 'Segel':
                    self.Leistung_DES += item.leistung
                elif item.art == 'Umluftkühler':
                    self.Leistung_ULK += item.leistung
                else:
                    self.Leistung_SON += item.leistung
        if self.raumtyp in ['Gekühlt','Beheizt und gekühlt']:
            if self.Vol_zu and self.T_zu > -273.15 and self.T_raum > -273.15:
                self.Leistung_ZUL = round(self.zuluftleistungsfaktor * (self.Vol_zu * 1000 * 1.2 * 1.006 * (self.T_raum - self.T_zu) / 3600),2)

        self.Leistung_Gesamt = self.Leistung_DES + self.Leistung_ZUL + self.Leistung_ULK + self.Leistung_SON
        self.kuehllast_Gesamt = self.kuehllast_Gebaeude + self.Kuehllast_Labor_Raum
        self.Bilanz = self.Leistung_Gesamt - float(self.kuehllast_Gesamt)
        if not self.kuehllast_Gesamt:
            self.Prozent = 1
        else:
            self.Prozent = float(self.Leistung_Gesamt)/self.kuehllast_Gesamt
    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.elem.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)
        
        wert_schreiben("IGF_K_KühllastGesamt", self.kuehllast_Gesamt)
        wert_schreiben("IGF_K_KühlleistungBilanz", self.Bilanz)
        wert_schreiben("IGF_K_KühlleistungRaum", self.Leistung_Gesamt)
        wert_schreiben("IGF_K_KühlBilanzProzent", self.Prozent)

        wert_schreiben("IGF_K_DeS_Leistung", self.Leistung_DES)
        wert_schreiben("IGF_RLT_ZuluftKühlleistung", self.Leistung_ZUL)
        wert_schreiben("IGF_K_KA_Leistung", self.Leistung_SON)
        wert_schreiben("IGF_K_ULK_Leistung", self.Leistung_ULK)

Liste_Familie = []
for el in Liste_Datagrid:
    if el.checked:
        Liste_Familie.append(el)

if len(Liste_Familie) == 0:
    UI.TaskDialog.Show('Fehler','Keine Familie ausgewählt!')
    script.exit()

MEP_Dict = {}

with forms.ProgressBar(title='{value}/{max_value} Kälte', cancellable=True, step=10) as pb:
    for n, familie in enumerate(Liste_Familie):
        pb.title='{value}/{max_value} Exemplare von ' + familie.Name + ' --- ' + str(n+1) + ' / '+ str(len(Liste_Familie)) + 'Familien'
        for n1, elem in enumerate(familie.elems):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n1 + 1, len(familie.elems))
            bauteil = Bauteil(elem,familie.Paras[familie.Selectedindex],familie.Arts[familie.Selectedart])
            if not bauteil.Muster_Pruefen():
                if bauteil.spaceid:
                    if bauteil.spaceid not in MEP_Dict.keys():
                        MEP_Dict[bauteil.spaceid] = []
                    MEP_Dict[bauteil.spaceid].append(bauteil)

mep_liste = []

MEP_Raum_Liste = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).ToElementIds()

with forms.ProgressBar(title='{value}/{max_value} MEP-Räume',cancellable=True, step=10) as pb1:
    for n, mep in enumerate(MEP_Raum_Liste):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(MEP_Raum_Liste))
        if mep.IntegerValue in MEP_Dict.keys():
            mepraum = MEPRaum(mep,MEP_Dict[mep.IntegerValue])
        else:
            mepraum = MEPRaum(mep,[])
        mep_liste.append(mepraum)

# Werte zurückschreiben + Abfrage
if forms.alert("Berechnete Werte in MEP-Räume schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Heizlast vs Heizleistung')
        for n,mepraum in enumerate(mep_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mep_liste))
            mepraum.werte_schreiben()
        t.Commit()
