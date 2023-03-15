# coding: utf8
from rpw import revit,DB,UI
from pyrevit import forms,script
from pyIGF_config import Server_config
from pyIGF_filter import Fam_Exemplar_Filter
from pyIGF_logInfo import getlog
from pyIGF_class import tempsystemtypclass,tempclass_id
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights
from lib_systemtrennung import Fam_Exemplar_Filter,Verbinder

__title__ = "Systemtrennung"
__doc__ = """..."""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
doc = revit.doc
config_user = script.get_config('Systemtrennung')
server = Server_config()
config_server = server.get_config('Systemtrennung')
config_temp = script.get_config('Systemtrennung_temp')

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
    Familie_liste = config.Verbinder
except:
    Familie_liste = []

coll = DB.FilteredElementCollector(doc).OfClass(DB.Family)
Verbindername = []

for el in coll:
    if el.FamilyCategoryId.ToString() in ['-2008010','-2008049']:
        Verbindername.append(el.Name)
coll.Dispose()

Familie_liste_neu = list(set(Familie_liste).intersection(set(Verbindername)))

if len(Familie_liste_neu) == 0:
    UI.TaskDialog.Show('Fehler','Bitte wählen Sie eine Familie aus')
    script.exit()

Richtung_id_name_dict = {0: 'In', 1: 'Out'}
Richtung_name_id_dict = {'In': 0, 'Out': 1}
Richtung__IS = ObservableCollection[tempclass_id]()

for el in Richtung_id_name_dict.keys():
    temp = tempclass_id()
    temp.Name = Richtung_id_name_dict[el]
    temp.select_id = el
    Richtung__IS.Add(temp)


class System(tempsystemtypclass):
    def __init__(self,systyp):
        tempsystemtypclass.__init__(self,systyp)
        self.checked = False
        self.richtung_is = Richtung__IS
        self.selected_id = 0
    
    @property
    def selected_id(self):
        return self._selected_id
    @selected_id.setter
    def selected_id(self,value):
        self._selected_id = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        self._checked = value

system_duct = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsElementType()
system_rohr = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsElementType()
Liste_Rohre = ObservableCollection[System]()
Liste_Luft = ObservableCollection[System]()
for el in system_duct:
    typ = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    
    Liste_Luft.Add(System(typ))
for el in system_rohr:
    typ = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    Liste_Rohre.Add(System(typ))
system_duct.Dispose()
system_rohr.Dispose()


class Systemtrennung(forms.WPFWindow):
    def __init__(self,xaml_file,liste_luft,liste_rohr):           
        self.Liste_Luft = liste_luft
        self.Liste_rohr = liste_rohr
        self.Richtung_id_name_dict = {0: 'In', 1: 'Out'}
        self.Liste_bearbeiten = []
        self.tempcoll = ObservableCollection[System]()
        forms.WPFWindow.__init__(self,xaml_file)
        self.read_config()
        if self.ebene:
            self.click(self.ansicht)
        else:
            self.click(self.projekt)
        if self.system:
            self.click(self.luft)
            self.List_view.ItemsSource = self.Liste_Luft
        else:
            self.click(self.rohr)
            self.List_view.ItemsSource = self.Liste_rohr
        self.Dict_system = {}

    
    def read_config(self):
        try:
            self.ebene = config_temp.ansicht 
        except:
            self.ebene = config_temp.ansicht = True
        try:
            self.system = config_temp.system 
        except:
            self.system = config_temp.system = True
        try:
            self.Dict_system = config_temp.liste_duct
            for el in self.Liste_Luft:
                if el.Systemtyp in config_temp.liste_duct.keys():
                    el.selected_id = config_temp.liste_duct[el.Systemtyp]
                 
                    el.checked = True
        except:
            config_temp.liste_duct = {}
        try:           
            for el in self.Liste_rohr:
                if el.Systemtyp in config_temp.liste_pipe.keys():
                    el.selected_id = config_temp.liste_pipe[el.Systemtyp]
                    self.Dict_system[el.Systemtyp] = self.Richtung_id_name_dict[el.selected_id]
                    el.checked = True
            print(self.Dict_system)
        except:
            config_temp.liste_pipe = {}
    def coll2list(self):
        liste_pipe = {}
        liste_duct = {}
        for elem in self.Liste_rohr:
            if elem.checked:
                liste_pipe[elem.Systemtyp] = elem.selected_id
        for elem in self.Liste_Luft:
            if elem.checked:
                liste_duct[elem.Systemtyp] = elem.selected_id
        return liste_pipe,liste_duct

    def write_config(self):
        try:
            config_temp.ansicht = self.ebene
        except:
            pass
        try:
            config_temp.system = self.system
        except:
            pass
        try:
            config_temp.liste_pipe,config_temp.liste_duct = self.coll2list()
        except:
            pass
        script.save_config()

    
    def click(self,button):
        button.FontWeight = FontWeights.Bold
    def back(self,button):
        button.FontWeight = FontWeights.Normal
    def back_sys(self):
        self.back(self.luft)
        self.back(self.rohr)
    def back_ebe(self):
        self.back(self.ansicht)
        self.back(self.projekt)
    def ebe_anschit(self,sender,args):
        self.back_ebe()
        self.click(self.ansicht)
        self.ebene = True
    def ebe_doc(self,sender,args):
        self.back_ebe()
        self.click(self.projekt)
        self.ebene = False
    def luftauswahl(self,sender,args):
        self.back_sys()
        self.click(self.luft)
        self.system = True
        self.List_view.ItemsSource = self.Liste_Luft
    def rohrauswahl(self,sender,args):
        self.back_sys()
        self.click(self.rohr)
        self.system = False
        self.List_view.ItemsSource = self.Liste_rohr
    def search_txt_changed(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.systemtyp.Text.upper()
        if text_typ in ['',None]:
            if self.system:
                self.List_view.ItemsSource = self.Liste_Luft
            else:
                self.List_view.ItemsSource = self.Liste_rohr
        else:
            if self.system:
                for item in self.Liste_Luft:
                    if item.Systemtyp.upper().find(text_typ) != -1:
                        self.tempcoll.Add(item)
            else:
                for item in self.Liste_rohr:
                    if item.Systemtyp.upper().find(text_typ) != -1:
                        self.tempcoll.Add(item)
            self.List_view.ItemsSource = self.tempcoll
        self.List_view.Items.Refresh()
    def check(self,sender,args):
        for item in self.List_view.Items:
            item.checked = True
        self.List_view.Items.Refresh()

    def uncheck(self,sender,args):
        for item in self.List_view.Items:
            item.checked = False
        self.List_view.Items.Refresh()

    def toggle(self,sender,args):
        for item in self.List_view.Items:
            value = item.checked
            item.checked = not value
        self.List_view.Items.Refresh()
    def checkchanged(self,sender,args):
        self.Dict_system.clear()
        Checked = sender.IsChecked
        if self.List_view.SelectedItem is not None:
            try:
                if sender.DataContext in self.List_view.SelectedItems:
                    for item in self.List_view.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.List_view.Items.Refresh()
                else:
                    pass
            except:
                pass
        for el in self.List_view.Items:
            if el.checked:
                self.Dict_system[el.Systemtyp] = self.Richtung_id_name_dict[el.selected_id]

    def daten_ermitteln(self):
        self.Dict_system.clear()
        for el in self.List_view.Items:
            if el.checked:
                self.Dict_system[el.Systemtyp] = self.Richtung_id_name_dict[el.selected_id]

    def schreiben(self,sender,args):
        self.daten_ermitteln()
        self.write_config()
        self.Liste_bearbeiten.Clear()
        for elem in Familie_liste_neu:
            Famile_coll = Fam_Exemplar_Filter(fam_name=elem,ansicht = self.ebene,doc=doc)
            for el in Famile_coll:
                typ_sys = el.LookupParameter('Systemtyp').AsValueString()
                if typ_sys in self.Dict_system.keys():
                    self.Liste_bearbeiten.append(Verbinder(el.Id,self.Dict_system[typ_sys],doc))
            Famile_coll.Dispose()
        if any(self.Liste_bearbeiten):
            t_sch = DB.Transaction(doc,'Daten schreiben')
            t_sch.Start()
            with forms.ProgressBar(title="{value}/{max_value} Verbinder ausgewählt", cancellable=True, step=1) as pb:
                for n, verbinder in enumerate(self.Liste_bearbeiten):
                    if pb.cancelled:
                        t_sch.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(self.Liste_bearbeiten))
                    try:
                        verbinder.werte_schreiben()
                    except:
                        pass
            doc.Regenerate()
            t_sch.Commit()

        self.Close()
        
    def trennen(self,sender,args):
        self.daten_ermitteln()
        self.write_config()
        self.Liste_bearbeiten.Clear()
        for elem in Familie_liste_neu:
            Famile_coll = Fam_Exemplar_Filter(fam_name=elem,ansicht = self.ebene,doc=doc)
            for el in Famile_coll:
                typ_sys = el.LookupParameter('Systemtyp').AsValueString()
                if typ_sys in self.Dict_system.keys():
                    self.Liste_bearbeiten.append(el)
            Famile_coll.Dispose()

        if any(self.Liste_bearbeiten):
            t_trenn = DB.Transaction(doc,'Trennen')
            t_trenn.Start()
            with forms.ProgressBar(title="{value}/{max_value} Verbinder ausgewählt", cancellable=True, step=1) as pb:
                for n, verbinder in enumerate(self.Liste_bearbeiten):
                    if pb.cancelled:
                        t_trenn.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(self.Liste_bearbeiten))
                    conn_l_nr = verbinder.LookupParameter('ConnectorID_Rohre').AsInteger()
                    conn_v_nr = verbinder.LookupParameter('ConnectorID_Verbinder').AsInteger()
                    elem_id = verbinder.LookupParameter('UniqueId_Rohre').AsString()
                    leitung = doc.GetElement(elem_id)
                    print(leitung)
                    try:
                        leitung = doc.GetElement(elem_id)
                        for conn in verbinder.MEPModel.ConnectorManager.Connectors: 
                            if conn.Id == conn_v_nr: conn_v = conn
                        for conn in leitung.ConnectorManager.Connectors: 
                            if conn.Id == conn_l_nr: conn_l = conn
                        conn_v.DisconnectFrom(conn_l)
                        print(conn_v,conn_l)
                    except:
                        pass
            t_trenn.Commit()
        self.Close()

    def verbinden(self,sender,args):
        self.daten_ermitteln()
        self.Liste_bearbeiten.Clear()
        for elem in Familie_liste_neu:
            Famile_coll = Fam_Exemplar_Filter(fam_name=elem,ansicht = self.ebene,doc=doc)
            for el in Famile_coll:
                typ_sys = el.LookupParameter('Systemtyp').AsValueString()
                if typ_sys in self.Dict_system.keys():
                    self.Liste_bearbeiten.append(el)
            Famile_coll.Dispose()
        if any(self.Liste_bearbeiten):
            t_verb = DB.Transaction(doc,'Verbinden')
            t_verb.Start()
            with forms.ProgressBar(title="{value}/{max_value} Verbinder ausgewählt", cancellable=True, step=1) as pb:
                for n, verbinder in enumerate(self.Liste_bearbeiten):
                    if pb.cancelled:
                        t_verb.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(self.Liste_bearbeiten))
                    conn_l_nr = verbinder.LookupParameter('ConnectorID_Rohre').AsInteger()
                    conn_v_nr = verbinder.LookupParameter('ConnectorID_Verbinder').AsInteger()
                    elem_id = verbinder.LookupParameter('UniqueId_Rohre').AsString()
           
                   
                    try:
                        leitung = doc.GetElement(elem_id)
                        for conn in verbinder.MEPModel.ConnectorManager.Connectors: 
                            if conn.Id == conn_v_nr: conn_v = conn
                        for conn in leitung.ConnectorManager.Connectors: 
                            if conn.Id == conn_l_nr: conn_l = conn
                        conn_v.ConnectTo(conn_l)
                    except:
                        pass

            t_verb.Commit()
        self.Close()
    def close(self,sender,args):
        self.Close()

wpfwindows = Systemtrennung("window.xaml",Liste_Luft,Liste_Rohre)
wpfwindows.ShowDialog()
# try:
#     wpfwindows.ShowDialog()
# except:
#     pass