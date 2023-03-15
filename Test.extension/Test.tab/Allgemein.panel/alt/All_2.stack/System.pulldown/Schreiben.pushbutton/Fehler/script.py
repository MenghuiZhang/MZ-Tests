# coding: utf8
from rpw import revit,DB,UI
from pyrevit import forms,script
from pyIGF_config import Server_config
from pyIGF_filter import Fam_Exemplar_Filter
from pyIGF_logInfo import getlog
from pyIGF_class import tempsystemtypclass,tempclass_id
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights

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
        self.ebene = True
        self.system = True
        self.click(self.luft)
        self.click(self.ansicht)
        self.List_view.ItemsSource = self.Liste_Luft
        self.Dict_system = {}
    
    def click(self,button):
        from System.Windows import FontWeights
        button.FontWeight = FontWeights.Bold
    def back(self,button):
        from System.Windows import FontWeights
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

    def schreiben(self,sender,args):
        # coding: utf8
        from lib_systemtrennung import Fam_Exemplar_Filter,Verbinder
        from pyrevit import revit, UI, forms
        from Autodesk.Revit.DB import Transaction
        #doc = revit.doc
        self.Liste_bearbeiten.Clear()
        for elem in Familie_liste_neu:
            Famile_coll = Fam_Exemplar_Filter(fam_name=elem,ansicht = self.ebene,doc=self.doc)
            for el in Famile_coll:
                typ_sys = el.LookupParameter('Systemtyp').AsValueString()
                if typ_sys in self.Dict_system.keys():
                    self.Liste_bearbeiten.append(Verbinder(el.Id,self.Dict_system[typ_sys],self.doc))
            Famile_coll.Dispose()
        print(self.Liste_bearbeiten)
        if any(self.Liste_bearbeiten):
            #doc = __revit__.ActiveUIDocument.Document
            t = Transaction(self.doc,'Daten schreiben')
            t.Start()
            with forms.ProgressBar(title="{value}/{max_value} Verbinder ausgewählt", cancellable=True, step=1) as pb:
                for n, verbinder in enumerate(self.Liste_bearbeiten):
                    if pb.cancelled:
                      #  t.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(self.Liste_bearbeiten))
                    # try:
                    #     verbinder.werte_schreiben()
                    # except:
                    #     pass
            #self.doc.Regenerate()
            t.Commit()
        
    def trennen(self,sender,args):
        # coding: utf8
        from lib_systemtrennung import Fam_Exemplar_Filter,Verbinder
        from pyrevit import forms,script
        import Autodesk.Revit.DB as DB
        self.Liste_bearbeiten.Clear()
        for elem in self.list_fam:
            Famile_coll = Fam_Exemplar_Filter(fam_name=elem,ansicht = self.ebene,doc=self.doc)
            for el in Famile_coll:
                typ_sys = el.LookupParameter('Systemtyp').AsValueString()
                if typ_sys in self.Dict_system.keys():
                    self.Liste_bearbeiten.append(el)
            Famile_coll.Dispose()
        if any(self.Liste_bearbeiten):
            t = DB.Transaction(self.doc,'Trennen')
            t.Start()
            with forms.ProgressBar(title="{value}/{max_value} Verbinder ausgewählt", cancellable=True, step=1) as pb:
                for n, verbinder in enumerate(self.Liste_bearbeiten):
                    if pb.cancelled:
                        t.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(self.Liste_bearbeiten))
                    conn_l_nr = verbinder.LookupParameter('ConnectorID_Rohre').AsInteger()
                    conn_v_nr = verbinder.LookupParameter('ConnectorID_Verbinder').AsInteger()
                    elem_id = verbinder.LookupParameter('UniqueId_Rohre').AsString()
                    leitung = self.doc.GetElement(elem_id)
                    for conn in verbinder.MEPModel.ConnectorManager.Connectors: 
                        if conn.Id == conn_v_nr: conn_v = conn
                    for conn in leitung.ConnectorManager.Connectors: 
                        if conn.Id == conn_l_nr: conn_l = conn
                    conn_v.DisconnectFrom(conn_l)
                    # try:
                    #     leitung = self.doc.GetElement(elem_id)
                    #     for conn in verbinder.MEPModel.ConnectorManager.Connectors: 
                    #         if conn.Id == conn_v_nr: conn_v = conn
                    #     for conn in leitung.ConnectorManager.Connectors: 
                    #         if conn.Id == conn_l_nr: conn_l = conn
                    #     conn_v.DisconnectFrom(conn_l)
                    # except:
                    #     pass
            self.doc.Regenerate()
            t.Commit()
    def verbinden(self,sender,args):
        # coding: utf8
        from lib_systemtrennung import Fam_Exemplar_Filter,Verbinder
        from pyrevit import forms,script,DB
        self.Liste_bearbeiten.Clear()
        for elem in self.list_fam:
            Famile_coll = Fam_Exemplar_Filter(fam_name=elem,ansicht = self.ebene,doc=self.doc)
            for el in Famile_coll:
                typ_sys = el.LookupParameter('Systemtyp').AsValueString()
                if typ_sys in self.Dict_system.keys():
                    self.Liste_bearbeiten.append(el)
            Famile_coll.Dispose()
        if any(self.Liste_bearbeiten):
            t = DB.Transaction(self.doc,'Verbinden')
            t.Start()
            with forms.ProgressBar(title="{value}/{max_value} Verbinder ausgewählt", cancellable=True, step=1) as pb:
                for n, verbinder in enumerate(self.Liste_bearbeiten):
                    if pb.cancelled:
                        t.RollBack()
                        script.exit()
                    pb.update_progress(n + 1, len(self.Liste_bearbeiten))
                    conn_l_nr = verbinder.LookupParameter('ConnectorID_Rohre').AsInteger()
                    conn_v_nr = verbinder.LookupParameter('ConnectorID_Verbinder').AsInteger()
                    elem_id = verbinder.LookupParameter('UniqueId_Rohre').AsString()
                    try:
                        leitung = self.doc.GetElement(elem_id)
                        for conn in verbinder.MEPModel.ConnectorManager.Connectors: 
                            if conn.Id == conn_v_nr: conn_v = conn
                        for conn in leitung.ConnectorManager.Connectors: 
                            if conn.Id == conn_l_nr: conn_l = conn
                        conn_v.ConnectTo(conn_l)
                    except:
                        pass
            self.doc.Regenerate()
            t.Commit()
    def close(self,sender,args):
        self.Close()

wpfwindows = Systemtrennung("window.xaml",Liste_Luft,Liste_Rohre,Familie_liste_neu,doc)
wpfwindows.Show()
# try:
#     wpfwindows.Show()
# except:
#     pass

# systemname_id=DB.ElementId(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
# systemname_prov=DB.ParameterValueProvider(systemname_id)
# param_equality=DB.FilterStringEquals() # equality class
# systemname_value_rule=DB.FilterStringRule(systemname_prov,param_equality,systenname,True)
# systemname_filter=DB.ElementParameterFilter(systemname_value_rule)

# systemtyp_id=DB.ElementId(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
# systemtyp_prov=DB.ParameterValueProvider(systemtyp_id)
# systemtyp_value_rule=DB.FilterStringRule(systemtyp_prov,param_equality,System_typ,True)
# systemtyp_filter=DB.ElementParameterFilter(systemtyp_value_rule)

# duct_coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctCurves).WherePasses(systemname_filter).WherePasses(systemtyp_filter).WhereElementIsNotElementType()
# elements = duct_coll.ToElementIds()
# print(len(elements))
# duct_coll.Dispose()



# class Verbinder(object):
#     def __init__(self,elementid,fliessrichtung):
#         self.fliessrichtung = fliessrichtung 
#         self.elem_id = elementid
#         self.elem = doc.GetElement(self.elem_id)
#         self.connid_Leitung = ''
#         self.connid_verbinder = ''
#         self.elemid_Leitung = ''
#         self.Daten_ermitteln()
    
#     @property
#     def conns(self):
#         return self.elem.MEPModel.ConnertorManager.Connectors
#     @property
#     def connid_verbinder(self):
#         return self.connid_verbinder
#     @property
#     def connid_Leitung(self):
#         return self.connid_Leitung
#     @property
#     def elemid_Leitung(self):
#         return self.elemid_Leitung
#     def Daten_ermitteln(self):
#         for conn in self.conns:
#             if conn:
#                 pass
#             self.connid_verbinder = conn.Id
#             for ref in conn.Allrefs:
#                 try:
#                     if ref.Owner.Category.Id.ToString() in ['-2008000','']:
#                         self.connid_Leitung = ref.Id
#                         self.elemid_Leitung = ref.Owner.Id.ToString
#                         break
#                 except:
#                     pass
#             if self.connid_Leitung:
#                 break

#     def werte_schreiben(self):
#         def wert_schreiben(param,value):
#             try:
#                 self.elem.LookupParameter(param).Set(value)
#             except:
#                 logger.error('Fehler beim Werte-Schreiben in Parameter "{}".format(param)')
#         wert_schreiben('ConnertorId_Rohre',self.connid_Leitung)
#         wert_schreiben('ConnertorId_Verbinder',self.connid_verbinder)
#         wert_schreiben('ElementID_Rohre',self.elemid_Leitung)

# Liste_richtig = []
# Liste_Fehler = []
# with forms.ProgressBar(title="{value}/{max_value} Elemente in Projekt ausgewählt", cancellable=True, step=10) as pb:   
#     for n, elemid in enumerate(elemids):
#         if pb.cancelled:
#             script.exit()
#         pb.update_progress(n + 1, len(elemids))
#         verbinder = Verbinder(elemid)
#         if verbinder.connid_Leitung:
#             Liste_richtig.append(verbinder)
#         else:
#             Liste_Fehler.append(elemid.ToString())




# if len(Liste_richtig) == 0:
#     UI.TaskDialog.Show(__title__,'Keine Verbinder ausgewählt')
#     script.exit()
# else:
#     print('{} Verbinder ausgewählt'.format(len(Liste_richtig)))

# # Werte zuückschreiben + Abfrage
# if forms.alert('Werte schreiben?', ok=False, yes=True, no=True):
#     t = DB.Transaction(doc,'Werte schreiben')
#     t.Start()
#     with forms.ProgressBar(title="{value}/{max_value} Verbinder in Projekt ausgewählt", cancellable=True, step=10) as pb1:
#         for n1, verbinder in enumerate(Liste_richtig):
#             if pb1.cancelled:
#                 t.Rollback()
#                 script.exit()
#             pb.update_progress(n1 + 1, len(Liste_richtig))
#             verbinder.werte_schreiben()
            
#     t.Commit()

# if len(Liste_Fehler) != 0:
#     logger.error('Problem mit Verbinder: /n')
#     logger.error(Liste_Fehler)