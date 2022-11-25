# coding: utf8
from IGF_log import getlog
from pyrevit import UI, DB

from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection
from Heizkoerper import config,uidoc,doc,AUSWAHL_HEIZKOERPER_IS,\
    HeizkoerperInstance,HK_Familie,logger,script,MEP_Raum,Liste_Params,\
        ChangeParameterWert,PickMEP,ExternalEvent,ChangeParameterWertForAll

__title__ = "2.03 Daten von MEP in HK schreiben"
__doc__ = """Heizleistung von MEP-Räume in Heizkörper schreiben

-------------------------
HK Category: HLS-Bauteile
-------------------------
Parameter in HK: manuell definiert, Exemplar Parameter
-------------------------
Parameter in MEP-Räume:
-------------------------
LIN_BA_CALCULATED_HEATING_LOAD: Heizlast Gesamt

IGF_H_HeizlastZuluft: Heizlast Zuluft

Heizlast für HK: LIN_BA_CALCULATED_HEATING_LOAD - IGF_H_HeizlastZuluft


!!!!!!!!!!!!!!!!!!!!!!!!
gleichmäßig: HK_KL = gesamt HK_HL / Anzahl

übernehmen: erst HL eingeben, dann in HK schreiben. 

"""

__authors__ = "Menghui Zhang"

__min_revit_ver__ = 2020

__max_revit_ver__ = 2022

try:
    getlog(__title__)
except:
    pass




class Familienauswahl(WPFWindow):
    def __init__(self,HLS_IS):
        self.HLS_IS = HLS_IS
        WPFWindow.__init__(self, 'GUI_Heizkoerper.xaml')
        self.class_HK_Familie = HK_Familie
        self.lv_HK.ItemsSource = self.HLS_IS
        self.config = config
    
    def write_config(self):
        try:
            liste = []
            for el in self.HLS_IS:
                liste.append([el.FamilienName,el.nennleistungName,el.heizleistungName])
            self.config.HK_Familie = liste
        except:
            pass
        script.save_config()

    def cancel(self,sender,args):
        self.Close()
        script.exit()

    def OK(self,sender,args):
        self.write_config()
        self.Close()

    def Add(self,sender,args):
        self.HLS_IS.Add(self.class_HK_Familie())
        self.lv_HK.Items.Refresh()

    def dele(self,sender,args):
        try:
            index = []
            for el in self.lv_HK.SelectedItems:
                index.append(self.HLS_IS.IndexOf(el))
            index.sort(reverse=True)
            for el in index:
                self.HLS_IS.RemoveAt(el)
            self.lv_HK.Items.Refresh()
        except:pass
    
    def selected_fam_changed(self,sender,args):
        self.lv_HK.Items.Refresh()

    def selected_NL_changed(self,sender,args):
        Item = sender.DataContext
        Item.get_nennleistungName()
        name = Item.nennleistungName
        if self.lv_HK.SelectedItem is not None:
            try:
                if Item in self.lv_HK.SelectedItems:
                    for item in self.lv_HK.SelectedItems:
                        try:
                            item.nennleistungName = name
                            item.get_nennleistungindex()
                            item.get_nennleistungName()
                        except:
                            pass
                    self.lv_HK.Items.Refresh()
                else:
                    pass
            except:
                pass
        self.lv_HK.Items.Refresh()

    def selected_HL_changed(self,sender,args):
        Item = sender.DataContext
        Item.get_heizleistungName()
        name = Item.heizleistungName
        if self.lv_HK.SelectedItem is not None:
            try:
                if Item in self.lv_HK.SelectedItems:
                    for item in self.lv_HK.SelectedItems:
                        try:
                            item.heizleistungName = name
                            item.get_heizleistungindex()
                            item.get_heizleistungName()
                        except:
                            pass
                    self.lv_HK.Items.Refresh()
                else:
                    pass
            except:
                pass
        self.lv_HK.Items.Refresh()

    def HK_Selection_Changed(self,sender,args):
        if self.lv_HK.SelectedIndex != -1:self.delete.IsEnabled = True
        else:self.delete.IsEnabled = False

FamilienDialog = Familienauswahl(AUSWAHL_HEIZKOERPER_IS)
try:
    FamilienDialog.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienDialog.Close()
    script.exit()

temp_Liste = []
Dict_MEP = {}

def get_Daten_Liste():
    for HK in AUSWAHL_HEIZKOERPER_IS:
        for elemid in HK.elemids:
            if elemid not in temp_Liste:
                temp_Liste.append(elemid)
                hkinstance = HeizkoerperInstance(elemid.ToString(),HK.nennleistungName,HK.heizleistungName)
                if hkinstance.raumid not in Dict_MEP.keys():
                    Dict_MEP[hkinstance.raumid] = [hkinstance]
                else:Dict_MEP[hkinstance.raumid].append(hkinstance)

get_Daten_Liste()

Dict_MEP_Itemssource = {}

def get_Itemssource():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).\
        WhereElementIsNotElementType().ToElements()
    for el in spaces:
        try:
            name = el.Number + ' - ' +el.LookupParameter('Name').AsString()
            mepraum = MEP_Raum(el.Id)
            Dict_MEP_Itemssource[name] = mepraum

            if el.Id.ToString() in Dict_MEP.keys():
                mepraum.HK_Bauteile = ObservableCollection[HeizkoerperInstance](Dict_MEP[el.Id.ToString()])
        except Exception as e:logger.error(e)

get_Itemssource()

class MEP_Uebersicht(WPFWindow):
    def __init__(self):
        self.MEP_IS = Dict_MEP_Itemssource
        WPFWindow.__init__(self, 'GUI_Window.xaml')
        self.Nummer.ItemsSource = sorted(self.MEP_IS.keys())
        self.param.ItemsSource = sorted(Liste_Params)
        self.mepraum = None
        self.Param_HK = None
        self.config = config
        self.script = script
        self.changeparameter = ChangeParameterWert()
        self.changeparameterEvent = ExternalEvent.Create(self.changeparameter)
        self.changeparameterforall = ChangeParameterWertForAll()
        self.changeparameterforallEvent = ExternalEvent.Create(self.changeparameterforall)
        self.pickmep = PickMEP()
        self.pickmepEvent = ExternalEvent.Create(self.pickmep)

        try:
            self.read_config()
        except:pass
    
    def set_up(self):
        if self.mepraum == None:
            return
        if self.Param_HK != None:
            self.mepraum.Param_HK = self.Param_HK
            self.mepraum.get_HKL()
            self.HKHL.Text = self.mepraum.Leistung

        self.lv_HK.ItemsSource = self.mepraum.HK_Bauteile
    
    def param_sel_changed(self,sender,args):
        try:
            self.Param_HK = self.param.SelectedItem.ToString()
            self.set_up()
            self.write_config()
        except:pass

    def raum_sel_changed(self,sender,args):
        try:
            self.mepraum = self.MEP_IS[self.Nummer.SelectedItem.ToString()]
            self.set_up()
        except:pass
    
    def read_config(self):
        try:
            self.Param_HK = self.config.HK_Heizlast
            self.param.Text = self.Param_HK
        except:
            pass

    
    def write_config(self):
        try:
            self.config.HK_Heizlast = self.Param_HK
        except:
            pass
        self.script.save_config()

    def durchgehen(self,sender,args):
        self.changeparameterforallEvent.Raise()

    def overwrite(self,sender,args):
        self.changeparameter.raum = self.Nummer.SelectedItem.ToString()
        self.changeparameter.bauteile = self.mepraum.HK_Bauteile
        self.changeparameterEvent.Raise()

    def gleich(self,sender,args):
        if self.mepraum.HK_Bauteile.Count == 0:return
        summe = sum([x.Nennleistung for x in self.mepraum.HK_Bauteile])
        for el in self.mepraum.HK_Bauteile:
            el.Heizleistung = round(self.mepraum.LeistungInDouble / summe * el.Nennleistung,2)
        self.lv_HK.Items.Refresh()
        self.changeparameter.raum = self.Nummer.SelectedItem.ToString()
        self.changeparameter.bauteile = self.mepraum.HK_Bauteile
        self.changeparameterEvent.Raise()

    def pickmepraum(self,sender,args):
        self.pickmepEvent.Raise()
    
MEP_Uebersicht_GUI = MEP_Uebersicht()  
MEP_Uebersicht_GUI.changeparameterforall.GUI_MEPRaum = MEP_Uebersicht_GUI
MEP_Uebersicht_GUI.pickmep.GUI_MEPRaum = MEP_Uebersicht_GUI
MEP_Uebersicht_GUI.Show() 

# dict_Bauteile = {}
# for el in FamilienDialog.ListeH:
#     dict_Bauteile[el.Famname] = el.HLeistung

# class HeizL(object):
#     def __init__(self,elemid,param):
#         self.elemid = elemid
#         self.elem = doc.GetElement(self.elemid)
#         self.Param = self.elem.LookupParameter(param)
#         self.werte = get_value(self.Param)
#         self.Space = self.elem.Space[doc.Phases[0]]
#         if self.Space:
#             self.Spaceid = self.Space.Id.ToString()
#         else:
#             self.Spaceid = ''
#         self.name = 'HZK'
#         self.typ = self.elem.Name

# class MEPRaum(object):
#     def __init__(self,elemid):
#         self.elemid = elemid
#         self.elem = doc.GetElement(self.elemid)
#         try:
#             self.HL_Gesamt = get_value(self.elem.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD'))
#         except:
#             self.HL_Gesamt = 0
#         try:
#             self.HL_Zuluft = get_value(self.elem.LookupParameter('IGF_H_HeizlastZuluft'))
#         except:
#             self.HL_Zuluft = 0
#         self.HL_HK = self.HL_Gesamt - self.HL_Zuluft

# dict_Bauteile_MEP = {}
# for el in dict_Bauteile.keys():
#     coll = Element_FamilyFilter(el)
#     for elem in coll:
#         heizl = HeizL(elem.Id,dict_Bauteile[el])
#         if heizl.Spaceid:
#             if heizl.Spaceid not in dict_Bauteile_MEP.keys():
#                 dict_Bauteile_MEP[heizl.Spaceid] = ObservableCollection[HeizL]()
#                 dict_Bauteile_MEP[heizl.Spaceid].Add(heizl)
#             else:
#                 dict_Bauteile_MEP[heizl.Spaceid].Add(heizl)

# global WEITER
# WEITER = False

# class Familienbearbeiten(WPFWindow):
#     def __init__(self, xaml_file_name):
#         WPFWindow.__init__(self, xaml_file_name)
#         self.elem = PickMEP()
#         self.mepraum = ''
#         self.dataGrid.ItemsSource = ObservableCollection[HeizL]()
#         self.items = []
#         self.set_up()

#     def set_up(self):
#         self.mepraum = MEPRaum(self.elem.Id)
#         self.Nummer.Text = self.mepraum.elem.Number
#         self.GesamtHL.Text = str(round(self.mepraum.HL_Gesamt))
#         self.ZuluftHL.Text = str(round(self.mepraum.HL_Zuluft))
#         self.HKHL.Text = str(round(self.mepraum.HL_HK))
#         if self.elem.Id.ToString() in dict_Bauteile_MEP.keys():
#             self.items = dict_Bauteile_MEP[self.elem.Id.ToString()]
#             i = 0
#             for elem in self.items:
#                 i += 1
#                 elem.name = 'HZK-' + str(i)
#             self.dataGrid.ItemsSource = self.items
#         else:
#             UI.TaskDialog.Show(__title__,"Keine HZK im ausgewählten Raum gefunden! Bitte wählen Sie ein neu Raum aus.")
#             self.elem = PickMEP()
#             self.set_up()

#     def close(self,sender,args):
#         self.Close()

#     def gleich(self,sender,args):
#         t = DB.Transaction(doc,self.mepraum.elem.Number)
#         t.Start()
#         for el in self.items:
#             el.werte = round(float(self.HKHL.Text) / self.items.Count)
#             el.Param.SetValueString(str(el.werte))
#         doc.Regenerate()
#         t.Commit()
#         t.Dispose()
#         self.dataGrid.Items.Refresh()
    
#     def weiter(self,sender,args):
#         global WEITER
#         WEITER = True
#         self.Close()
#         # self.elem = PickMEP()
#         # self.set_up()
#         # self.show_dialog()

#     def overwrite(self,sender,args):
#         self.dataGrid.Items.Refresh()
#         t = DB.Transaction(doc,self.mepraum.elem.Number)
#         t.Start()
#         for el in self.items:
#             el.Param.SetValueString(str(el.werte))
#         doc.Regenerate()
#         t.Commit()
#         t.Dispose()

# Bauteilebearbeiten = Familienbearbeiten('final_bearbeiten.xaml')
# try:
#     Bauteilebearbeiten.ShowDialog()
# except Exception as e:
#     logger.error(e)
#     Bauteilebearbeiten.Close()
#     script.exit()

# while(WEITER):
#     WEITER = False
#     Bauteilebearbeiten = Familienbearbeiten('final_bearbeiten.xaml')
#     try:
#         Bauteilebearbeiten.ShowDialog()
#     except Exception as e:
#         logger.error(e)
#         Bauteilebearbeiten.Close()
#         script.exit()
