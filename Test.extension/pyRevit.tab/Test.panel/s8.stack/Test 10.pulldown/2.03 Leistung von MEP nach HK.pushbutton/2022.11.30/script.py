# coding: utf8
from IGF_log import getlog
from pyrevit.forms import WPFWindow
from externalevent import ExternaleventListe,Dict_MEP_Itemssource,ExternalEvent,FamilienDialog,BLACK,RED,Regex


__title__ = "2.03 Daten von MEP in HK schreiben"
__doc__ = """Heizleistung von MEP-Räume in Heizkörper schreiben

Raumheizlast für HK: LIN_BA_CALCULATED_HEATING_LOAD
Raumtemperatur: LIN_BA_DESIGN_HEATING_TEMPERATURE
Vor-,Ruecklauftemperatur: Temperatur von Medium

[2022.11.10]

"""

__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

class MEP_Uebersicht(WPFWindow):
    def __init__(self):
        self.MEP_IS = Dict_MEP_Itemssource
        self.BLACK = BLACK
        self.RED = RED
        self.regex2 = Regex("[^0-9]+")
        WPFWindow.__init__(self, 'GUI_Window.xaml')
        self.Nummer.ItemsSource = sorted(self.MEP_IS.keys())
        self.mepraum = None
        self.ExternaleventListe = ExternaleventListe()
        self.Externalevent = ExternalEvent.Create(self.ExternaleventListe)
            
    def set_up(self):
        if self.mepraum == None:
            return
        self.HKHL.Text = self.mepraum.Leistung
        self.Heizleistung.Text = self.mepraum.HKLeistung
        self.auswertung.Text = self.mepraum.auswertung
        self.auswertung.Foreground = self.mepraum.foreground
        self.lv_HK.ItemsSource = self.mepraum.HK_Bauteile
    
    def raum_sel_changed(self,sender,args):
        try:
            self.mepraum = self.MEP_IS[self.Nummer.SelectedItem]
            self.set_up()
        except Exception as e:print(e)
    
    def Heizleistungpro(self,sender,args):
        self.ExternaleventListe.ExecuteApp  =self.ExternaleventListe.HKL_Gleich_PRO
        self.Externalevent.Raise()

    def Heizleistungmep(self,sender,args):
        self.ExternaleventListe.ExecuteApp  =self.ExternaleventListe.HKL_Gleich_MEP
        self.Externalevent.Raise()

    def anderungmanuell(self,sender,args):
        self.ExternaleventListe.ExecuteApp  =self.ExternaleventListe.Aederung_MEP
        self.Externalevent.Raise()

    def datenaktuel(self,sender,args):
        self.ExternaleventListe.ExecuteApp  =self.ExternaleventListe.DatenAktualisieren
        self.Externalevent.Raise()
    
    def selectedchanegd(self,sender,e):
        if self.lv_HK.SelectedIndex != -1:
            self.ExternaleventListe.ExecuteApp  =self.ExternaleventListe.Selecteditem
            self.Externalevent.Raise()
    
    def hlchanged(self,sender,e):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK.SelectedItem is not None:
            if Item in self.lv_HK.SelectedItems:
                for item in self.lv_HK.SelectedItems:
                    item.Heizleistung = text 
                    if float(item.Heizleistung) > item.Nennleistung:
                        item.fore = self.RED
                    else:
                        item.fore = self.BLACK
        else:
            if float(Item.Heizleistung) > Item.Nennleistung:
                Item.fore = self.RED
            else:
                Item.fore = self.BLACK

    def textinput(self, sender, args):
        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True

GUI = MEP_Uebersicht()  
GUI.ExternaleventListe.GUI = GUI
if FamilienDialog.result == True:GUI.Show() 