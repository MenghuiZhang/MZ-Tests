# coding: utf8
from IGF_log import getlog,getloglocal
from eventhandler import revit,forms,ADDISO,REMOVEISO,EINGABE,ExternalEvent,IS_ISO,Liste_Systemtyp,Liste_Systemtyp_1,\
    ObservableCollection,Systemtyp,REMOVEISO_LUFT,ADDISO_LUFT,IS_ISO_LUFT,EINGABE_LUFT
import os
from System.Windows import GridLength

__title__ = "Dämmung"
__doc__ = """
Parameter:
IGF_X_Vorgabe_ISO_Dicke: Vorgabe_ISO_Dicke
IGF_X_Vorgabe_ISO_Art: Vorgabe_ISO_Art

Category: Rohr Systeme
sind Typparameter.

[2022.04.12]
Version: 1.0

"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc


class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.addISO = ADDISO()
        self.addISOluft = ADDISO_LUFT()
        self.removeISO = REMOVEISO()
        self.removeISOluft = REMOVEISO_LUFT()
        self.classEingabe = EINGABE
        self.classEingabe_luft = EINGABE_LUFT
        self.len_0 = GridLength(0.0)
        self.len_310 = GridLength(310.0)

        self.addISOEvent = ExternalEvent.Create(self.addISO)
        self.addISOEvent_Luft = ExternalEvent.Create(self.addISOluft)
        self.removeISOEvent = ExternalEvent.Create(self.removeISO)
        self.removeISOEvent_Luft = ExternalEvent.Create(self.removeISOluft)

        self.vorhandenart = IS_ISO[0].elem
        self.vorhandendicke_mm = ''
        self.vorhandendicke_pro = '100'

        self.vorhandenart_luft = IS_ISO_LUFT[0].elem
        self.vorhandendicke_luft = '30'

        self.ISO_LUFT = IS_ISO_LUFT
        self.ISO_ROHR = IS_ISO
                    
        forms.WPFWindow.__init__(self,'window.xaml')
        self.systemtyp_lv_rohr.ItemsSource = Liste_Systemtyp
        self.Liste_Systemtyp = Liste_Systemtyp
        self.tempcoll = ObservableCollection[Systemtyp]()
        self.systemtyp_lv_luft.ItemsSource = Liste_Systemtyp_1
        self.Liste_Systemtyp1 = Liste_Systemtyp_1
        self.tempcoll1 = ObservableCollection[Systemtyp]()

        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    def manuellbearbeiten_rohr(self, sender, args):
        wpfEingabe = self.classEingabe()
        wpfEingabe.isotyp.ItemsSource = self.ISO_ROHR
        wpfEingabe.iso = self.ISO_ROHR
        if self.vorhandenart:
            liste = wpfEingabe.iso
            for el in liste:
                if el.elem.ToString() == self.vorhandenart.ToString():
                    wpfEingabe.isotyp.SelectedItem = el
                    break
        if self.vorhandendicke_mm:
            wpfEingabe.mm.IsChecked = True
            wpfEingabe.pro.IsChecked = False
            wpfEingabe.dicke_mm.Text = self.vorhandendicke_mm
            wpfEingabe.dicke_pro.Text = ''
            wpfEingabe.dicke_mm.IsEnabled = True
            wpfEingabe.dicke_pro.IsEnabled = False
        elif self.vorhandendicke_pro:
            wpfEingabe.mm.IsChecked = False
            wpfEingabe.pro.IsChecked = True
            wpfEingabe.dicke_mm.Text = ''
            wpfEingabe.dicke_pro.Text = self.vorhandendicke_pro
            wpfEingabe.dicke_mm.IsEnabled = False
            wpfEingabe.dicke_pro.IsEnabled = True

        wpfEingabe.ShowDialog()
        self.vorhandenart = wpfEingabe.isotyp.SelectedItem.elem
        self.vorhandendicke_mm = wpfEingabe.dicke_mm.Text
        self.vorhandendicke_pro = wpfEingabe.dicke_pro.Text
    
    def manuellbearbeiten_luft(self, sender, args):
        wpfEingabe = self.classEingabe_luft()
        wpfEingabe.isotyp.ItemsSource = self.ISO_LUFT
        wpfEingabe.iso = self.ISO_LUFT
        if self.vorhandenart_luft:
            liste = wpfEingabe.iso
            for el in liste:
                if el.elem.ToString() == self.vorhandenart_luft.ToString():
                    wpfEingabe.isotyp.SelectedItem = el
                    break
        if self.vorhandendicke_luft:
            wpfEingabe.dicke_mm.Text = self.vorhandendicke_luft

        wpfEingabe.ShowDialog()
        self.vorhandenart_luft = wpfEingabe.isotyp.SelectedItem.elem
        self.vorhandendicke_luft = wpfEingabe.dicke_mm.Text
    
    def changetomanuell(self,sender,arg):
        self.button_manuell_rohr.IsEnabled = True

    def changetosystem_rohr(self,sender,arg):
        self.button_manuell.IsEnabled = False
    
    def modus_systemtyp_luft(self, sender, args):
        self.Height = 560
        self.systemtyp_lv_luft.Height = 250.0
        self.filterdock_luft.Height = 25.0
    
    def modus_system_luft(self,sender,arg):
        self.Height = 310
        self.systemtyp_lv_luft.Height = 0
        self.filterdock_luft.Height = 0.0

    def modus_elems_luft(self,sender,arg):
        self.Height = 310
        self.systemtyp_lv_luft.Height = 0
        self.filterdock_luft.Height = 0.0

    def changetomanuell_luft(self,sender,arg):
        self.button_manuell_luft.IsEnabled = True

    def changetosystem_luft(self,sender,arg):
        self.button_manuell_luft.IsEnabled = False
    
    def modus_systemtyp_rohr(self, sender, args):
        self.Height = 560
        self.systemtyp_lv_rohr.Height = 250.0
        self.filterdock_rohr.Height = 25.0
    
    def modus_system_rohr(self,sender,arg):
        self.Height = 310
        self.systemtyp_lv_rohr.Height = 0.0
        self.filterdock_rohr.Height = 0.0

    def modus_elems_rohr(self,sender,arg):
        self.Height = 310
        self.systemtyp_lv_rohr.Height = 0.0
        self.filterdock_rohr.Height = 0.0
        

    def addisoclick_rohr(self, sender, args):
        self.addISO.elem = self.system_elems.IsChecked
        self.addISO.system = self.system_sel.IsChecked
        self.addISO.vorhandenbearbeiten = self.anpassen.IsChecked
        self.addISO.rohr = self.pipe.IsChecked
        self.addISO.rohraccessory = self.pipeaccessory.IsChecked
        self.addISO.rohrformteil = self.pipefitting.IsChecked
        self.addISO.flexrohr = self.softpipe.IsChecked
        self.addISO.typ = [el for el in self.Liste_Systemtyp if el.checked]
        if self.manuell.IsChecked:
            self.addISO.vorhandenart = self.vorhandenart
            self.addISO.vorhandendicke_mm = self.vorhandendicke_mm
            self.addISO.vorhandendicke_pro = self.vorhandendicke_pro
        else:
            self.addISO.vorhandenart = None
            self.addISO.vorhandendicke_mm = ''
            self.addISO.vorhandendicke_pro = ''

        self.addISOEvent.Raise() 
    
    def addisoclick_luft(self, sender, args):
        self.addISOluft.elem = self.system_elems.IsChecked
        self.addISOluft.system = self.system_sel.IsChecked
        self.addISOluft.vorhandenbearbeiten = self.anpassen_luft.IsChecked
        self.addISOluft.duct = self.duct_luft.IsChecked
        self.addISOluft.ductaccessory = self.ductaccess_luft.IsChecked
        self.addISOluft.ductformteil = self.ductfitting_luft.IsChecked
        self.addISOluft.flexduct = self.softduct_luft.IsChecked
        self.addISOluft.typ = [el for el in self.Liste_Systemtyp1 if el.checked]
        if self.manuell.IsChecked:
            self.addISOluft.vorhandenart = self.vorhandenart
            self.addISOluft.vorhandendicke_mm = self.vorhandendicke_mm
        else:
            self.addISOluft.vorhandenart = None
            self.addISOluft.vorhandendicke_mm = ''

        self.addISOEvent_Luft.Raise() 

    def removeisoclick_rohr(self, sender, args):
        self.removeISO.elem = self.system_elems.IsChecked
        self.removeISO.system = self.system_sel.IsChecked
        self.removeISO.rohr = self.pipe.IsChecked
        self.removeISO.rohraccessory = self.pipeaccessory.IsChecked
        self.removeISO.rohrformteil = self.pipefitting.IsChecked
        self.removeISO.flexrohr = self.softpipe.IsChecked
        self.removeISO.typ = [el for el in self.Liste_Systemtyp if el.checked]
        self.removeISOEvent.Raise()
    
    def removeisoclick_luft(self, sender, args):
        self.removeISOluft.elem = self.system_elems_luft.IsChecked
        self.removeISOluft.system = self.system_sel_luft.IsChecked
        self.removeISOluft.duct = self.duct_luft.IsChecked
        self.removeISOluft.ductaccessory = self.ductaccess_luft.IsChecked
        self.removeISOluft.ductformteil = self.ductfitting_luft.IsChecked
        self.removeISOluft.flexduct = self.softduct_luft.IsChecked
        self.removeISOluft.typ = [el for el in self.Liste_Systemtyp1 if el.checked]
        self.removeISOEvent_Luft.Raise()

    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.systemtyp_lv.SelectedItem is not None:
            try:
                if sender.DataContext in self.systemtyp_lv.SelectedItems:
                    for item in self.systemtyp_lv.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.systemtyp_lv.Items.Refresh()
                else:
                    pass
            except:
                pass
    
    def checkedchanged_luft(self, sender, args):
        Checked = sender.IsChecked
        if self.systemtyp_lv_luft.SelectedItem is not None:
            try:
                if sender.DataContext in self.systemtyp_lv_luft.SelectedItems:
                    for item in self.systemtyp_lv_luft.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.systemtyp_lv_luft.Items.Refresh()
                else:
                    pass
            except:
                pass

    def serchtextchanged(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.systemtyp_lv_rohr.ItemsSource = self.Liste_Systemtyp

        else:
            for item in self.Liste_Systemtyp:
                if item.name.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.systemtyp_lv_rohr.ItemsSource = self.tempcoll
        self.systemtyp_lv_rohr.Items.Refresh()
    
    def serchtextchanged_luft(self, sender, args):
        self.tempcoll1.Clear()
        text_typ = self.suche_luft.Text.upper()
        if text_typ in ['',None]:
            self.systemtyp_lv.ItemsSource = self.Liste_Systemtyp

        else:
            for item in self.Liste_Systemtyp1:
                if item.name.upper().find(text_typ) != -1:
                    self.tempcoll1.Add(item)
            self.systemtyp_lv_luft.ItemsSource = self.tempcoll
        self.systemtyp_lv_luft.Items.Refresh()
        
wind = AktuelleBerechnung()
wind.Show()