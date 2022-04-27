# coding: utf8
from IGF_log import getlog,getloglocal
from eventhandler import revit,forms,ADDISO,REMOVEISO,EINGABE,ExternalEvent,IS_ISO,Liste_Systemtyp,Liste_Systemtyp_1
import os
from System.Windows import GridLength

__title__ = "Rohrdämmung"
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
        self.removeISO = REMOVEISO()
        self.classEingabe = EINGABE
        self.len_0 = GridLength(0.0)
        self.len_310 = GridLength(310.0)

        self.addISOEvent = ExternalEvent.Create(self.addISO)
        self.removeISOEvent = ExternalEvent.Create(self.removeISO)

        self.vorhandenart = IS_ISO[0].elem
        self.vorhandendicke_mm = ''
        self.vorhandendicke_pro = '100'
                    
        forms.WPFWindow.__init__(self,'window.xaml')
        self.systemtyp_lv.ItemsSource = Liste_Systemtyp
        self.Liste_Systemtyp = Liste_Systemtyp
        self.tempcoll = Liste_Systemtyp_1

        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    def manuellbearbeiten(self, sender, args):
        wpfEingabe = self.classEingabe()
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
    
    def changetomanuell(self,sender,arg):
        self.button_manuell.IsEnabled = True

    def changetosystem(self,sender,arg):
        self.button_manuell.IsEnabled = False
    
    def modus_systemtyp(self, sender, args):
        self.Height = 560
        self.grid.RowDefinitions[3].Height = self.len_310
    
    def modus_system(self,sender,arg):
        self.Height = 250
        self.grid.RowDefinitions[3].Height = self.len_0

    def modus_elems(self,sender,arg):
        self.Height = 250
        self.grid.RowDefinitions[3].Height = self.len_0

    def addisoclick(self, sender, args):
        self.addISO.elem = self.system_elems.IsChecked
        self.addISO.system = self.system_sel.IsChecked
        self.addISO.vorhandenbearbeiten = self.anpassen.IsChecked
        self.addISO.rohr = self.pipe.IsChecked
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

    def removeisoclick(self, sender, args):
        self.removeISO.elem = self.system_elems.IsChecked
        self.removeISO.system = self.system_sel.IsChecked
        self.removeISO.rohr = self.pipe.IsChecked
        self.removeISO.rohrformteil = self.pipefitting.IsChecked
        self.removeISO.flexrohr = self.softpipe.IsChecked
        self.removeISO.typ = [el for el in self.Liste_Systemtyp if el.checked]
        self.removeISOEvent.Raise()
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
    def serchtextchanged(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.systemtyp_lv.ItemsSource = self.Liste_Systemtyp

        else:
            for item in self.Liste_Systemtyp:
                if item.name.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.systemtyp_lv.ItemsSource = self.tempcoll
        self.systemtyp_lv.Items.Refresh()
        
wind = AktuelleBerechnung()
wind.Show()