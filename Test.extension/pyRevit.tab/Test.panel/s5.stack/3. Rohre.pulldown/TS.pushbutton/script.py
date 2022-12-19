# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
import os
from eventhandler import SELECT,config,DIMENSIONIEREN,UEBERNEHMEN,ExternalEvent,Key,_params,NUMMERIEREN
from System.Windows.Forms import OpenFileDialog

__title__ = "3. Dimension übernehmen_TGA"
__doc__ = """

Parameter: 
IGF_X_SM_Durchmesser
IGF_X_RVT_TS_Nr

[2022.08.29]
Version: 2.0
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

logger = script.get_logger()

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.minvalue = 0
        self.maxvalue = 100
        self.value = 1
        self.PB_text = ''
        self._params = _params
        self.select = SELECT()
        self.dimensionieren = DIMENSIONIEREN()
        self.uebernehmen = UEBERNEHMEN()
        self.nummerieren = NUMMERIEREN()
        self.OpenFileDialog = OpenFileDialog
        self.Key = Key
        # self.DialogResult = DialogResult.OK

        self.selectevent = ExternalEvent.Create(self.select)
        self.dimensionierenevent = ExternalEvent.Create(self.dimensionieren)
        self.uebernehmenevent = ExternalEvent.Create(self.uebernehmen)
        self.nummerierenevent = ExternalEvent.Create(self.nummerieren)

        self.config = config
        
        self.allebauteile = None
        self.script = script
        
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'IGF.png'))
        self.btid.ItemsSource = sorted(self._params)
        self.read_config()

        

    def read_config(self):
        try:
            param = self.config.param
            if param not in self._params:
                self.config.param = ''
            else:
                self.btid.SelectedItem = param
        except:
            pass
            
        try:
            excel = self.config.excel
            self.excel.Text = excel
        except:
            pass
    
    def write_config(self):
        try:
            self.config.param = self.btid.SelectedItem.ToString()
        except:
            pass
        try:
            self.config.excel = self.excel.Text
        except:
            pass
        self.script.save_config()
    
    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        else:
            args.Handled = True

    def durchsuchen(self, sender, args):
        dialog = self.OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "Excel"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog().ToString() == 'OK':
            self.excel.Text = dialog.FileName
        else:
            dialog.FileName = self.excel.Text

    def auswaehlen(self,sender,args):
        self.selectevent.Raise()
    def schreiben(self,sender,args):
        self.uebernehmenevent.Raise()
    def nummer(self,sender,args):
        self.nummerierenevent.Raise()
    def changedimension(self,sender,args):
        self.dimensionierenevent.Raise()

       
einstellung = AktuelleBerechnung()
einstellung.select.GUI = einstellung
einstellung.dimensionieren.GUI = einstellung
einstellung.uebernehmen.GUI = einstellung
einstellung.nummerieren.GUI = einstellung
einstellung.Show()