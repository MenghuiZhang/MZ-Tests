# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
import os
from eventhandler import config,ExternalEvent,_params,Externalliste
from System.Text.RegularExpressions import Regex

__title__ = "3.1 Dimension übernehmen (TGA Modell)"
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
        self.regex2 = Regex("[^0-9]+")
        self.value = 1
        self.PB_text = ''
        self._params = _params
        self.externalliste = Externalliste()
        self.externallisteevent = ExternalEvent.Create(self.externalliste)
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
    
    def textinpuut(self, sender, args):   
        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True

    def durchsuchen(self, sender, args):
        self.externalliste.ExecuteApp = self.externalliste.ordnersclect
        self.externallisteevent.Raise()

    def auswaehlen(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.select
        self.externallisteevent.Raise()

    def nummer(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.nummeriren
        self.externallisteevent.Raise()

    def exportbauteilliste(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.Bauteillisteexport
        self.externallisteevent.Raise()
       
einstellung = AktuelleBerechnung()
einstellung.externalliste.GUI = einstellung
einstellung.Show()