# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
import os
from eventhandler import config,ExternalEvent,_params,Externalliste

__title__ = "3.2 Dimension übernehmen (Schema)"
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
        self.externalliste = Externalliste()
        self.externallisteevent = ExternalEvent.Create(self.externalliste)
        self.config = config
        
        self.allebauteile = None
        self.script = script
        
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'IGF.png'))
        self.btid.ItemsSource = sorted(self._params)
        self.read_config()
        self.maxwert = 0

        

    def read_config(self):
        try:
            param = self.config.param
            if param not in self._params:
                self.config.param = ''
            else:
                self.btid.SelectedItem = param
        except:
            pass
    
    def write_config(self):
        try:
            self.config.param = self.btid.SelectedItem
        except:
            pass

        self.script.save_config()


    def durchsuchen(self, sender, args):
        self.externalliste.ExecuteApp = self.externalliste.ordnersclect
        self.externallisteevent.Raise()

    def excel_textchanged(self,sender,e):
        text = sender.Text
        if text:
            try:
                if text.find('IGF_X_Information ans Schema_Rohre_') != -1:
                    zahl = text[text.find('IGF_X_Information ans Schema_Rohre_'):text.find('.xlsx')].replace('IGF_X_Information ans Schema_Rohre_','')
                    minzahl = zahl[:zahl.find('_')]
                    maxzahl = zahl[zahl.find('_')+1:]
                    self.maxwert = maxzahl
                    self.startnr.Text = minzahl
            except Exception as e:
                print(e)

    def auswaehlen(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.select
        self.externallisteevent.Raise()

    def nummer(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.nummeriren
        self.externallisteevent.Raise()

    def datenschreiben(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.Datenschreiben
        self.externallisteevent.Raise()

    def dimensionieren(self,sender,args):
        self.externalliste.ExecuteApp = self.externalliste.Dimensionieren
        self.externallisteevent.Raise()
       

       
einstellung = AktuelleBerechnung()
einstellung.externalliste.GUI = einstellung
einstellung.Show()