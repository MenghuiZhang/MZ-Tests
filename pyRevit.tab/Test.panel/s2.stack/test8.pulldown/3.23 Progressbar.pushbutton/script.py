# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
import os
from eventhandler import UEBERNEHMEN,ExternalEvent,TemplateItemBase

__title__ = "1. Progressbar"
__doc__ = """

Parameter: IGF_X_SM_Durchmesser

[2022.08.16]
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

logger = script.get_logger()

class ProgressBarItem(TemplateItemBase):
    def __init__(self):
        TemplateItemBase.__init__(self)
        self._value = 0
        self._maxvalue = 100
        self._minvalue = 0
        self._PB_text = '0/100'

    @property
    def value(self):
        return self._value
    @value.setter
    def value(self,value):
        if value != self._value:
            self._value = value
            self.RaisePropertyChanged('value')
    
    @property
    def maxvalue(self):
        return self._maxvalue
    @maxvalue.setter
    def maxvalue(self,value):
        if value != self._maxvalue:
            self._maxvalue = value
            self.RaisePropertyChanged('maxvalue')
    
    @property
    def minvalue(self):
        return self._minvalue
    @minvalue.setter
    def minvalue(self,value):
        if value != self._minvalue:
            self._minvalue = value
            self.RaisePropertyChanged('minvalue')
    
    @property
    def PB_text(self):
        return self._PB_text
    @PB_text.setter
    def PB_text(self,value):
        if value != self._PB_text:
            self._PB_text = value
            self.RaisePropertyChanged('PB_text')

class AktuelleBerechnung(forms.WPFWindow,ProgressBarItem):
    def __init__(self):
        ProgressBarItem.__init__(self)    
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.uebernehmen = UEBERNEHMEN()
        self.uebernehmenevent = ExternalEvent.Create(self.uebernehmen)
        
              
    def changeuebernehmen(self,sender,args):
        self.uebernehmenevent.Raise()

       
einstellung = AktuelleBerechnung()
einstellung.uebernehmen.GUI = einstellung
einstellung.Show()