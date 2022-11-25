# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import TransitionFitting,ExternalEvent,TransitionFitting1
import os


__title__ = "Übergang"
__doc__ = """

[2022.04.13]
Version: 1.1
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

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.transitionfitting = TransitionFitting()
        self.transitionfitting1 = TransitionFitting1()
        self.transitionfittingEvent = ExternalEvent.Create(self.transitionfitting)
        self.transitionfittingEvent1 = ExternalEvent.Create(self.transitionfitting1)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
        self.co0 = ''
        self.co1 = ''
        self.l0 = ''
    
    def erstellen(self, sender, args):
        self.transitionfittingEvent.Raise()   
        self.transitionfittingEvent1.Raise()   

wind = AktuelleBerechnung()
wind.transitionfitting.GUI = wind
wind.transitionfitting1.GUI = wind
wind.Show()