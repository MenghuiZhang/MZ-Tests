# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import VERBINDEN,ExternalEvent
import os


__title__ = "Hide Elements"
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
        self.transitionfitting = VERBINDEN()
        self.transitionfittingEvent = ExternalEvent.Create(self.transitionfitting)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    def erstellen(self, sender, args):
        self.transitionfittingEvent.Raise()   
       
wind = AktuelleBerechnung()
wind.transitionfitting.GUI = wind
wind.Show()
