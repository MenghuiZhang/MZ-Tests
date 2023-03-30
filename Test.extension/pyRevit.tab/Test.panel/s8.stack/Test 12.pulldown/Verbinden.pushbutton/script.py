# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import ExternalEventListe,ExternalEvent,DB
import os
from IGF_Forms import WPFWindow
from System.Text.RegularExpressions import Regex

# from System.Windows.Input import Key


__title__ = "Verbinden"
__doc__ = """

[2022.08.01]
Version: 1.1
"""
__authors__ = "Menghui Zhang"


class AktuelleBerechnung(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.externaleventliste = ExternalEventListe()
        self.externaleventliste.GUI = self
        self.externalevent = ExternalEvent.Create(self.externaleventliste)
        
    def starten(self, sender, args):
        self.externalevent.Raise()   
 

wind = AktuelleBerechnung()
wind.Show()
