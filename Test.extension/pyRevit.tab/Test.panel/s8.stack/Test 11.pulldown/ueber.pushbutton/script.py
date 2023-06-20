# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import TransitionFitting,ExternalEvent,TransitionFitting1,ExternaleventListe
import os
from System.Windows.Input import Key
from IGF_Forms import WPFWindow

__title__ = "Übergang"
__doc__ = """

[2022.08.01]
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

class AktuelleBerechnung(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.externaleventliste = ExternaleventListe()
        self.externalevent = ExternalEvent.Create(self.externaleventliste)
        self.externaleventliste.GUI = self
        self.transitionfitting = TransitionFitting()
        self.transitionfittingEvent = ExternalEvent.Create(self.transitionfitting)
        self.Key = Key
        self.transitionfitting1 = TransitionFitting1()
        self.transitionfittingEvent1 = ExternalEvent.Create(self.transitionfitting1)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    def erstellen(self, sender, args):
        self.externaleventliste.ExecuteApp = self.externaleventliste.Formteilerstellen
        self.externalevent.Raise()
        # self.transitionfittingEvent.Raise()   
    
    def aktu(self, sender, args):
        self.externaleventliste.ExecuteApp = self.externaleventliste.Bogenversprungerstellen
        self.externalevent.Raise()
        # self.transitionfittingEvent1.Raise()   


    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        else:
            args.Handled = True

       
wind = AktuelleBerechnung()
wind.transitionfitting.GUI = wind
wind.transitionfitting1.GUI = wind

wind.Show()
