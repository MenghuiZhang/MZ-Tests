# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import ExternalEvent,BESCHIFTUNG2,Aktualisieren
import os


__title__ = "Leistung"
__doc__ = """

Beschriftung ausrichten
gilt nur in Grundriss.

1. auf X(Y) ausrichten.
2. fixierte Beschriftung auswählen
3. zu verschiebenen Beschriftungen auswählen
4. Fertig stellen klicken.

[2022.08.08]
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


class GUI(forms.WPFWindow):
    def __init__(self):
        self.Beschriftung1 = BESCHIFTUNG2()
        self.BeschriftungEvent1 = ExternalEvent.Create(self.Beschriftung1)
        self.Beschriftung2 = Aktualisieren()
        self.BeschriftungEvent2 = ExternalEvent.Create(self.Beschriftung2)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
         
    def manuell(self, sender, args):
        self.BeschriftungEvent1.Raise()   
    def aktu(self, sender, args):
        self.BeschriftungEvent2.Raise()   

gui = GUI()
gui.Beschriftung1.GUI = gui
gui.Beschriftung2.GUI = gui
gui.Show()