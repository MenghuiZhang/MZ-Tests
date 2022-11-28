# coding: utf8
from pyrevit.forms import WPFWindow
import os
from eventhandler import ExternalEvent,ANZEIGEN
import time

__title__ = "Pick MEP"
__doc__ = """


[2021.10.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"


class Suche(WPFWindow):
    def __init__(self):
        self.time = time
        self.start = self.time.time() 
        self.anzeigen = ANZEIGEN()
        self.anzeigenevent = ExternalEvent.Create(self.anzeigen)
        WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))

    def ok(self, sender, args):
        self.anzeigenevent.Raise()          


suche = Suche()
suche.anzeigen.class_GUI = suche
suche.Show()