# coding: utf8
from pyrevit import forms
from eventhandler import Verbinden,ExternalEvent

__title__ = "Rohre Verbinden"
__doc__ = """
Rohre Verbinden

[2022.04.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"



class Suche(forms.WPFWindow):
    def __init__(self):
        self.v = Verbinden()
        self.vevent = ExternalEvent.Create(self.v)

        forms.WPFWindow.__init__(self, 'window.xaml')


    def connect(self, sender, args):
        self.vevent.Raise()


Suche = Suche()
Suche.Show()