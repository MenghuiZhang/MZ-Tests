# coding: utf8
from pyrevit import forms
from eventhandler import ExternalEvent,SHANG

__title__ = "Abstand"
__doc__ = """
Elements verschieben

[2022.04.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"



class Suche(forms.WPFWindow):
    def __init__(self):

        self.s = SHANG()

        self.sevent = ExternalEvent.Create(self.s)

        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
    
    def shang(self, sender, args):
        self.s.zahl = 260 #- float(self.zahl.Text)
        self.sevent.Raise()

suche = Suche()
suche.Show()