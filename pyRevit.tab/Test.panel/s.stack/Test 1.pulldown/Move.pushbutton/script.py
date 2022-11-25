# coding: utf8
from pyrevit import forms
from eventhandler import LINKS,RECHTS,OBEN,UNTEN,ExternalEvent,SHANG,XIA

__title__ = "Verschieben"
__doc__ = """
Elements verschieben

[2022.04.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"



class Suche(forms.WPFWindow):
    def __init__(self):
        self.l = LINKS()
        self.r = RECHTS()
        self.o = OBEN()
        self.u = UNTEN()
        self.s = SHANG()
        self.x = XIA()
        self.levent = ExternalEvent.Create(self.l)
        self.revent = ExternalEvent.Create(self.r)
        self.oevent = ExternalEvent.Create(self.o)
        self.uevent = ExternalEvent.Create(self.u)
        self.sevent = ExternalEvent.Create(self.s)
        self.xevent = ExternalEvent.Create(self.x)
        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
    
    def shang(self, sender, args):
        self.s.zahl = 260 - float(self.zahl.Text)
        self.sevent.Raise()



Suche = Suche()
Suche.Show()