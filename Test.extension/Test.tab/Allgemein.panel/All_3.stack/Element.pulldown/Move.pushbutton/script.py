# coding: utf8
from pyrevit import forms
from eventhandler import LINKS,RECHTS,OBEN,UNTEN,ExternalEvent,VORNE,HINTER

__title__ = "Verschieben"
__doc__ = """
Elements verschieben

[2022.09.26]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

class Suche(forms.WPFWindow):
    def __init__(self):
        self.l = LINKS()
        self.r = RECHTS()
        self.o = OBEN()
        self.u = UNTEN()
        self.v = VORNE()
        self.h = HINTER()
        self.levent = ExternalEvent.Create(self.l)
        self.revent = ExternalEvent.Create(self.r)
        self.oevent = ExternalEvent.Create(self.o)
        self.uevent = ExternalEvent.Create(self.u)
        self.vevent = ExternalEvent.Create(self.v)
        self.hevent = ExternalEvent.Create(self.h)
        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)


    def link(self, sender, args):
        self.levent.Raise()

    def recht(self, sender, args):
        self.revent.Raise()
    
    def hinter(self, sender, args):
        self.hevent.Raise()

    def vorne(self, sender, args):
        self.vevent.Raise()
    
    def oben(self, sender, args):
        self.oevent.Raise()

    def unten(self, sender, args):
        self.uevent.Raise()


suche = Suche()
suche.h.GUI = suche
suche.v.GUI = suche
suche.l.GUI = suche
suche.r.GUI = suche
suche.o.GUI = suche
suche.u.GUI = suche
suche.Show()