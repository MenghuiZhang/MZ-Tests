# coding: utf8
from pyrevit import forms
from eventhandler import LINKS,RECHTS,OBEN,UNTEN,ExternalEvent,SHANG,XIA

__title__ = "管道翻弯"
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
        forms.WPFWindow.__init__(self, 'window.xaml')


    def link(self, sender, args):
        self.l.zahl = float(self.zahl.Text) * (-1)
        self.levent.Raise()

    def recht(self, sender, args):

        self.r.zahl = float(self.zahl.Text)
        self.revent.Raise()
    
    def oben(self, sender, args):

        self.o.zahl = float(self.zahl.Text)
        self.oevent.Raise()
    def unten(self, sender, args):
        self.u.zahl = float(self.zahl.Text) * (-1)
        self.uevent.Raise()
    
    def shang(self, sender, args):

        self.s.zahl = float(self.zahl.Text)
        self.sevent.Raise()

    def xia(self, sender, args):
        self.x.zahl = float(self.zahl.Text) * (-1)
        self.xevent.Raise()


Suche = Suche()
Suche.Show()