# coding: utf8
from IGF_Forms import WPFWindow
from eventhandler import ANZEIGEN,ExternalEvent


__title__ = "Pick MEP"
__doc__ = """


[2021.10.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"


class Suche(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.externaleventliste = ANZEIGEN()
        self.externaleventliste.GUI = self
        self.externalevent = ExternalEvent.Create(self.externaleventliste)

    def ok(self, sender, args):
        self.externalevent.Raise()          


suche = Suche()
suche.Show()