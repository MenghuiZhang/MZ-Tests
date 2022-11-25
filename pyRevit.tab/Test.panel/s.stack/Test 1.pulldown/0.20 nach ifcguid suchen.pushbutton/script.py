# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit.forms import WPFWindow
import os
from eventhandler import Visibility,ExternalEvent,ANZEIGEN,SELECT

__title__ = "0.20 nach IFC GUID suchen"
__doc__ = """
Element nach ifcguid suchen

[2021.10.19]
Version: 1.0
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

class Suche(WPFWindow):
    def __init__(self):
        self.anzeigen = ANZEIGEN()
        self.select = SELECT()
        self.anzeigenevent = ExternalEvent.Create(self.anzeigen)
        self.selectevent = ExternalEvent.Create(self.select)
        self.Hidden = Visibility.Hidden
        self.Visible = Visibility.Visible
        WPFWindow.__init__(self, 'window.xaml')
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))

    def Anzeigen(self, sender, args):
        # ifcguid = self.ifcguid.Text
        # self.anzeigen.ifcguid = ifcguid
        self.anzeigenevent.Raise()

    def OK(self, sender, args):
        return
        ifcguid = self.ifcguid.Text
        self.select.ifcguid = ifcguid
        self.selectevent.Raise()

    def Abbrechen(self, sender, args):
        self.Close()

suche = Suche()
suche.anzeigen.Guiclass = suche
suche.Show()