# coding: utf8
from IGF_log import getlog
import os
from eventhandler import ExternalEvent,ExternalEvenetListe
from IGF_Forms import WPFWindow

__title__ = "0.21 Material"
__doc__ = """

[2023.03.21]
Version: 2.0
"""
__authors__ = "Menghui Zhang"

try:getlog(__title__)
except:pass

class Suche(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.externaleventliste = ExternalEvenetListe()
        self.externaleventliste.GUI = self
        self.externalevent = ExternalEvent.Create(self.externaleventliste)        
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))
        self.material = 'Blech'

    def blech(self, sender, args):
        self.material = 'Blech'
        self.externalevent.Raise()

    def pps(self, sender, args):
        self.material = 'PPs'
        self.externalevent.Raise()
    
    def v2(self, sender, args):
        self.material = 'V2'
        self.externalevent.Raise()

suche = Suche()
suche.Show()