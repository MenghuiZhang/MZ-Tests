# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit
from pyrevit import script, forms
from eventhandler import CHANGEFAMILY,LISTE_IS,ExternalEvent
import os


__title__ = "3.60 BSK wechseln"
__doc__ = """

[2022.04.13]
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

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc


class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.changefamily = CHANGEFAMILY()
        self.changefamilyEvent = ExternalEvent.Create(self.changefamily)
        self.liste_is = LISTE_IS                
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.alt.ItemsSource = sorted(self.liste_is.keys())
        self.neu.ItemsSource = sorted(self.liste_is.keys())
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    
    def start(self, sender, args):
        self.changefamilyEvent.Raise()   
       
wind = AktuelleBerechnung()
wind.changefamily.GUI = wind
wind.Show()
