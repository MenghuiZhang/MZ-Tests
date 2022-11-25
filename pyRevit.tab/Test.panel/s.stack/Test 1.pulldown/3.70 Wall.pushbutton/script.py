# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit
from pyrevit import script, forms
from eventhandler import CHANGEFAMILY,ExternalEvent,WALLS,LEVELS
import os


__title__ = "3.70 Raumtrennung -> Wände"
__doc__ = """

[2022.04.13]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc


class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.changefamily = CHANGEFAMILY()
        self.changefamilyEvent = ExternalEvent.Create(self.changefamily)

        self.walls = WALLS
        self.levels = LEVELS

        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.wall.ItemsSource = sorted(self.walls.keys())
        self.EU.ItemsSource = sorted(self.levels.keys())
        self.EO.ItemsSource = sorted(self.levels.keys())

        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
        
    def start(self, sender, args):
        self.changefamilyEvent.Raise()   
       
wind = AktuelleBerechnung()
wind.changefamily.class_GUI = wind
wind.Show()
