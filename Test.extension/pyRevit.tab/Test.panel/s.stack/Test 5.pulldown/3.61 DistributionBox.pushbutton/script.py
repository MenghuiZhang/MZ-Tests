# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit
from pyrevit import script, forms
from eventhandler import Spiegeln,ExternalEvent


__title__ = "3.61 DistributionBox Spiegeln"
__doc__ = """

DistributionBox Spiegeln
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

uidoc = revit.uidoc
doc = revit.doc

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.mirrorelement = Spiegeln()
        self.mirrorelementEvent = ExternalEvent.Create(self.mirrorelement)             
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)

    def changetomodell(self, sender, args):
        self.mirrorelement.modell = True
        self.mirrorelement.ansicht = False
        self.mirrorelement.auswahl = False

    def changetoansicht(self, sender, args):
        self.mirrorelement.modell = False
        self.mirrorelement.ansicht = True
        self.mirrorelement.auswahl = False

    def changetoselect(self, sender, args):
        self.mirrorelement.modell = False
        self.mirrorelement.ansicht = False
        self.mirrorelement.auswahl = True  
    
    def start(self, sender, args):
        self.mirrorelementEvent.Raise()   
       
wind = AktuelleBerechnung()
wind.Show()
