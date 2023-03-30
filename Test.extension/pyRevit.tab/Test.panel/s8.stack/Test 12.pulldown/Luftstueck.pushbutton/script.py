# coding: utf8
# from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import ExternalEventListe,ExternalEvent,DB
import os
from IGF_Forms import WPFWindow
from System.Text.RegularExpressions import Regex

# from System.Windows.Input import Key


__title__ = "Kanalerstellen"
__doc__ = """

[2022.08.01]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

# try:
#     getlog(__title__)
# except:
#     pass

# try:
#     getloglocal(__title__)
# except:
#     pass

_dict = {}
coll = DB.FilteredElementCollector(__revit__.ActiveUIDocument.Document).OfClass(DB.Mechanical.DuctType ).ToElements()
for elem in coll:
    _dict[elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()] = elem.Id

class AktuelleBerechnung(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.externaleventliste = ExternalEventListe()
        self.externaleventliste.GUI = self
        self.externalevent = ExternalEvent.Create(self.externaleventliste)
        self.DUCT_Dict = _dict
        self.ducttype.ItemsSource = sorted(self.DUCT_Dict.keys())
        self.regex = Regex("[^0-9]+")
        
    def starten(self, sender, args):
        self.externalevent.Raise()   
 
    def Textinput(self, sender, args):
        try:args.Handled = self.regex.IsMatch(args.Text)
        except:args.Handled = True
       
wind = AktuelleBerechnung()
wind.Show()
