# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import VERSCHIEBEN,ExternalEvent
from System.Windows.Input import Key
import os


__title__ = "Abstand"
__doc__ = """

Formteile und Zubehöre auf eingegebenen Abstand.

für Luftkanal- und Rohrteile.

[2022.08.01]
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

class GUI(forms.WPFWindow):
    def __init__(self):
        self.Key = Key
        self.verschieben = VERSCHIEBEN()
        self.verschiebenEvent = ExternalEvent.Create(self.verschieben)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

    def aktualisieren(self, sender, args):
        self.verschiebenEvent.Raise()   
    
    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        
        else:
            args.Handled = True
       
gui = GUI()
gui.verschieben.GUI = gui
gui.Show()