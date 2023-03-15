# coding: utf8
from pyrevit import forms
from eventhandler import ExternalEvent,Drehen
from System.Text.RegularExpressions import Regex

__title__ = "Rotate"
__doc__ = """
Elements verschieben

[2022.04.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"



class ElementDrehen(forms.WPFWindow):
    def __init__(self):
        self.drehen = Drehen()
        self.drehenevent = ExternalEvent.Create(self.drehen)
        self.regex1 = Regex("[^0-9,]+")
        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)


    def rotate(self, sender, args):
        self.drehenevent.Raise()
    
    def close(self,sender,e):
        self.Close()
    
    def movewindow(self,sender,e):
        self.DragMove()
    
    def textinput(self, sender, args):
        try:
            if sender.Text in ['',None]:
                args.Handled = self.regex1.IsMatch(args.Text)
            elif sender.Text.find(',') != -1 and args.Text == ',':
                args.Handled = True
            else:
                args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True

drehen = ElementDrehen()
drehen.drehen.GUI = drehen
drehen.Show()