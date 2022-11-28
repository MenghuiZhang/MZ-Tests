# coding: utf8
from pyrevit import forms
from eventhandler import XRotate,YRotate,ZRotate,ExternalEvent
import math

__title__ = "Rotate"
__doc__ = """
Elements verschieben

[2022.04.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"



class Suche(forms.WPFWindow):
    def __init__(self):
        self.x = XRotate()
        self.y = YRotate()
        self.z = ZRotate()
        self.pi = math.pi

        self.xevent = ExternalEvent.Create(self.x)
        self.yevent = ExternalEvent.Create(self.y)
        self.zevent = ExternalEvent.Create(self.z)

        forms.WPFWindow.__init__(self, 'window.xaml',handle_esc=False)


    def xrotate(self, sender, args):
        self.x.zahl = float(self.zahl.Text)/360*2*self.pi
        self.xevent.Raise()

    def yrotate(self, sender, args):

        self.y.zahl = float(self.zahl.Text)/360*2*self.pi
        self.yevent.Raise()
    
    def zrotate(self, sender, args):

        self.z.zahl = float(self.zahl.Text)/360*2*self.pi
        self.zevent.Raise()



Suche = Suche()
Suche.Show()