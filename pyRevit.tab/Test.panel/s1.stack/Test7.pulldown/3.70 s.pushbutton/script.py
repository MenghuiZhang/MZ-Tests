# coding: utf8
from math import sqrt,pi
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import ListeExternalEvent,ExternalEvent
from System.Windows.Input import Key

__title__ = "Kanalrechner"

__doc__ = """


[2022.10.18]
Version: 1.0
"""
__authors__ = "Menghui Zhang"


try:getlog()
except:pass
try:getloglocal()
except:pass

class GUI(forms.WPFWindow):
    def __init__(self):
        self.Key = Key
        self.sqrt = sqrt
        self.pi = pi
        self.BreiteListe = [
            75, 90, 100, 110, 125, 140, 150, 175, 200, 225, 250, 275, 300, 325, 
            350, 375, 400, 425, 450, 475, 500, 550, 600, 650, 700, 750, 800, 850, 
            900, 950, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 
            1500, 1550, 1600, 1650, 1700, 1750, 1800, 1850, 1900, 1950, 2000, 2050, 
            2100, 2150, 2200, 2250, 2300, 2350, 2400, 2450, 2500
            ]
        
        self.DurchmesserListe = [
            75, 90, 100, 110, 125, 140, 150, 160, 175, 190, 200, 210, 225, 240, 250,
            260, 275, 290, 300, 310, 315, 325, 340, 350, 360, 375, 390, 400, 425, 450, 
            475, 500, 550, 575, 600, 625, 650, 675, 700, 725, 750, 775, 800, 825, 850, 
            875, 900, 925, 950, 975, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 
            1400, 1450, 1500, 1550, 1600, 1650, 1700, 1750, 1800, 1850, 1900, 1950, 2000,
            2050, 2100, 2150, 2200, 2250, 2300, 2350, 2400, 2450, 2500
                    ]


        self.externaleventhandler = ListeExternalEvent()
        self.externalevent = ExternalEvent.Create(self.externaleventhandler)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        
        self.breite.ItemsSource = sorted(self.BreiteListe)
        self.Hoehe.ItemsSource = sorted(self.BreiteListe)
        self.Durchmesser.ItemsSource = sorted(self.DurchmesserListe)
    
    def Setkey1(self, sender, args):
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        elif args.Key == self.Key.OemComma or args.Key == self.Key.Decimal:
            if sender.Text.Length == 0:
                args.Handled = True
            else:
                zahl_new = sender.Text.replace(',', '.') + '.'
                try:
                    float(zahl_new)
                    args.Handled = False
                except:
                    args.Handled = True
        else:
            args.Handled = True
    
    def Setkey(self, sender, args):
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        
        else:
            args.Handled = True
        
    def vol_vor_changed(self, sender, args):
        text = self.vol_vor.Text
        if text in ['',None,'0']:
            text = 0
            self.quer.Text = '0'
            self.rund_gesch.Text = '0'
            self.eck_gesch.Text = '0'
            return
        else:
            tempo = self.sch_vor.Text.replace(',','.')
            if tempo in ['',None,'0']:
                tempo = 0
                self.quer.Text = '0'
                self.rund_gesch.Text = '0'
                self.eck_gesch.Text = '0'
                return
            else:
                try:
                    quer = self.quer_berechnen(int(text),float(tempo))
                except:
                    self.quer.Text = '0'
                    self.rund_gesch.Text = '0'
                    self.eck_gesch.Text = '0'
                    return
                try:
                    self.quer.Text = str(round(quer,2)).replace('.',',')
                    if self.breite.SelectedIndex == -1:b = 0
                    else:b = int(self.breite.SelectedItem)
                    if self.Hoehe.SelectedIndex == -1:h = 0
                    else:h = int(self.Hoehe.SelectedItem)
                    b,h = self.bh_berechenn(quer,b,h)
                    v = self.v_berechnen(int(text),b,h,0)
                    self.breite.SelectedItem = int(b)
                    self.Hoehe.SelectedItem = int(h)
                    self.eck_gesch.Text = str(v).replace('.',',')
                    if self.Durchmesser.SelectedIndex == -1:d = 0
                    else:d = int(self.Durchmesser.SelectedItem)
                    d = self.d_berechenn(quer,d)
                    v  =self.v_berechnen(int(text),0,0,d)
                    self.Durchmesser.SelectedItem = int(d)
                    self.rund_gesch.Text = str(v).replace('.',',')
                except:pass
    
    def quer_berechnen(self,vol,ges):
        quer = int(vol) / float(ges) / 3600
        return quer
    
    def bh_berechenn(self,quer,b = 0,h = 0):
        b = b
        h = h
        if b == 0 and h == 0:
            b = int(self.sqrt(quer*1000000))
            while (b not in self.BreiteListe):
                b += 1
            return b,b
        elif b != 0 and h == 0:
            h = int(quer*1000000/b)
            while (h not in self.BreiteListe):
                h += 1
            return b,h

        elif b == 0 and h != 0:
            b = int(quer*1000000/h)
            while (b not in self.BreiteListe):
                b += 1
            return b,h
        else:
            return b, h
        
    def d_berechenn(self,quer,d = 0):
        d = d
        if d == 0:
            d = int(self.sqrt(quer / self.pi) * 2000)
            while (d not in self.DurchmesserListe):
                d += 1
            return d
        
        else:
            return d
    
    def v_berechnen(self, vol, b = 0, h = 0 , d = 0):
        if d == 0:
            v = round(float(vol) / h / b * 1000000 / 3600,2)
            return v
        else:
            v = round(float(vol) / (self.pi * d * d / 4000000) / 3600,2)
            return v
    
    def start(self, sender, args):
        self.externalevent.Raise()
        
            
gui = GUI()
gui.externaleventhandler.GUI = gui
gui.Show()
