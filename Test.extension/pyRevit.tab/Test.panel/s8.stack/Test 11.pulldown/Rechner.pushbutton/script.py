# coding: utf8
from IGF_log import getlog
from pyrevit.forms import WPFWindow
import math
from System.Text.RegularExpressions import Regex


__title__ = "2.03 Hiezkörperexponent Rechner"
__doc__ = """Heizleistung von MEP-Räume in Heizkörper schreiben

Raumheizlast für HK: LIN_BA_CALCULATED_HEATING_LOAD
Raumtemperatur: LIN_BA_DESIGN_HEATING_TEMPERATURE
Vor-,Ruecklauftemperatur: Temperatur von Medium

[2022.11.10]

"""

__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

class MEP_Uebersicht(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, 'GUI_Window.xaml')
        self.mittletemp0 = 0
        self.mittletemp1 = 0
        self.regex1 = Regex("[^0-9.]+")
    
    def textinput(self, sender, args):
        try:
            if sender.Text in ['',None]:
                args.Handled = self.regex1.IsMatch(args.Text)
            elif sender.Text.find('.') != -1 and args.Text == '.':
                args.Handled = True
            else:
                args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True
            
    def mittleretemprechnen(self,Vt,Rt,Raum):
        if Vt.Text:vt = float(Vt.Text)
        else:vt = 0.0
        if Rt.Text:rt = float(Rt.Text)
        else:rt = 0.0
        if Raum.Text:tr = float(Raum.Text)
        else:tr = 0.0
        return (vt+ rt)/2-tr
    
    def Exponentberechnen(self):
        try:
            l0 = float(self.l0.Text)
            l1 = float(self.l1.Text)
            self.exporent.Text = str(round(math.log(l1/l0)/math.log(self.mittletemp1/self.mittletemp0),4))
            return
        except:
            self.exporent.Text = '0'
            return
        
    def l1changed(self,sender,e):
        try:self.Exponentberechnen()
        except:pass

    def l0changed(self,sender,e):
        try:self.Exponentberechnen()
        except:pass
    
    def Tv1changed(self,sender,e):
        try:self.mittletemp1 = self.mittleretemprechnen(self.Tv1,self.Tr1,self.Trr1)
        except:self.mittletemp1 = 0
        try:self.Exponentberechnen()
        except:pass
    def Tr1changed(self,sender,e):
        try:self.mittletemp1 = self.mittleretemprechnen(self.Tv1,self.Tr1,self.Trr1)
        except:self.mittletemp1 = 0
        try:self.Exponentberechnen()
        except:pass
    def Trr1changed(self,sender,e):
        try:self.mittletemp1 = self.mittleretemprechnen(self.Tv1,self.Tr1,self.Trr1)
        except:self.mittletemp1 = 0
        try:self.Exponentberechnen()
        except:pass
    def Tv0changed(self,sender,e):
        try:self.mittletemp0 = self.mittleretemprechnen(self.Tv0,self.Tr0,self.Trr0)
        except:self.mittletemp0 = 0
        try:self.Exponentberechnen()
        except:pass
    def Tr0changed(self,sender,e):
        try:self.mittletemp0 = self.mittleretemprechnen(self.Tv0,self.Tr0,self.Trr0)
        except:self.mittletemp0 = 0
        try:self.Exponentberechnen()
        except:pass
    def Trr0changed(self,sender,e):
        try:self.mittletemp0 = self.mittleretemprechnen(self.Tv0,self.Tr0,self.Trr0)
        except:self.mittletemp0 = 0
        try:self.Exponentberechnen()
        except:pass

        
        
    


GUI = MEP_Uebersicht()  
GUI.ShowDialog()