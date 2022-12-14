# coding: utf8
from math import sqrt,pi,ceil
from IGF_log import getlog,getloglocal
from pyrevit import forms,revit
from eventhandler import ListeExternalEvent,ExternalEvent,DB
from System.Windows.Input import Key,TraversalRequest,FocusNavigationDirection,Keyboard,KeyStates
from System.Windows import Visibility
from System.Text.RegularExpressions import Regex
from System import Action


__title__ = "Kanalrechner"

__doc__ = """


[2022.10.31]
Version: 1.3
"""
__authors__ = "Menghui Zhang"


try:getlog()
except:pass
try:getloglocal()
except:pass

class GUI(forms.WPFWindow):
    def __init__(self):
        self.Key = Key
        self.Keyboard = Keyboard
        self.KeyStates = KeyStates
        self.Visible = Visibility.Visible
        self.Hidden = Visibility.Hidden
        self.class_Action = Action
        self.RundListe = [63,71,80,90,100,112,125,140,150,160,180,200,224,250,280,300,315,355,400,450,500,560,600,630,710,800,900,1000,1120,1250,1400,1500,1600]
        self.sqrt = sqrt
        self.ceil = ceil
        self.elems = []
        self.uebertragenforcheck = False
        self.regex0 = Regex("[^0-9+-]+")
        self.regex1 = Regex("[^0-9,]+")
        self.regex2 = Regex("[^0-9]+")
        self.pi = pi
        self.TraversalRequest = TraversalRequest
        self.focusNavigationDirection = FocusNavigationDirection.Next
        self.ebenedict = {el.Name:el.Id for el in DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()}

        self.externaleventhandler = ListeExternalEvent()
        self.externalevent = ExternalEvent.Create(self.externaleventhandler)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.liste_rohr.ItemsSource = sorted(self.RundListe)
        self.refe.ItemsSource = sorted(self.ebenedict.keys())

    def Volvorgabeberechnen(self, sender, args):
        try:
            args.Handled = self.regex0.IsMatch(args.Text)
        except:
            args.Handled = True
    

    
    def Nichtnurganzzahl(self, sender, args):
        try:
            if sender.Text in ['',None]:
                args.Handled = self.regex1.IsMatch(args.Text)
            elif sender.Text.find(',') != -1 and args.Text == ',':
                args.Handled = True
            else:
                args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True

    def Nurganzzahl(self, sender, args):
        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True

    def Setkey3(self, sender, args):
        if args.Key == self.Key.Tab:
            try:self.MoveFocus(self.TraversalRequest(self.focusNavigationDirection))
            except Exception as e:print(e)
        elif args.Key == self.Key.Enter:
            try:self.Keyboard.FocusedElement.IsChecked = True
            except Exception as e:print(e)
        else:
            args.Handled = True

    def menge_berechnen(self, sender, args):
        try:
            self.menge_berechn()
            if not self.uebertragenforcheck:self.demension_berechn()
            self.uebertragenforcheck = False
        except:pass
    
    def get_vol(self):
        Text = self.eingabevol.Text
        vol = 0.0
        Liste = Text.split('+')
        for el in Liste:
            if el == '':
                pass
            elif el.find('-') != -1:
                liste = el.split('-')
                if liste[0] == '':
                    liste[0] = 0
                if liste[1] == '':
                    liste[1] = 0
                vol = vol + float(liste[0]) - float(liste[1])
            else:
                vol = vol + float(el)
        return vol
    
    def TextBox_GotFocus(self,sender,e):
        tb = sender
        tb.SelectAll()
        # tb.Dispatcher.BeginInvoke(self.Action(tb.SelectAll()))

    def menge_berechn(self):
        if self.eingabeb.Text in ['',None,'0']:b = 0
        else:b = int(self.eingabeb.Text)
        if self.eingabeh.Text in ['',None,'0']:h = 0
        else:h = int(self.eingabeh.Text)
        if self.eingabed.Text in ['',None,'0']:d = 0
        else:d = int(self.eingabed.Text)

        if self.modus_tem.IsChecked:
            if self.eingabevol.Text in ['',None,'0']:vol = 0.0
            else:
                try:vol = self.get_vol()
                except:vol = 0.0
                
            if vol == 0:
                self.resulteck.Text = str(0)
                self.resultrund.Text = str(0)
            else:
                if b and h:
                    v = self.tem_berechnen(vol,b,h,0)
                    self.resulteck.Text = str(v).replace('.',',')
                if d:
                    v = self.tem_berechnen(vol,0,0,d)
                    self.resultrund.Text = str(v).replace('.',',')

        elif self.modus_vol.IsChecked:
            if self.eingabetem.Text in ['',None,'0']:tem = 0.0
            else:tem = float(self.eingabetem.Text.replace(',','.'))
            if tem == 0:
                self.resulteck.Text = str(0)
                self.resultrund.Text = str(0)
            else:
                if b and h:
                    t = self.vol_berechnen(tem,b,h,0)
                    Text = str(t).replace('.',',')
                    nach = Text[Text.find(','):]
                    vor = Text[:Text.find(',')]
                    if len(vor) >= 7:
                        vor = vor[:-6] + '.' + vor[-6:-3] + '.' + vor[-3:] 
                    elif len(vor) >= 4:
                        vor = vor[:-3] + '.' + vor[-3:] 
                    else:pass
                    
                    self.resulteck.Text = vor + nach
                if d:
                    t = self.vol_berechnen(tem,0,0,d)
                    Text = str(t).replace('.',',')
                    nach = Text[Text.find(','):]
                    vor = Text[:Text.find(',')]
                    if len(vor) >= 7:
                        vor = vor[:-6] + '.' + vor[-6:-3] + '.' + vor[-3:] 
                    elif len(vor) >= 4:
                        vor = vor[:-3] + '.' + vor[-3:] 
                    else:pass
                    self.resultrund.Text = vor + nach

    def demension_berechn(self):
        if self.eingabevol.Text in ['',None,'0']:vol = 0.0
        else:vol = float(self.eingabevol.Text.replace(',','.'))
        if self.eingabetem.Text in ['',None,'0']:tem = 0.0
        else:tem = float(self.eingabetem.Text.replace(',','.'))
        if vol == 0.0 or tem == 0.0:
            return
        else:
            if self.modus_b.IsChecked:
                if self.eingabeh.Text in ['',None,'0']:h = 0
                else:h = int(self.eingabeh.Text)
                if h == 0:return
                else:
                    if self.bh_raster.Text in ['',None,'0']:raster = 0
                    else:raster = int(self.bh_raster.Text)
                    b = vol/tem*1000000/h / 3600
                    if raster == 0:b = self.ceil(b)
                    else:
                        if b%raster > 0:
                            b = self.ceil(b/raster) * raster
                    self.eingabeb.Text = str(int(b))
                
            elif self.modus_h.IsChecked:
                if self.eingabeb.Text in ['',None,'0']:b = 0
                else:b = int(self.eingabeb.Text)
                if b == 0:return
                else:
                    if self.bh_raster.Text in ['',None,'0']:raster = 0
                    else:raster = int(self.bh_raster.Text)
                    h = vol/tem*1000000/b / 3600
                    if raster == 0:h = self.ceil(h)
                    else:
                        if h%raster > 0:
                            h = self.ceil(h/raster) * raster
                    self.eingabeh.Text = str(int(h))
            elif self.modus_d.IsChecked:
                d = self.ceil(self.sqrt(float(vol)/tem / self.pi) * 100 / 3)
                d_temp = int(d)
                while (d_temp not in self.RundListe):
                    d_temp += 1
                    if d_temp > 1600:
                        d_temp = 1600
                        break
                try:self.liste_rohr.SelectedItem = d_temp
                except:pass

                self.eingabed.Text = str(int(d))

    def tem_berechnen(self, vol, b = 0, h = 0 , d = 0):
        if d == 0:
            v = round(vol / h / b * 1000000 / 3600,2)
            return v
        else:
            v = round(vol / (self.pi * d * d / 4000000) / 3600,2)
            return v

    def vol_berechnen(self, tem, b = 0, h = 0 , d = 0):
        if d == 0:
            vol = round(3600 * tem / 1000000 * b * h,2)
            return vol
        else:
            vol = round(3600 * tem * self.pi * d * d / 4000000,2)
            return vol

    def changetovol(self, sender, args):
        self.eingabevol.IsEnabled = False
        self.eingabetem.IsEnabled = True
        self.eingabeb.IsEnabled = True
        self.eingabeh.IsEnabled = True
        self.eingabed.IsEnabled = True
        self.modus_eck_art.Text = 'eck Volumenstrom:'
        self.modus_rund_art.Text = 'rund Volumenstrom:'
        self.modus_eck_einheit.Text = 'm³/h'
        self.modus_rund_einheit.Text = 'm³/h'
        self.resultrund.Visibility = self.Visible
        self.resulteck.Visibility = self.Visible  
        
        self.menge_berechn()

    def changetotem(self, sender, args):
        self.eingabevol.IsEnabled = True
        self.eingabetem.IsEnabled = False
        self.eingabeb.IsEnabled = True
        self.eingabeh.IsEnabled = True
        self.eingabed.IsEnabled = True
        self.modus_eck_art.Text = 'eck Geschwindigkeit:'
        self.modus_rund_art.Text = 'rund Geschwindigkeit:'
        self.modus_eck_einheit.Text = 'm/s'
        self.modus_rund_einheit.Text = 'm/s'
        self.resultrund.Visibility = self.Visible
        self.resulteck.Visibility = self.Visible
        self.menge_berechn()
    
    def changetob(self, sender, args):
        self.eingabevol.IsEnabled = True
        self.eingabetem.IsEnabled = True
        self.eingabeb.IsEnabled = False
        self.eingabeh.IsEnabled = True
        self.eingabed.IsEnabled = True
        self.modus_eck_art.Text = ''
        self.modus_rund_art.Text = ''
        self.modus_eck_einheit.Text = ''
        self.modus_rund_einheit.Text = ''
        self.resultrund.Visibility = self.Hidden
        self.resulteck.Visibility = self.Hidden
        self.demension_berechn()
    
    def changetoh(self, sender, args):
        self.eingabevol.IsEnabled = True
        self.eingabetem.IsEnabled = True
        self.eingabeb.IsEnabled = True
        self.eingabeh.IsEnabled = False
        self.eingabed.IsEnabled = True
        self.modus_eck_art.Text = ''
        self.modus_rund_art.Text = ''
        self.modus_eck_einheit.Text = ''
        self.modus_rund_einheit.Text = ''
        self.resultrund.Visibility = self.Hidden
        self.resulteck.Visibility = self.Hidden
        try:
            self.demension_berechn()
        except Exception as e:print(e)
    
    def changetod(self, sender, args):
        self.eingabevol.IsEnabled = True
        self.eingabetem.IsEnabled = True
        self.eingabeb.IsEnabled = True
        self.eingabeh.IsEnabled = True
        self.eingabed.IsEnabled = False
        self.modus_eck_art.Text = ''
        self.modus_rund_art.Text = ''
        self.modus_eck_einheit.Text = ''
        self.modus_rund_einheit.Text = ''
        self.resultrund.Visibility = self.Hidden
        self.resulteck.Visibility = self.Hidden
        self.demension_berechn()

    def changetooben(self, sender, args):
        self.obentext.IsEnabled = True
        self.mittetext.IsEnabled = False
        self.untentext.IsEnabled = False

    def changetomitte(self, sender, args):
        self.obentext.IsEnabled = False
        self.mittetext.IsEnabled = True
        self.untentext.IsEnabled = False

    def changetounten(self, sender, args):
        self.obentext.IsEnabled = False
        self.mittetext.IsEnabled = False
        self.untentext.IsEnabled = True

    def schreibenabmessung(self, sender, args):
        self.externaleventhandler.ExcuteApp = self.externaleventhandler.Dimension
        self.externalevent.Raise()

    def updategui(self, sender, args):
        self.externaleventhandler.ExcuteApp = self.externaleventhandler.GUIAktualiesieren
        self.externalevent.Raise()

    def schreibenheight(self, sender, args):
        self.externaleventhandler.ExcuteApp = self.externaleventhandler.HEIGHTAnpassen
        self.externalevent.Raise()
    def uebernehmen(self, sender, args):
        self.uebertragenforcheck = True
        if self.liste_rohr.SelectedIndex != -1:
            self.eingabed.Text = self.liste_rohr.SelectedItem.ToString()
    
    


gui = GUI()
gui.externaleventhandler.GUI = gui
gui.Show()