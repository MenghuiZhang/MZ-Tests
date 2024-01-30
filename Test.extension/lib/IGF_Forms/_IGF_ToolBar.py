# coding: utf8
from IGF_Forms import WPFPanel,forms
import os
from IGF_Forms._toolbareventhandler import ExternalEvent,ExternaleventListe,Grundriss,IS_Ebenen,DB,IS_DuctType,IS_PipeType,IS_FamilieType
from System.Text.RegularExpressions import Regex
from System.Windows.Input import Key,TraversalRequest,FocusNavigationDirection,Keyboard
from System.Windows import Visibility
from IGF_Allgemein._Rohrsegment import RundKanal
from pyrevit.forms import utils
from System.Windows.Input import MouseWheelEventArgs

XAML_FILES_DIR = os.path.dirname(__file__)

class ToolBar(WPFPanel):
    panel_title = "IGF_ToolBar"
    panel_id = "3110e336-f81c-4927-87da-4e0d39d4d64a"
    panel_source = os.path.join(XAML_FILES_DIR, '_IGF_ToolBar.xaml')

    def __init__(self):
        WPFPanel.__init__(self)
        self.regex = Regex("[^0-9.]+")
        self.regex1 = Regex("[^0-9]+")

        self.externaleventliste = ExternaleventListe()
        self.externaleventliste.GUI = self
        self.externalevent = ExternalEvent.Create(self.externaleventliste)

        self.ComboBox_winkel.ItemsSource = ['15','30','45','60','90']

        self.IS_Ebenen = IS_Ebenen
        self.CB_Reference.ItemsSource = self.IS_Ebenen

        self.IS_DuctType = IS_DuctType
        self.CB_DuctType.ItemsSource = self.IS_DuctType

        self.IS_PipeType = IS_PipeType
        self.CB_PipeType.ItemsSource = self.IS_PipeType

        self.IS_FamilieType = IS_FamilieType
        self.CB_FamilieType.ItemsSource = self.IS_FamilieType

        self.Key = Key
        self.Keyboard = Keyboard
        self.Visible = Visibility.Visible
        self.Hidden = Visibility.Hidden
        self.uebertragenforcheck = False

        self.volumenstromeingaben = Regex("[^0-9+-]+")
        self.tempoeingaben = Regex("[^0-9.]+")
        self.regex2 = Regex("[^0-9]+")
        self.TraversalRequest = TraversalRequest
        self.focusNavigationDirection = FocusNavigationDirection.Next

        self.RundKanal = RundKanal
        self.CB_RundRohr.ItemsSource = sorted(self.RundKanal)
        self.kanalberechnen = True

        self.cls_MouseWheelEventArgs = MouseWheelEventArgs

        self.SchnitteGrundriss = Grundriss(None)
        self.ListeSchnitte = self.SchnitteGrundriss.schnitt_Liste
        self.dg_Schnitte.ItemsSource = self.ListeSchnitte
    
    def dg_mausewhell(self, sender, e):
        tmp = self.cls_MouseWheelEventArgs(e.MouseDevice,e.Timestamp,e.Delta)
        tmp.RoutedEvent = self.SV_Familie.MouseWheelEvent
        tmp.Source = sender
        self.SV_Familie.RaiseEvent(tmp)
        e.Handled = True
    
    def dg_mausewhell_Projekt(self, sender, e):
        tmp = self.cls_MouseWheelEventArgs(e.MouseDevice,e.Timestamp,e.Delta)
        tmp.RoutedEvent = self.SV_Projekt.MouseWheelEvent
        tmp.Source = sender
        self.SV_Projekt.RaiseEvent(tmp)
        e.Handled = True

    def IsShowed(self,sender,e):
        # self.setup_owner()
        self.externaleventliste.ExcuteApp = self.externaleventliste.set_up
        self.externalevent.Raise()
    
    def BT_ChangeFamilietyp_Aktive(self,sender,e):
        self.BT_ChangeFamilietyp.IsEnabled = True
    
    def ReferenzFamilieAktualisieren(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ReferenzFamilieAktualisieren
        self.externalevent.Raise()

    def ChangeFamilietyp(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ChangeFamilietyp
        self.externalevent.Raise()
    
    def SchnitteVerschieben(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.SchnitteVerschieben
        self.externalevent.Raise()
    
    def KanalstueckEinfuegen(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.KanalstueckEinfuegen
        self.externalevent.Raise()
    
    def RohrstueckEinfuegen(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.RohrstueckEinfuegen
        self.externalevent.Raise()
    
    def SharedParameterdurchsuchen(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.SharedParameterdurchsuchen
        self.externalevent.Raise()
    
    def Familie_BT_Reset(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.FamilyParameterAktualiesieren
        self.externalevent.Raise()

    def ElementTrennen(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ElementTrennen
        self.externalevent.Raise()

    def Familie_BT_Erstellen(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.FamilyParameterErstellen
        self.externalevent.Raise()
    
    def ElementVerbinden(self,sender,e):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ElementVerbinden
        self.externalevent.Raise()
    
    def UebergangAktiv(self,sender,e):
        self.GB_Ubergang.IsEnabled = True
        self.GB_Versprung.IsEnabled = False
    
    def UebergangInaktiv(self,sender,e):
        self.GB_Ubergang.IsEnabled = False
        self.GB_Versprung.IsEnabled = True
    
    def Erstellen_Auto_Aktiv(self,sender,e):
        self.RB_Maneull_length.IsEnabled = False
    
    def Erstellen_Auto_inaktiv(self,sender,e):
        self.RB_Maneull_length.IsEnabled = True
    
    def Anpassen_Auto_Aktiv(self,sender,e):
        self.RB_Maneull_length_anpassen.IsEnabled = False
    
    def Anpassen_Auto_Inaktiv(self,sender,e):
        self.RB_Maneull_length_anpassen.IsEnabled = True
        
    def Verbindungerstellen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.Verbindungerstellen
        self.externalevent.Raise()
    
    def UebergangAnpassen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.UebergangAnpassen
        self.externalevent.Raise()
    
    def CloseWindow(self, sender, args):
        self.Close()
    
    def movewindow(self, sender, args):
        self.DragMove()
    
    def AbstandAnpassen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.AbstandAnpassen
        self.externalevent.Raise()
    
    def ElemAnzeigen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ElemAnzeigen
        self.externalevent.Raise()
    
    def ElemAuswaehlen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ElemAuswaehlen
        self.externalevent.Raise()
    
    def RaumAnzeigen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.RaumAnzeigen
        self.externalevent.Raise()
    
    def RaumAuswaehlen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.RaumAuswaehlen
        self.externalevent.Raise()


    def textinput(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find('.') != -1:
                if args.Text == '.':args.Handled = True
                return
        try:
            args.Handled = self.regex.IsMatch(args.Text)
        except:
            args.Handled = True
        
    def textinput1(self, sender, args):
        try:
            args.Handled = self.regex.IsMatch(args.Text)
        except:
            args.Handled = True
    
    def VolumenstromEingaben(self, sender, args):
        try:args.Handled = self.volumenstromeingaben.IsMatch(args.Text)
        except:args.Handled = True
    
    def Nichtnurganzzahl(self, sender, args):
        try:
            if sender.Text in ['',None]:
                args.Handled = self.tempoeingaben.IsMatch(args.Text)
            elif sender.Text.find('.') != -1 and args.Text == '.':
                args.Handled = True
            else:
                args.Handled = self.tempoeingaben.IsMatch(args.Text)
        except:
            args.Handled = True

    def Nurganzzahl(self, sender, args):
        try: args.Handled = self.regex2.IsMatch(args.Text)
        except: args.Handled = True

    def Setkey3(self, sender, args):
        if args.Key == self.Key.Tab:
            try:self.MoveFocus(self.TraversalRequest(self.focusNavigationDirection))
            except Exception as e:print(e)
        elif args.Key == self.Key.Enter:
            try:self.Keyboard.FocusedElement.IsChecked = True
            except Exception as e:print(e)
        else:
            args.Handled = True

    def KanalBerechnen(self, sender, args):
        try:
            self._Kanalberechnen()
            if self.RB_Durchmesser.IsChecked:
                if self.kanalberechnen:
                    self._KanalDimensionberechnen()
                
            else:
                self._KanalDimensionberechnen()
        except Exception as e:print(e)
    
    def get_volumenstrom(self):
        Text = self.TB_Volumenstrom.Text
        vol = 0.0
        try:vol = eval(Text)
        except:pass
        return vol
    
    def TextBox_GotFocus(self,sender,e):
        tb = sender
        tb.SelectAll()

    def _Kanalberechnen(self):
        if self.TB_Breite.Text in ['',None,'0']:b = 0
        else:b = float(self.TB_Breite.Text)
        if self.TB_Hoehe.Text in ['',None,'0']:h = 0
        else:h = float(self.TB_Hoehe.Text)
        if self.TB_Durchmesser.Text in ['',None,'0']:d = 0
        else:d = float(self.TB_Durchmesser.Text)

        if self.RB_Volumenstrom.IsChecked:
            if self.TB_Volumenstrom.Text in ['',None,'0']:vol = 0.0
            else:vol = self.get_volumenstrom()
                
            if vol == 0:
                self.TB_ResulEck.Text = str(0)
                self.TB_ResulRund.Text = str(0)
            else:
                if b and h:
                    tempo = self.externaleventliste.Tempo_berechnen(vol,b,h)
                    self.TB_ResulEck.Text = str(tempo)
                if d:
                    tempo = self.externaleventliste.Tempo_berechnen(vol,None,None,d)
                    self.TB_ResulRund.Text = str(tempo)

        elif self.RB_Tempo.IsChecked:
            if self.TB_Tempo.Text in ['',None,'0']:tempo = 0.0
            else:tempo = float(self.TB_Tempo.Text.replace(',','.'))
            if tempo == 0:
                self.TB_ResulEck.Text = str(0)
                self.TB_ResulRund.Text = str(0)
            else:
                if b and h:
                    vol = self.externaleventliste.Volumenstrom_berechnen(tempo,b,h)
                    Text = str(vol)
                    nach = Text[Text.find('.'):]
                    vor = Text[:Text.find('.')]
                    if len(vor) >= 7:
                        vor = vor[:-6] + ' ' + vor[-6:-3] + ' ' + vor[-3:] 
                    elif len(vor) >= 4:
                        vor = vor[:-3] + ' ' + vor[-3:] 
                    else:pass
                    self.TB_ResulEck.Text = vor + nach

                if d:
                    vol = self.externaleventliste.Volumenstrom_berechnen(tempo,None,None,d)
                    Text = str(vol)
                    nach = Text[Text.find('.'):]
                    vor = Text[:Text.find('.')]
                    if len(vor) >= 7:
                        vor = vor[:-6] + ' ' + vor[-6:-3] + ' ' + vor[-3:] 
                    elif len(vor) >= 4:
                        vor = vor[:-3] + ' ' + vor[-3:] 
                    else:pass
                    self.TB_ResulRund.Text = vor + nach

    def _KanalDimensionberechnen(self):
        if self.TB_Volumenstrom.Text in ['',None,'0']:vol = 0.0
        else:vol = float(self.TB_Volumenstrom.Text.replace(',','.'))
        if self.TB_Tempo.Text in ['',None,'0']:tem = 0.0
        else:tem = float(self.TB_Tempo.Text.replace(',','.'))
        if vol == 0.0 or tem == 0.0:
            return
        
        else:
            vol = self.get_volumenstrom()
            if self.RB_Hoehe.IsChecked:
                if self.TB_Hoehe.Text in ['',None,'0']:h = 0
                else:h = float(self.TB_Hoehe.Text)
                if h == 0:return
                else:
                    if self.RB_Raster.Text in ['',None,'0']:raster = 0
                    else:raster = int(float(self.RB_Raster.Text))
                    b = self.externaleventliste.Dimension_berechnen(vol,tem,h,None,None)
                    if raster == 0:pass
                    else:
                        if b%raster > 0:
                            b = int(b/raster) * raster + raster
                    self.TB_Breite.Text = str(int(b))
                
            elif self.RB_Bereite.IsChecked:
                if self.TB_Breite.Text in ['',None,'0']:b = 0
                else:b = float(self.TB_Breite.Text)
                if b == 0:return
                else:
                    if self.RB_Raster.Text in ['',None,'0']:raster = 0
                    else:raster = int(float(self.RB_Raster.Text))
                    h = self.externaleventliste.Dimension_berechnen(vol,tem,b,None,None)
                    if raster == 0:
                        pass
                    else:
                        if h%raster > 0:
                            h = int(h/raster) * raster + raster
                    self.TB_Hoehe.Text = str(int(h))

            elif self.RB_Durchmesser.IsChecked:
                d = self.externaleventliste.Dimension_berechnen(vol,tem,None,None,None)
                d_temp = float(d)
                while (d_temp not in self.RundKanal):
                    d_temp += 1
                    if d_temp > 1600:
                        d_temp = 1600
                        break
                try:self.CB_RundRohr.SelectedItem = d_temp
                except:pass

                self.TB_Durchmesser.Text = str(int(d))
    
    def Modus_Vol_Actived(self, sender, args):
        self.TB_Volumenstrom.IsEnabled = True
        self.TB_Tempo.IsEnabled = False
        self.TB_Breite.IsEnabled = True
        self.TB_Hoehe.IsEnabled = True
        self.TB_Durchmesser.IsEnabled = True
        self.modus_eck_art.Text = 'eck Geschwindigkeit:'
        self.modus_rund_art.Text = 'rund Geschwindigkeit:'
        self.modus_eck_einheit.Text = 'm/s'
        self.modus_rund_einheit.Text = 'm/s'
        self.TB_ResulRund.Visibility = self.Visible
        self.TB_ResulEck.Visibility = self.Visible  
        self._Kanalberechnen()

    def Modus_Tempo_Actived(self, sender, args):
        self.TB_Volumenstrom.IsEnabled = False
        self.TB_Tempo.IsEnabled = True
        self.TB_Breite.IsEnabled = True
        self.TB_Hoehe.IsEnabled = True
        self.TB_Durchmesser.IsEnabled = True
        self.modus_eck_art.Text = 'eck Volumenstrom:'
        self.modus_rund_art.Text = 'rund Volumenstrom:'
        self.modus_eck_einheit.Text = 'm³/h'
        self.modus_rund_einheit.Text = 'm³/h'
        self.TB_ResulRund.Visibility = self.Visible
        self.TB_ResulEck.Visibility = self.Visible
        self._Kanalberechnen()
    
    def Modus_Hoehe_Actived(self, sender, args):
        self.TB_Volumenstrom.IsEnabled = True
        self.TB_Tempo.IsEnabled = True
        self.TB_Breite.IsEnabled = False
        self.TB_Hoehe.IsEnabled = True
        self.TB_Durchmesser.IsEnabled = False
        self.modus_eck_art.Text = ''
        self.modus_rund_art.Text = ''
        self.modus_eck_einheit.Text = ''
        self.modus_rund_einheit.Text = ''
        self.TB_ResulRund.Visibility = self.Hidden
        self.TB_ResulEck.Visibility = self.Hidden
        try:self._KanalDimensionberechnen()
        except Exception as e:print(e)
    
    def Modus_Breite_Actived(self, sender, args):
        self.TB_Volumenstrom.IsEnabled = True
        self.TB_Tempo.IsEnabled = True
        self.TB_Breite.IsEnabled = True
        self.TB_Hoehe.IsEnabled = False
        self.TB_Durchmesser.IsEnabled = False
        self.modus_eck_art.Text = ''
        self.modus_rund_art.Text = ''
        self.modus_eck_einheit.Text = ''
        self.modus_rund_einheit.Text = ''
        self.TB_ResulRund.Visibility = self.Hidden
        self.TB_ResulEck.Visibility = self.Hidden
        try:self._KanalDimensionberechnen()
        except Exception as e:print(e)
    
    def Modus_Durchmesser_Actived(self, sender, args):
        self.TB_Volumenstrom.IsEnabled = True
        self.TB_Tempo.IsEnabled = True
        self.TB_Breite.IsEnabled = False
        self.TB_Hoehe.IsEnabled = False
        self.TB_Durchmesser.IsEnabled = False
        self.modus_eck_art.Text = ''
        self.modus_rund_art.Text = ''
        self.modus_eck_einheit.Text = ''
        self.modus_rund_einheit.Text = ''
        self.TB_ResulRund.Visibility = self.Hidden
        self.TB_ResulEck.Visibility = self.Hidden
        try:self._KanalDimensionberechnen()
        except Exception as e:print(e)

    def changetooben(self, sender, args):
        self.TB_Oben.IsEnabled = True
        self.TB_Mitte.IsEnabled = False
        self.TB_Unten.IsEnabled = False

    def changetomitte(self, sender, args):
        self.TB_Oben.IsEnabled = False
        self.TB_Mitte.IsEnabled = True
        self.TB_Unten.IsEnabled = False

    def changetounten(self, sender, args):
        self.TB_Oben.IsEnabled = False
        self.TB_Mitte.IsEnabled = False
        self.TB_Unten.IsEnabled = True

    def ChangeDimensionDuct(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ChangeDimensionDuct
        self.externalevent.Raise()

    def KanaldatenAusRevit(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.KanaldatenAusRevit
        self.externalevent.Raise()

    def ChangeHeightDuct(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.ChangeHeightDuct
        self.externalevent.Raise()
    
    def RundDimensionUebernehmen(self, sender, args):
        self.kanalberechnen = False
        if self.CB_RundRohr.SelectedIndex != -1:
            self.TB_Durchmesser.Text = str(int(self.CB_RundRohr.SelectedItem))
        self.kanalberechnen = True
    
       
        