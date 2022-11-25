# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import LISTE_SCHACHT,IS_Laboranschluss,IS_Schacht,ListeExternalEvent,ExternalEvent,MEPRaum,ObservableCollection,Laboranschlusss
from System.Windows.Input import Key

__title__ = "Raumluft"

__doc__ = """

Min/Max/Nacht/TiefeNacht in Parameter Volumenstrom übertragen

[2022.05.19]
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
        self.berechnung_nach_Liste = {
            "Fläche",
            "Luftwechsel",
            "Person",
            "manuell",
            "nurZUMa",
            "nurABMa",
            "nurZU_Fläche",
            "nurZU_Luftwechsel",
            "nurZU_Person",
            "nurAB_Fläche",
            "nurAB_Luftwechsel",
            "nurAB_Person",
            "keine",
                    }
        self.berechnung_nach = {
            "Fläche":'1',
            "Luftwechsel":'2',
            "Person":'3',
            "manuell":'4',
            "nurZUMa":'5',
            "nurABMa":'6',
            "nurZU_Fläche":'5.1',
            "nurZU_Luftwechsel":'5.2',
            "nurZU_Person":'5.3',
            "nurAB_Fläche":'6.1',
            "nurAB_Luftwechsel":'6.2',
            "nurAB_Person":'6.3',
            "keine":'9',
                    }
        self.einheit_liste = {
            "Fläche":'m³/h pro m²',
            "Luftwechsel":'-1/h',
            "Person":'m3/h pro P',
            "manuell":'m³/h',
            "nurZUMa":'m³/h',
            "nurABMa":'m³/h',
            "nurZU_Fläche":'m³/h pro m²',
            "nurZU_Luftwechsel":'-1/h',
            "nurZU_Person":'m3/h pro P',
            "nurAB_Fläche":'m³/h pro m²',
            "nurAB_Luftwechsel":'-1/h',
            "nurAB_Person":'m3/h pro P',
            "keine":'',
                }
        self.LISTE_SCHACHT = LISTE_SCHACHT
        self.IS_Laboranschluss = IS_Laboranschluss
        self.IS_Schacht = IS_Schacht
        self.schachtelemids = [el.elem.Id.ToString() for el in self.IS_Schacht]
        self.externaleventhandler = ListeExternalEvent()
        self.externalevent = ExternalEvent.Create(self.externaleventhandler)
        self.ItemtempMEP = ObservableCollection[MEPRaum]()
        self.ItemtempLabor = ObservableCollection[Laboranschlusss]()
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        

        self.bezugsname.ItemsSource = sorted(self.berechnung_nach_Liste)
          
        self.Anschlussart.ItemsSource = self.IS_Laboranschluss
        self.Anschlussdetail.ItemsSource = self.ItemtempLabor

        self.LVSchacht.ItemsSource = self.IS_Schacht
        self.LVSchacht0.ItemsSource = self.ItemtempMEP
        self.set_up_is()
    
    def set_up_is(self):
        self.SchachtRZU.ItemsSource = sorted(self.LISTE_SCHACHT)
        self.SchachtRAB.ItemsSource = sorted(self.LISTE_SCHACHT)
        self.SchachtTZU.ItemsSource = sorted(self.LISTE_SCHACHT)
        self.SchachtTAB.ItemsSource = sorted(self.LISTE_SCHACHT)
        self.Schacht24h.ItemsSource = sorted(self.LISTE_SCHACHT)
        self.SchachtLAB.ItemsSource = sorted(self.LISTE_SCHACHT)
        
    def auswahl(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SelectRaum
        self.externalevent.Raise()
    
    def schreiben_schacht(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenSchachtAnlagen
        self.externalevent.Raise()
    
    def raunluftschreiben(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenRaumluft
        self.externalevent.Raise()

    def manuelllab(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenLabor
        self.externalevent.Raise()
    
    def Labordetailsschreiben(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenLaborDetail
        self.externalevent.Raise()
    
    def Schachtakt(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.UpdateRaumInfo
        self.externalevent.Raise()

    def Schachtschreiben(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenSchacht0
        self.externalevent.Raise()
    
    def Schachtfinalschreiben(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenSchacht1
        self.externalevent.Raise()
    
    def bezugsnameselectionchanged(self, sender, args):
        self.einheit.Text = self.einheit_liste[sender.SelectedItem.ToString()]
        
    def anchanged(self, sender, args):
        self.ItemtempLabor.Clear()
        ab24h = 0
        ablabmin = 0
        ablabmax = 0
        for el in self.IS_Laboranschluss:
            if el.Anzahl not in ['',None]:
                self.ItemtempLabor.Add(el)
                el.get_ganzahl()
                if el.art == '24h':
                    ab24h += el.ganzahl * el.min
                else:
                    ablabmin += el.ganzahl * el.min
                    ablabmax += el.ganzahl * el.max
        
        self.Anschlussdetail.ItemsSource = self.ItemtempLabor
        self.Anschlussdetail.Items.Refresh()
        self.labminresult.Text = str(ablabmin)
        self.labmaxresult.Text = str(ablabmax)
        self.ab24hresult.Text = str(ab24h)
    
    def Setkey(self, sender, args):
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
    
    def Setkey1(self, sender, args):
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        
        else:
            args.Handled = True
            
gui = GUI()
gui.externaleventhandler.GUI = gui
gui.Show()
