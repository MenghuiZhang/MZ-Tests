# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import LISTE_SCHACHT,IS_Laboranschluss,IS_Schacht,ListeExternalEvent,ExternalEvent,MEPRaum,ObservableCollection,Laboranschlusss,labanschluss_Dict,LEVEL_Dict
from System.Windows.Input import Key
from System.Text.RegularExpressions import Regex

__title__ = "3.90 Schächte&Anlagen zuweisen"

__doc__ = """

[2022.10.17]
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
        self.labanschluss_Dict = labanschluss_Dict
        self.Levels = LEVEL_Dict
        self.LaboranschluesseDict = {}
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
        self.regex1 = Regex("[^0-9,]+")
        self.regex2 = Regex("[^0-9]+")
        self.regex3 = Regex("[^0-9-]+")
        self.IS_Laboranschluss = IS_Laboranschluss
        self.IS_Schacht = IS_Schacht
        self.schachtelemids = [el.elem.Id.ToString() for el in self.IS_Schacht]
        self.externaleventhandler = ListeExternalEvent()
        self.externalevent = ExternalEvent.Create(self.externaleventhandler)
        self.ItemtempMEP = ObservableCollection[MEPRaum]()
        self.ItemtempLabor = ObservableCollection[Laboranschlusss]()
        self.ItemtempLabor1 = ObservableCollection[Laboranschlusss]()
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        

        self.bezugsname.ItemsSource = sorted(self.berechnung_nach_Liste)
          
        self.Anschlussart.ItemsSource = self.IS_Laboranschluss
        self.Anschlussdetail.ItemsSource = self.IS_Laboranschluss

        self.LVSchacht.ItemsSource = self.IS_Schacht
        self.LVSchacht0.ItemsSource = self.ItemtempMEP
        self.labanschluss.ItemsSource = sorted(self.labanschluss_Dict.keys())
        self.ebene.ItemsSource = sorted(self.Levels.keys())
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

    def Datenaktualisieren(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.UpdateRaumInfo_Labor
        self.externalevent.Raise()

    def Schachtschreiben(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenSchacht0
        self.externalevent.Raise()
    
    def Schachtfinalschreiben(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenSchacht1
        self.externalevent.Raise()
    
    def anpassen(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.Laboranschlussanpassen
        self.externalevent.Raise()
    
    def laborplatz(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.SchreibenLaborDetail_Anschluss
        self.externalevent.Raise()
    
    def suchetextchanged(self, sender, args):
        text = sender.Text
        if not text:
            self.Anschlussart.ItemsSource = self.IS_Laboranschluss
        else:
            for el in self.IS_Laboranschluss:
                if el.Name.upper().find(text.upper()) != -1:
                    self.ItemtempLabor1.Add(el)
            self.Anschlussart.ItemsSource = self.ItemtempLabor1
        self.Anschlussart.Items.Refresh()
    
    def familychanged(self, sender, args):
        if self.labanschluss.SelectedIndex != -1:
            self.externaleventhandler.ExcuseApp = self.externaleventhandler.get_Laboranschluss
            self.externalevent.Raise()

    def bezugsnameselectionchanged(self, sender, args):
        self.einheit.Text = self.einheit_liste[sender.SelectedItem.ToString()]
    
    def Getluftmenge(self):
        ab24h = 0
        ablabmin = 0
        ablabmax = 0
        for el in self.IS_Laboranschluss:
            if el.Anzahl not in ['',None]:
                el.get_ganzahl()
                if el.art == '24h':
                    ab24h += el.ganzahl * el.min
                else:
                    ablabmin += el.ganzahl * el.min
                    ablabmax += el.ganzahl * el.max

        self.labminresult.Text = str(ablabmin)
        self.labmaxresult.Text = str(ablabmax)
        self.ab24hresult.Text = str(ab24h)
        
    def anchanged(self, sender, args):
        self.Getluftmenge()

    def Textinput0(self, sender, args):

        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True
    
    def Textinput3(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find('-') != -1:
                if args.Text == '-':args.Handled = True
                return
        try:
            args.Handled = self.regex3.IsMatch(args.Text)
        except:
            args.Handled = True
    
    def Textinput1(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find(',') != -1:
                if args.Text == ',':args.Handled = True
                return
        try:
            args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True
            
gui = GUI()
gui.externaleventhandler.GUI = gui
gui.Show()


