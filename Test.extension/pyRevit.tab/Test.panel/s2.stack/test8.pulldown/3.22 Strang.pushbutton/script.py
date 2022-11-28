# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
import os
from eventhandler import SELECT,config,DIMENSIONIEREN,UEBERNEHMEN,ExternalEvent,Key,TemplateItemBase,TemplateItem,Schwarz,Rot
from System.Collections.ObjectModel import ObservableCollection

__title__ = "1. Sprinkler-Strang dimensionieren"
__doc__ = """

Parameter: IGF_X_SM_Durchmesser

[2022.08.16]
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

logger = script.get_logger()

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.Liste = ObservableCollection[TemplateItem]()
        self.Key = Key
        self.Schwarz = Schwarz
        self.Rot = Rot
        self.minvalue = 0
        self.maxvalue = 100
        self.value = 1
        self.PB_text = ''
        self.select = SELECT()
        self.dimensionieren = DIMENSIONIEREN()
        self.uebernehmen = UEBERNEHMEN()
        self.selectevent = ExternalEvent.Create(self.select)
        self.dimensionierenevent = ExternalEvent.Create(self.dimensionieren)
        self.uebernehmenevent = ExternalEvent.Create(self.uebernehmen)
        # ProgressBarItem.__init__(self)
        self.config = config
        self.class_TemplateItem = TemplateItem
        self.allebauteile = None
        self.script = script
        self.read_config()
    
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.lv.ItemsSource = self.Liste
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
        

    def read_config(self):
        try:
            if self.config.Einstellungen:
                for el in self.config.Einstellungen:
                    self.Liste.Add(self.class_TemplateItem(el[2],el[0],el[1]))
        except:
            pass
        if self.Liste.Count == 0:
            self.Liste.Add(self.class_TemplateItem('20',1,None))
    
    def write_config(self):
        Liste = []
        try:
            for el in self.Liste:
                Liste.append([el.von,el.bis,el.dimension])
        except:
            pass
        self.config.Einstellungen = Liste
        self.script.save_config()
    
    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        else:
            args.Handled = True
            
    def von_changed(self, sender, args):
        value = sender.Text
        item = sender.DataContext
        index = self.Liste.IndexOf(item)
        item.von = value
        if value != '' and value != None:
            if index > 0:
                for n in range(1,self.Liste.Count):
                    alt = self.Liste.Item[n-1]
                    try:
                        alt.bis = str(int(self.Liste.Item[n].von)-1)
                        if int(alt.bis) < int(alt.von):
                            alt.farbe_bis = self.Rot
                            self.Liste.Item[n].farbe_von = self.Rot
                        else:
                            alt.farbe_bis = self.Schwarz
                            self.Liste.Item[n].farbe_von = self.Schwarz
                    except Exception as e:print(e)   

    def Add(self, sender, args):
        self.Liste.Add(self.class_TemplateItem('20',None,None))
        self.lv.Items.Refresh()

    def dele(self, sender, args):
        self.Liste.Remove(self.lv.SelectedItem)
        for n in range(1,self.Liste.Count):
            alt = self.Liste.Item[n-1]
            try:
                alt.bis = str(int(self.Liste.Item[n].von)-1)
                if alt.bis < alt.von:
                    alt.farbe_bis = Rot
                    self.Liste.Item[n].farbe_von = Rot
                else:
                    alt.farbe_bis = Schwarz
                    self.Liste.Item[n].farbe_von = Schwarz
            except:pass   
        self.lv.Items.Refresh()

    def cancel(self, sender, args):
        self.Close()
    
    def Selection_Changed(self,sender,args):
        if self.lv.SelectedIndex != -1:self.Remove.IsEnabled = True
        else:self.Remove.IsEnabled = False
    
    def auswaehlen(self,sender,args):
        self.selectevent.Raise()
    def dimension(self,sender,args):
        self.dimensionierenevent.Raise()
    def changeuebernehmen(self,sender,args):
        self.uebernehmenevent.Raise()

       
einstellung = AktuelleBerechnung()
einstellung.select.GUI = einstellung
einstellung.dimensionieren.GUI = einstellung
einstellung.uebernehmen.GUI = einstellung
einstellung.Show()