# coding: utf8
import clr
from pyrevit import revit, DB, script, UI
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from System.Windows import Visibility,Thickness  
from System.Windows import FontWeights
from config import Server_config
from userinfo import getlog,getloglocal
from family import CFI
from System.Windows.Input import Key

__title__ = "Einstellungen"
__doc__ = """Grundeinstellungen von Pläne festlegen"""
__author__ = "Menghui Zhang"

try:getlog(__title__)
except:pass
try:getloglocal(__title__)
except:pass

logger = script.get_logger()
config = script.get_config('Plaene_local')
serverconfiglocal = script.get_config('Plaene_server')
serverpath = script.get_config('serverconfig')

try:
    if serverpath.path:
        server_config = Server_config(serverpath.path)
    else:
        server_config = Server_config()
except:
    server_config = Server_config()

serverconfig = server_config.get_config('Plaene')

uidoc = revit.uidoc
doc = revit.doc

hide = Visibility.Hidden
show = Visibility.Visible

titleblock = CFI.get_alltbType()
viewport = CFI.get_alllafType()

class PlaeneUI(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, 'window.xaml')        
        try:
            self.read_config(self.conf)
        except:
            pass
        self.plankopf.ItemsSource = titleblock
        self.HA.ItemsSource = viewport
        self.LG.ItemsSource = viewport     
        
    def turntolocal(self):
        try:self.bz_l.Text = config.bz_l
        except:self.bz_l.Text = config.bz_l = '10'
        try:self.bz_r.Text = config.bz_r
        except:self.bz_r.Text = config.bz_r = '10'
        try:self.bz_o.Text = config.bz_o
        except:self.bz_o.Text = config.bz_o = '10'
        try:self.bz_u.Text = config.bz_u
        except:self.bz_u.Text = config.bz_u = '10'
        try:
            for n,el in enumerate(viewport):
                try:
                    if el.name == config.HA:
                        self.HA.SelectedIndex = n
                except:pass
                try:
                    if el.name == config.LG:
                        self.LG.SelectedIndex = n
                except:pass
        except:pass
        try:
            for n,el in enumerate(titleblock):
                try:
                    if el.name == config.plankopf:
                        self.plankopf.SelectedIndex = n
                except:pass
        except:pass

        try:self.pk_l.Text = config.pk_l
        except:self.pk_l.Text = config.pk_l = '20'
        try:self.pk_r.Text = config.pk_r
        except:self.pk_r.Text = config.pk_r = '5'
        try:self.pk_o.Text = config.pk_o
        except:self.pk_o.Text = config.pk_o = '5'
        try:self.pk_u.Text = config.pk_u
        except:self.pk_u.Text = config.pk_u = '5'

        try:self.raster.IsChecked = config.raster
        except:self.raster.IsChecked = config.raster = True
        try:self.Haupt.IsChecked = config.Haupt
        except:self.Haupt.IsChecked = config.Haupt = True
        try:self.legend.IsChecked = config.legend
        except:self.legend.IsChecked = config.legend = True

        try:self.gewerke.Text = config.gewerke
        except:pass
    
    def nichteditable(self):
        try:self.bz_l.IsEnabled = False
        except:pass
        try:self.bz_r.IsEnabled = False
        except:pass
        try:self.bz_o.IsEnabled = False
        except:pass
        try:self.bz_u.IsEnabled = False
        except:pass
        try:self.HA.IsEnabled = False
        except:pass
        try:self.LG.IsEnabled = False
        except:pass
        try:self.plankopf.IsEnabled = False
        except:pass
        try:self.pk_l.IsEnabled = False
        except:pass
        try:self.pk_r.IsEnabled = False
        except:pass
        try:self.pk_o.IsEnabled = False
        except:pass
        try:self.pk_u.IsEnabled = False
        except:pass
        try:self.raster.IsEnabled = False
        except:pass
        try:self.Haupt.IsEnabled = False
        except:pass
        try:self.legend.IsEnabled = False
        except:pass
        try:self.gewerke.IsEnabled = False
        except:pass
    
    def editable(self):
        try:self.bz_l.IsEnabled = True
        except:pass
        try:self.bz_r.IsEnabled = True
        except:pass
        try:self.bz_o.IsEnabled = True
        except:pass
        try:self.bz_u.IsEnabled = True
        except:pass
        try:self.HA.IsEnabled = True
        except:pass
        try:self.LG.IsEnabled = True
        except:pass
        try:self.plankopf.IsEnabled = True
        except:pass
        try:self.pk_l.IsEnabled = True
        except:pass
        try:self.pk_r.IsEnabled = True
        except:pass
        try:self.pk_o.IsEnabled = True
        except:pass
        try:self.pk_u.IsEnabled = True
        except:pass
        try:self.raster.IsEnabled = True
        except:pass
        try:self.Haupt.IsEnabled = True
        except:pass
        try:self.legend.IsEnabled = True
        except:pass
        try:self.gewerke.IsEnabled = True
        except:pass

    def turntoserver(self):
        try:self.bz_l.Text = serverconfig.bz_l
        except:self.bz_l.Text = serverconfig.bz_l = '10'
        try:self.bz_r.Text = serverconfig.bz_r
        except:self.bz_r.Text = serverconfig.bz_r = '10'
        try:self.bz_o.Text = serverconfig.bz_o
        except:self.bz_o.Text = serverconfig.bz_o = '10'
        try:self.bz_u.Text = serverconfig.bz_u
        except:self.bz_u.Text = serverconfig.bz_u = '10'
        try:
            for n,el in enumerate(viewport):
                try:
                    if el.name == serverconfig.HA:
                        self.HA.SelectedIndex = n
                except:pass
                try:
                    if el.name == serverconfig.LG:
                        self.LG.SelectedIndex = n
                except:pass
        except:pass
        try:
            for n,el in enumerate(titleblock):
                try:
                    if el.name == serverconfig.plankopf:
                        self.plankopf.SelectedIndex = n
                except:pass
        except:pass

        try:self.pk_l.Text = serverconfig.pk_l
        except:self.pk_l.Text = serverconfig.pk_l = '20'
        try:self.pk_r.Text = serverconfig.pk_r
        except:self.pk_r.Text = serverconfig.pk_r = '5'
        try:self.pk_o.Text = serverconfig.pk_o
        except:self.pk_o.Text = serverconfig.pk_o = '5'
        try:self.pk_u.Text = serverconfig.pk_u
        except:self.pk_u.Text = serverconfig.pk_u = '5'

        try:self.raster.IsChecked = serverconfig.raster
        except:self.raster.IsChecked = serverconfig.raster = True
        try:self.Haupt.IsChecked = serverconfig.Haupt
        except:self.Haupt.IsChecked = serverconfig.Haupt = True
        try:self.legend.IsChecked = serverconfig.legend
        except:self.legend.IsChecked = serverconfig.legend = True

        try:self.gewerke.Text = serverconfig.gewerke
        except:pass

    def read_config(self):
        if config.local:
            self.loca.IsChecked == True
            self.turntolocal()
            self.editable()
        else:
            self.serv.IsChecked == True
            self.turntoserver()
            self.nichteditable()

    def write_config(self,configdatei):
        try:configdatei.bz_l = self.bz_l.Text
        except:pass
        try:configdatei.bz_r = self.bz_r.Text
        except:pass
        try:configdatei.bz_o = self.bz_o.Text
        except:pass
        try:configdatei.bz_u = self.bz_u.Text
        except:pass
        try:configdatei.HA = self.HA.SelectedItem.name            
        except:pass
        try:configdatei.LG = self.LG.SelectedItem.name            
        except:pass
        try:configdatei.plankopf = self.plankopf.SelectedItem.name            
        except:pass
        try:configdatei.pk_l = self.pk_l.Text
        except:pass
        try:configdatei.pk_r = self.pk_r.Text
        except:pass
        try:configdatei.pk_o = self.pk_o.Text
        except:pass
        try:configdatei.pk_u = self.pk_u.Text
        except:pass
        try:configdatei.raster = self.raster.IsChecked
        except:pass
        try:configdatei.Haupt = self.Haupt.IsChecked
        except:pass
        try:configdatei.legend = self.legend.IsChecked
        except:pass
        try:configdatei.gewerke = self.gewerke.Text
        except:pass

    def rueck(self, sender, args):
        self.loca.IsChecked = True
        self.turntoserver()
        self.write_config(config)
        self.write_config(serverconfiglocal)

    def ok(self, sender, args):
        if self.loca.IsChecked:
            self.write_config(config)
            script.save_config()
        else:
            self.write_config(serverconfiglocal)
            script.save_config()
            self.write_config(serverconfig)
            server_config.save_config()
        self.Close()
  
    def abbrechen(self, sender, args):
        self.Close()
        script.exit()

    def serve(self, sender, args):
        self.turntoserver()
        self.nichteditable()
       
    def Local(self, sender, args):
        self.turntolocal()
        self.editable()
    
    def Scaledeaktiv(self, sender, args):
        return
    def Scaleaktiv(self, sender, args):
        return
    def richtungdeaktiv(self, sender, args):
        return
    def richtungaktiv(self, sender, args):
        return
    def Setkey(self, sender, args):   
        if ((args.Key >= Key.D0 and args.Key <= Key.D9) or (args.Key >= Key.NumPad0 and args.Key <= Key.NumPad9) \
            or args.Key == Key.Delete or args.Key == Key.Back or args.Key == Key.Enter):
            args.Handled = False
        else:
            args.Handled = True
        
Planfenster = PlaeneUI()

try:
    Planfenster.ShowDialog()
except Exception as e:
    logger.error(e)
    Planfenster.Close()
    script.exit()
