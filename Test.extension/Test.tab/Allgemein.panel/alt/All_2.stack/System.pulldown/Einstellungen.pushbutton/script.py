# coding: utf8
import clr
from pyrevit import revit, DB, script, UI
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from System.Windows import FontWeights
from pyIGF_config import Server_config

__title__ = "0. Familie"
__doc__ = """Grundeinstellungen für Systemtrennung festlegen
Version: 1.0"""
__author__ = "Menghui Zhang"

try:
    from pyIGF_logInfo import getlog
    getlog(__title__)
except:
    pass

class Familie(object):
    def __init__(self):
        self.SelID = 0
    @property
    def SelIS(self):
        return self._SelIS
    @SelIS.setter
    def SelIS(self,value):
        self._SelIS = value
    @property
    def SelID(self):
        return self._SelID
    @SelID.setter
    def SelID(self,value):
        self._SelID = value


class FamilieIS(object):
    def __init__(self):
        self.Name = ''
        self.id = 0
    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self,value):
        self._Name = value
    @property
    def id(self):
        return self._id
    @id.setter
    def id(self,value):
        self._id = value

logger = script.get_logger()
output = script.get_output()
config = script.get_config('Systemtrennung')
Server_Config = Server_config()
server_config = Server_Config.get_config('Systemtrennung')

uidoc = revit.uidoc
doc = revit.doc


coll = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
Verbinder = []
Verbindername = []

for el in coll:
    if el.FamilyCategoryId.ToString() in ['-2008010','-2008049']:
        Verbinder.append(el)


for el in Verbinder:
    name = el.Name
    Verbindername.append(name)
Verbindername.sort()

Liste_FAMIS = ObservableCollection[FamilieIS]()
dict_id_name = {}
dict_name_id = {}
for selid,el in enumerate(Verbindername):
    temp = FamilieIS()
    temp.Name = el
    temp.id = selid
    dict_id_name[selid] = el
    dict_name_id[el] = selid
    Liste_FAMIS.Add(temp)

def read_config(config):
    try:
        return config.Verbinder
    except:
        try:
            config.Verbinder = []
            return []
        except:
            return []

config_user = read_config(config)
config_server = read_config(server_config)

Liste_user = ObservableCollection[Familie]()
Liste_server = ObservableCollection[Familie]()
for elem in Verbindername:
    temp_user = Familie()
    temp_server = Familie()
    temp_user.SelIS = Liste_FAMIS
    temp_server.SelIS = Liste_FAMIS
    if elem in config_user:
        temp_user.SelID = dict_name_id[elem]
        Liste_user.Add(temp_user)
    if elem in config_server:
        temp_server.SelID = dict_name_id[elem]
        Liste_server.Add(temp_server)

class PlaeneUI(WPFWindow):
    def __init__(self, xaml_file_name,liste_user,liste_server):
        self.liste_user = liste_user
        self.liste_server = liste_server
        WPFWindow.__init__(self, xaml_file_name)
        try:
            self.confi = config.getconfig
            print(config.getconfig)
            if self.confi == True:
                self.ListView_Familie.ItemsSource = self.liste_server
                self.serv.FontWeight = FontWeights.Bold
                self.conf = server_config
            else:
                self.loca.FontWeight = FontWeights.Bold
                self.conf = config
                self.ListView_Familie.ItemsSource = self.liste_user
        except:
            self.conf = server_config
            try:
                config.getconfig = True
            except:
                pass
            self.serv.FontWeight = FontWeights.Bold
            self.ListView_Familie.ItemsSource = self.liste_server
    
    @property
    def server_config(self):
        return Server_Config.get_config('Systemtrennung')
    @property
    def user_config(self):
        return script.get_config('Systemtrennung')
    
    def write_config_user(self):       
        try:
            self.user_config.Verbinder = [dict_id_name[item.SelID] for item in self.liste_user]
        except:
            pass
        script.save_config()

    def write_config_server(self):       
        try:
            self.server_config.Verbinder = [dict_id_name[item.SelID] for item in self.liste_server]
        except:
            pass
        Server_Config.save_config()
    
    def Abfrage(self):
        buttons = UI.TaskDialogCommonButtons.No | UI.TaskDialogCommonButtons.Yes
        Task = UI.TaskDialog.Show('Abfrage','Möchten Sie die Server-Einstellungen ändern?',buttons)
        if Task.ToString() == 'Yes':
            return True
        elif Task.ToString() == 'No':
            return False

    def Server_Pruefen(self):
        conf = self.server_config.Verbinder
        neu = []
        for item in self.ListView_Familie.ItemsSource:
            name = dict_id_name[item.SelID]
            neu.append(name)
 
        pr = sorted(conf) != sorted(neu)
        return pr
     
    def ok(self, sender, args):
        if self.user_config.getconfig:
            if self.Server_Pruefen():
                if self.Abfrage():
                    self.write_config_server()
                    script.save_config()
                    Server_Config.save_config()
        else:
            self.write_config_user()
        self.Close()

    def anwenden(self, sender, args):
        if self.user_config.getconfig:
            if self.Server_Pruefen():
                if self.Abfrage():
                    self.write_config_server()
                    Server_Config.save_config()
                    script.save_config()
        else:
            self.write_config_user()

    def abbrechen(self, sender, args):
        self.Close()

    def serve(self, sender, args):
        self.serv.FontWeight = FontWeights.Bold
        self.loca.FontWeight = FontWeights.Normal
        try:
            if not self.user_config.getconfig:
                self.write_config_user()
                
        except:
            pass
        try:
            self.user_config.getconfig = True
            script.save_config()
        except:
            pass
        self.ListView_Familie.ItemsSource = self.liste_server
       
    def local(self, sender, args):
        self.loca.FontWeight = FontWeights.Bold
        self.serv.FontWeight = FontWeights.Normal
        try:
            if self.user_config.getconfig:
                if self.Server_Pruefen():
                    if self.Abfrage():
                        self.write_config_server()
                        Server_Config.save_config()  
                    else:
                        pass
        except:
            pass
        try:
            self.user_config.getconfig = False
            script.save_config()
        except:
            pass
        self.ListView_Familie.ItemsSource = self.liste_user
        
        
    def rueck(self, sender, args):
        try:
            config.Verbinder = server_config.Verbinder
            temp_coll = ObservableCollection[Familie]()
            for item in config.Verbinder:
                temp = Familie()
                temp.SelID = dict_name_id[item]
                temp.SelIS = Liste_FAMIS 
                temp_coll.Add(temp)
                self.ListView_Familie.ItemsSource = temp_coll
                self.liste_user = temp_coll
        except:
            pass
        self.loca.FontWeight = FontWeights.Bold
        self.serv.FontWeight = FontWeights.Normal

    def add(self, sender, args):
        temp = Familie()
        temp.SelIS = Liste_FAMIS
        self.ListView_Familie.ItemsSource.Add(temp)
        
    def dele(self, sender, args):
        if self.ListView_Familie.SelectedItem is not None:
            self.ListView_Familie.ItemsSource.Remove(self.ListView_Familie.SelectedItem)
    
Planfenster = PlaeneUI("window.xaml",Liste_user,Liste_server)

try:
    Planfenster.ShowDialog()
except Exception as e:
    logger.error(e)
    Planfenster.Close()