# coding: utf8
from pyrevit.userconfig import PyRevitConfig
import os

class Server_config(object):
    def __init__(self,path = r'R:\Vorlagen\_IGF\_pyRevit\config.ini'):
        if not os.path.exists(path):
            with open(path, 'w') as configfile:
                configfile.write('[Server Configuration]')
        self.server_config = PyRevitConfig(cfg_file_path=path,config_type='User')
    
    @property
    def server_config(self):
        return self._server_config
    
    @server_config.setter
    def server_config(self,value):
        self._server_config = value
    
    @staticmethod
    def getserverconfig(p):
        if p:return Server_config(p)
        else:return Server_config()

    @staticmethod
    def get_serverconfig(self,section = 'Default'):
        config = None
        try:
            config = Server_config().server_config.get_section(section)
        except:
            config = Server_config().server_config.add_section(section)
        return config

    def get_config(self,section = 'Default'):
        config = None
        try:
            config = self.server_config.get_section(section)
        except:
            config = self.server_config.add_section(section)
        return config
    
    def save_config(self):
        self.server_config.save_changes()    