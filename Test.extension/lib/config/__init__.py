# coding: utf8
import ConfigParser
import os
import json
home_dir = os.path.expanduser("~") + '\\DC\\ACCDocs\\IGF\\IGF_pyrevit\\Projektdateien\\xx_Projektconfig'

appdata = os.getenv('APPDATA')

class ProjektConfig(object):
    def __init__(self,nummername = '',lokal = False):
        self.config = ConfigParser.ConfigParser()
        self.nummername = nummername
        self.home_dir_0 = os.path.expanduser("~") + '\\DC\\ACCDocs\\IGF\\IGF_pyrevit\\Project Files\\xx_Projektconfig'
        self.home_dir_1 = os.path.expanduser("~") + '\\DC\\ACCDocs\\IGF\\IGF_pyrevit\\Projektdateien\\xx_Projektconfig'
        if os.path.exists(self.home_dir_0):
            self.config_rvt = self.home_dir_0 + '\\' + self.nummername
        else:
            self.config_rvt = self.home_dir_1 + '\\' + self.nummername
            
        self.configfile = appdata + '\\pyRevit'
        self.lokal = lokal
        self._cfg_file_path_local = self.configfile + '\\pyRevit_config.conf'
        self._cfg_file_path_cental = self.config_rvt + '\\pyRevit_config.conf'
        self.isconfig = True
        self._cfg_file_path = self._cfg_file_path_cental
        self.config_ermitterl()

        if not self.isconfig:return
        
        self.config.read(self._cfg_file_path)
        self.section = ''
        
        
    
    @property
    def config_file(self):
        """Current config file path."""
        return self._cfg_file_path

    def config_ermitterl(self):
        if self.lokal:
            self._cfg_file_path = self._cfg_file_path_local
            if not os.path.exists(self._cfg_file_path_local):
                if not os.path.exists(self.configfile):
                    try:
                        os.makedirs(self.configfile)
                    except:
                        self.isconfig = False
                        return
                else:
                    try:
                        with open(self._cfg_file_path_local,'w') as file:
                            pass
                        return
                    except:
                        self.isconfig = False
                        return
            else:
                return

        if not os.path.exists(self._cfg_file_path_cental):
            self._cfg_file_path = self._cfg_file_path_cental
            if not os.path.exists(self.config_rvt):
                try:
                    os.makedirs(self.config_rvt)
                except:
                    try:
                        self.lokal = True
                        if not os.path.exists(self._cfg_file_path_local):
                            self._cfg_file_path = self._cfg_file_path_local
                            if not os.path.exists(self.configfile):
                                try:
                                    os.makedirs(self.configfile)
                                except:
                                    self.isconfig = False
                                    return
                            else:
                                try:
                                    with open(self._cfg_file_path_local,'w') as file:
                                        pass
                                except:
                                    self.isconfig = False
                                    return
                        else:
                            return
                    except:
                        self.isconfig = False
                        return

            try:
                with open(self._cfg_file_path_cental,'w') as file:
                    pass
            except:
                self.isconfig = False
                return

    
    @staticmethod
    def get_config_static(nummername,section='Default',lokal = False):
        klass = ProjektConfig(nummername,lokal)
        if klass.isconfig:
            klass.get_config(section)
            return klass
        else:
            return None

    
    def get_value(self,option):
        if self.config.has_option(self.section, option):
            try:
                value = self.config.get(self.section, option)
                try:
                    try:
                        return json.loads(value)  #pylint: disable=W0123
                    except Exception:
                        # try fix legacy formats
                        # cleanup python style true, false values
                        if value == "True":
                            value = json.dumps(True)
                        elif value == "False":
                            value = json.dumps(False)
                        # cleanup string representations
                        value = value.replace('\'', '"').encode('string-escape')
                        # try parsing again
                        try:
                            return json.loads(value)  #pylint: disable=W0123
                        except Exception:
                            # if failed again then the value is a string
                            # but is not encapsulated in quotes
                            # e.g. option = C:\Users\Desktop
                            value = value.strip()
                            if not value.startswith('(') \
                                    or not value.startswith('[') \
                                    or not value.startswith('{'):
                                value = "\"%s\"" % value
                            return json.loads(value)  #pylint: disable=W0123
                except Exception:
                    return value
            except:pass
        else:
            self.config.set(self.section,option,'')
            return ''
    
    def set_value(self,option,value):
        self.config.set(self.section,option,json.dumps(value,separators=(',', ':'),ensure_ascii=False))

    def get_config(self,section = 'Default'):
        if self.lokal:
            self.section = self.nummername+section
        else:
            self.section = section
        if self.config.has_section(self.section):
            return
        else:
            self.config.add_section(self.section)
    
    def save_config(self):
        with open(self._cfg_file_path,'w') as config_file:
            self.config.write(config_file) 