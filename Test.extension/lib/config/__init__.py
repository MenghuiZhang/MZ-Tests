# coding: utf8
import ConfigParser
import os
import json
home_dir = os.path.expanduser("~") + '\\DC\\ACCDocs\\IGF\\IGF_pyrevit\\Projektdateien\\xx_Projektconfig'
appdata = os.getenv('APPDATA')

class ProjektConfig(object):
    def __init__(self,nummername = ''):
        self.config = ConfigParser.ConfigParser()
        self.nummername = nummername
        self.config_rvt = home_dir + '\\' + self.nummername
        self.condigdatei = appdata + '\\pyRevit'
        self.lokal = False
        self._cfg_file_path0 = self.condigdatei + '\\pyRevit_config.ini'
        self._cfg_file_path1 = self.config_rvt + '\\pyRevit_config.conf'
        self.isconfig = True
        self._cfg_file_path = self._cfg_file_path1

        if not self.isconfig:return
        
        self.config.read(self._cfg_file_path)
        self.section = ''
        
        
    
    @property
    def config_file(self):
        """Current config file path."""
        return self._cfg_file_path

    def config_ermitterl(self):
        if not os.path.exists(self._cfg_file_path1):
            self._cfg_file_path = self._cfg_file_path1
            if not os.path.exists(self.config_rvt):
                try:
                    os.makedirs(self.config_rvt)
                except:
                    try:
                        if not os.path.exists(self._cfg_file_path0):
                            self._cfg_file_path = self._cfg_file_path0
                            if not os.path.exists(self.condigdatei):
                                try:
                                    os.makedirs(self.condigdatei)
                                except:
                                    self.isconfig = False
                                    return
                            else:
                                try:
                                    with open(self._cfg_file_path0,'w') as file:
                                        pass
                                except:
                                    self.isconfig = False
                                    return
                    except:
                        self.isconfig = False
                        return

            try:
                with open(self._cfg_file_path1,'w') as file:
                    pass
            except:
                self.isconfig = False
                return

    
    @staticmethod
    def get_config_static(nummername,section='Default'):
        klass = ProjektConfig(nummername)
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