# coding: utf8
from IGF_Forms import WPFWindow
import os
XAML_FILES_DIR = os.path.dirname(__file__)

class Template(WPFWindow):
    def __init__(self,methode):
        WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, '_template.xaml'),handle_esc=False)
        self.methode = methode
       
        