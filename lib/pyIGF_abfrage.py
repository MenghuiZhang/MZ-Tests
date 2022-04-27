# coding: utf8
from pyrevit import forms
from System.Windows import Visibility
import os
from System.Windows.Forms import OpenFileDialog,DialogResult

XAML_FILES_DIR = os.path.dirname(__file__)

class abfrage(forms.WPFWindow):
    def __init__(self,title = None,info =None,ja = False,anmerkung = None,nein = True,ja_text= None,nein_text=None):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, 'abfrage_GUI.xaml'))
        if title:
            self.Title = title
        if not info:
            self.info.Text = 'Keine Information '
        else:
            self.info.Text = info
        if anmerkung:
            self.anmerkung.Text = anmerkung

        if ja:
            self.ja_button.Visibility = Visibility.Visible
        else:
            self.ja_button.Visibility = Visibility.Hidden
        if nein:
            self.nein_button.Visibility = Visibility.Visible
        else:
            self.nein_button.Visibility = Visibility.Hidden
        if ja_text:
            self.ja_button.Content = ja_text
        if nein_text:
            self.nein_button.Content = nein_text
        self.antwort = ''


    def ja(self,sender,args):
        self.antwort = self.ja_button.Content
        self.Close()
    def nein(self,sender,args):
        self.antwort = self.nein_button.Content
        self.Close()


class Suche(forms.WPFWindow):
    def __init__(self,adresse = None):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, 'excel_adresse.xaml'))
        self.Adresse.Text = adresse

    def durchsuchen(self,sender,args):
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "Excel"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            self.Adresse.Text = dialog.FileName

    def ok(self, sender, args):
        self.Close()
