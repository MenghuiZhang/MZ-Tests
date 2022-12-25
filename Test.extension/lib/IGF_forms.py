# coding: utf8
from pyrevit import forms
from System.Windows import Visibility
import os
from System.Windows.Forms import OpenFileDialog,DialogResult, SaveFileDialog
import xlsxwriter
import sys
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
            self.ja_text = ja_text
        if nein_text:
            self.nein_button.Content = nein_text
            self.nein_text = nein_text
        self.antwort = ''

        self.ja_text = ja_text


    def ja(self,sender,args):
        self.antwort = self.ja_text
        self.Close()
    def nein(self,sender,args):
        self.antwort = self.nein_text
        self.Close()


class ExcelSuche(forms.WPFWindow):
    def __init__(self,exceladresse = None):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, 'excel_adresse.xaml'))
        self.Adresse.Text = exceladresse

    def durchsuchen(self,sender,args):
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "Excel"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            self.Adresse.Text = dialog.FileName
        else:
            dialog.FileName = self.Adresse.Text

    def ok(self, sender, args):
        self.Close()
    def abbrechen(self, sender, args):
        self.Close()
        sys.exit()
        

class Texteingeben(forms.WPFWindow):
    def __init__(self,text = None,label = None):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, 'Texteingabe.xaml'))
        self.text.Text = text
        self.label.Text = label

    def ok(self, sender, args):
        self.Close()

class Excelerstellen(forms.WPFWindow):
    def __init__(self,exceladresse = None):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, 'excel_adresse.xaml'))
        self.Adresse.Text = exceladresse
        self.Durch.Content = 'erstellen'

    def durchsuchen(self,sender,args):
        dialog = SaveFileDialog()
        
        dialog.Title = "Speichern unter"
        dialog.Filter = "Excel Dateien|*.xlsx"
        dialog.FilterIndex = 0
        dialog.RestoreDirectory = True
        if dialog.ShowDialog() == DialogResult.OK:
            workbook = xlsxwriter.Workbook(dialog.FileName)
            workbook.add_worksheet()
            workbook.close()
        else:
            dialog.FileName = self.Adresse.Text
        self.Adresse.Text = dialog.FileName

    def ok(self, sender, args):
        self.Close()