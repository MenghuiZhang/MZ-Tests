# coding: utf8
from IGF_log import getlog
import os
from eventhandler import ExternalEvent,ExternalEvenetListe
from IGF_Forms import WPFWindow
from System.Text.RegularExpressions import Regex

from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
from IGF_Funktionen._Parameter import wert_schreibenbase
from System.IO import FileInfo
from rpw import revit,DB

path = r"C:\Users\zhang\Desktop\Brandschutzklappe_MFC.xlsx"
# path1 = r"C:\Users\zhang\Desktop\MFC BSK UG _01.11.2024.xlsx"

ExcelPackage.LicenseContext = LicenseContext.NonCommercial
fs = FileInfo(path)
book = ExcelPackage(fs)

# fs1 = FileInfo(path1)
# book1 = ExcelPackage(fs1)

sheet = book.Workbook.Worksheets['BSK']

Liste = []
for n in range(721,809):
    _id = int(sheet.Cells[n,4].Value)
    Liste.append(_id)



__title__ = "0.21 Element-Suche (ElementId)"
__doc__ = """

[2023.03.21]
Version: 2.0
"""
__authors__ = "Menghui Zhang"

try:getlog(__title__)
except:pass

class Suche(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.regex = Regex("[^0-9]+")
        self.externaleventliste = ExternalEvenetListe()
        self.externaleventliste.GUI = self
        self.externalevent = ExternalEvent.Create(self.externaleventliste)  
        self.elemetid.ItemsSource = Liste
        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))

    def Anzeigen(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.anzeigen
        self.externalevent.Raise()

    def OK(self, sender, args):
        self.externaleventliste.ExcuteApp = self.externaleventliste.select
        self.externalevent.Raise()
    
    def Textinput(self, sender, args):
        try:args.Handled = self.regex2.IsMatch(args.Text)
        except:args.Handled = True

    def Abbrechen(self, sender, args):
        self.Close()

suche = Suche()
suche.Show()
 