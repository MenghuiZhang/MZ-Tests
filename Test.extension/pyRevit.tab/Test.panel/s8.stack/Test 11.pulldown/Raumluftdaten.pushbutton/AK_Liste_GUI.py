from pyrevit import forms, script
import os
from System.Windows.Forms import OpenFileDialog,DialogResult, SaveFileDialog
import xlsxwriter
XAML_FILES_DIR = os.path.dirname(__file__)


class Excelerstellen(forms.WPFWindow):
    def __init__(self, exceladresse=None):
        forms.WPFWindow.__init__(self, os.path.join(XAML_FILES_DIR, 'AK_Liste_GUI.xaml'))
        self.Adresse.Text = exceladresse

    def Speicherort(self, sender, args):
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

    def Export(self, sender, args):
        self.Close()

    def Abbrechen(self, sender, args):
        self.Close()
        script.exit()