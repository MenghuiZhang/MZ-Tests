# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script,forms
from IGF_forms import Texteingeben
import clr
import os
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Forms import OpenFileDialog,DialogResult
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal

__title__ = "0.61 importiert ExportDWGOption"
__doc__ = """
importiert ExportDWGOption


[2021.12.07]
Version: 1.0
"""
__author__ = "Menghui Zhang"


try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

logger = script.get_logger()
output = script.get_output()

projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number
config = script.get_config('ExportDWGOption-Import' + projectinfo)


coll = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ExportDWGSettings))
liste_optionen = {}
for el in coll:
    name = el.Name
    liste_optionen[name] = el.Id

if coll.GetElementCount() == 0:
    logger.info('Kein ExportDWGOption vorhanden')
ids = coll.ToElementIds()
coll.Dispose()

class DWGOption(object):
    def __init__(self,name,elemid):
        self.name = name
        self.id = elemid
        self.checked = False

Liste_Optionen = ObservableCollection[DWGOption]()
for key in sorted(liste_optionen.keys()):
    Liste_Optionen.Add(DWGOption(key,liste_optionen[key]))

class Import(forms.WPFWindow):
    def __init__(self, xaml_file_name,liste_option):
        self.liste_option = liste_option
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.ListView.ItemsSource = liste_option
        self.read_config()

    def read_config(self):
        try:
            if os.path.exists(config.Adresse):
                self.Adresse.Text = config.Adresse
            else:
                self.Adresse.Text = config.Adresse = ""
        except:
            self.Adresse.Text = config.Adresse = ""

    def write_config(self):
        config.Adresse = self.Adresse.Text
        script.save_config()

    def durchsuchen(self,sender,args):
        dialog = OpenFileDialog()
        dialog.Title = "Excel"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            self.Adresse.Text = dialog.FileName
        else:
            dialog.FileName = self.Adresse.Text

        self.write_config()
    
    def ok(self,sender,args):
        if not self.Adresse.Text:
            UI.TaskDialog.Show(__title__,'Keine Vorlage ausgewählt')
            return
        for el in self.liste_option:
            if el.checked:
                self.Close()
                return
        UI.TaskDialog.Show(__title__,'Keine ExportOption ausgewählt')
            
    
    def abbrechen(self,sender,args):
        self.Close()
        script.exit()

    def erstellen(self,sender,args):
        nameeingabe = Texteingeben(label='Name:')
        nameeingabe.Title = 'ExportDWGOption'
        nameeingabe.ShowDialog()
        name = nameeingabe.text.Text
        if not name:
            UI.TaskDialog.Show(__title__,'Keine Name eingegeben')
            return
        if name in liste_optionen.keys():
            UI.TaskDialog.Show(__title__,'eigegebene Name veretis verwendet')
            return

        t = DB.Transaction(doc,'DWGOption erstellen')
        t.Start()
        elem = DB.ExportDWGSettings.Create(doc,name)
        t.Commit()
        t.Dispose()
        dwgoption = DWGOption(name,elem.Id)
        for el in self.liste_option:
            el.checked = False
        self.liste_option.Add(dwgoption)
        dwgoption.checked = True
        self.ListView.Items.Refresh()
    
    def copy(self,sender,args):
        template = None
        for el in self.liste_option:
            if el.checked:
                template = el
                break
        if not template:
            UI.TaskDialog.Show(__title__,'Keine Option ausgewählt')
            return

        nameeingabe = Texteingeben(label='Name:')
        nameeingabe.Title = 'ExportDWGOption'
        nameeingabe.ShowDialog()
        name = nameeingabe.text.Text
        if not name:
            UI.TaskDialog.Show(__title__,'Keine Name eingegeben')
            return
        
        if name in liste_optionen.keys():
            UI.TaskDialog.Show(__title__,'eigegebene Name veretis verwendet')
            return
        
        t = DB.Transaction(doc,'DWGOption duplizieren')
        t.Start()
        option = doc.GetElement(template.id)
        elem = DB.ExportDWGSettings.Create(doc,name,option.GetDWGExportOptions())
        t.Commit()
        t.Dispose()
        dwgoption = DWGOption(name,elem.Id)
        for el in self.liste_option:
            el.checked = False
        self.liste_option.Add(dwgoption)
        dwgoption.checked = True
        self.ListView.Items.Refresh()
    
optionen = Import('window.xaml',Liste_Optionen)
try:
    optionen.show_dialog()
except Exception as e:
    logger.error(e)
    optionen.Close()
    script.exit()


option_dwg = None
for el in Liste_Optionen:
    if el.checked:
        option_dwg = el
        break

excelPath = optionen.Adresse.Text
exapplication = ex.ApplicationClass()

book1 = exapplication.Workbooks.Open(excelPath)
systemtyp_dict = {}
for sheet in book1.Worksheets:
    with forms.ProgressBar(title="{value}/{max_value} Zeile --- Excel lesen",cancellable=True, step=10) as pb:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            if pb.cancelled:
                Marshal.FinalReleaseComObject(sheet)
                Marshal.FinalReleaseComObject(book1)
                exapplication.Quit()
                Marshal.FinalReleaseComObject(exapplication)
                script.exit()
            pb.update_progress(row-1, rows)
            name = sheet.Cells[row, 1].Value2
            layername = sheet.Cells[row, 2].Value2
            layerfarbe = sheet.Cells[row, 3].Value2
            systemtyp_dict[name] = [layername,layerfarbe]

book1.Save()
book1.Close()

t = DB.Transaction(doc,'Export Option')
t.Start()
optionen_temp = doc.GetElement(option_dwg.id).GetDWGExportOptions()
layertable = optionen_temp.GetExportLayerTable()
dict_neu = {} 
for item in layertable:
    layinfo = item.Value
    Categoryname = item.Key.CategoryName
    name = item.Key.SubCategoryName

    if name in systemtyp_dict.keys():
        try:
            layinfo.LayerName = systemtyp_dict[name][0]
        except:
            pass
        try:
            if systemtyp_dict[name][1]:
                layinfo.ColorNumber = int(systemtyp_dict[name][1])
        except:
            pass
        try:
            layinfo.CutLayerName = systemtyp_dict[name][0]
        except:
            pass
        try:
            if systemtyp_dict[name][3]:
                layinfo.CutColorNumber = int(systemtyp_dict[name][1])
        except:
            pass
        
        
        dict_neu[item.Key] = item.Value
    
for el in dict_neu.keys():
    layertable.Remove(el)
    layertable.Add(el,dict_neu[el])

optionen_temp.SetExportLayerTable(layertable)
doc.GetElement(option_dwg.id).SetDWGExportOptions(optionen_temp)

t.Commit()