# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit,  DB
from pyrevit import script,forms
import xlsxwriter
from IGF_forms import Excelerstellen
import clr
import os

__title__ = "0.60 exportiert ExportDWGOption"
__doc__ = """
exportiert ExportDWGOption


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
config = script.get_config('ExportDWGOption-Export' + projectinfo)

adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

excel = Excelerstellen(exceladresse = adresse)
excel.show_dialog()

try:
    config.adresse = excel.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel gegeben')
    script.exit()

coll = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ExportDWGSettings))

if coll.GetElementCount() == 0:
    t_temp = DB.Transaction(doc)
    t_temp.Start('Option erstellen')
    DWG_option = DB.ExportDWGSettings.Create(doc,'Option temp')
    doc.Regenerate()
    t_temp.Commit()
else:
    t_temp = None
    dict_optionen = {}
    for el in coll:
        name = el.Name
        dict_optionen[name] = el
    option_selected = forms.SelectFromList.show(dict_optionen.keys(), button_name='Select')
    if not option_selected:
        logger.error('Keine Option ausgewählt')
        script.exit()
    DWG_option = dict_optionen[option_selected]
    
    

coll.Dispose()

class MEP_System:
    def __init__(self,name,pro_lay,pro_farbe):
        self.main = False
        self.name = name
        self.pro_lay = pro_lay
        self.pro_farbe = pro_farbe

dict_all = {}
ueberschrift = MEP_System('Name','Projekt_LayerName','Projekt_FarbeId')
layertable = DWG_option.GetDWGExportOptions().GetExportLayerTable()
for item in layertable:
    layinfo = item.Value
    Categoryname = item.Key.CategoryName
    name = item.Key.SubCategoryName
    Categorytype = layinfo.CategoryType.ToString()
    
    if not Categorytype in dict_all.keys():
        dict_all[Categorytype] = [ueberschrift]
    layername = layinfo.LayerName
    layerfarbeid = layinfo.ColorNumber
    if layerfarbeid == -1:
        layerfarbeid = ''

    mepsystem = MEP_System(name,layername,layerfarbeid)
    if not name:
        mepsystem.main = True
        mepsystem.name = Categoryname
    dict_all[Categorytype].append(mepsystem)
    
excel_path = excel.Adresse.Text
workbook = xlsxwriter.Workbook(excel_path)
cell_format = workbook.add_format()
cell_format.set_pattern(1)  # This is optional when using a solid fill.
cell_format.set_bg_color('green')

for key in dict_all.keys(): 
    worksheet = workbook.add_worksheet(key)
    liste = dict_all[key]

    for row in range(len(liste)):
        item = liste[row]
        if item.main:
            worksheet.write(row, 0, item.name,cell_format)
            worksheet.write(row, 1, item.pro_lay,cell_format)
            worksheet.write(row, 2, item.pro_farbe,cell_format)

        else:
            worksheet.write(row, 0, item.name)
            worksheet.write(row, 1, item.pro_lay)
            worksheet.write(row, 2, item.pro_farbe)



    worksheet.autofilter(0, 0, int(len(liste))-1, 2)
workbook.close()

if t_temp:
    t_temp.Start('temp. Option löschen')
    doc.Delete(DWG_option.Id)
    t_temp.Commit()
    t_temp.Dispose()
