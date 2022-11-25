# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit,  DB
from pyrevit import script
import xlsxwriter
from IGF_forms import ExcelSuche
import clr
import os

__title__ = "0.52 ExportDWGOption exportieren"
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
config = script.get_config('ExportDWGOption-Import' + projectinfo)

adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

excel = ExcelSuche(exceladresse = adresse)
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
    for el in coll:
        DWG_option = el
        break

coll.Dispose()

class MEP_System:
    def __init__(self,name,pro_lay,pro_farbe,sch_lay,sch_farbe):
        self.name = name
        self.pro_lay = pro_lay
        self.pro_farbe = pro_farbe
        self.sch_lay = sch_lay
        self.sch_farbe = sch_farbe


systemtyp_liste = []
ueberschrift = MEP_System('Systemtyp','Projekt_LayerName','Projekt_FarbeId','Schnitt_LayerName','Schnitt_FarbeId')
systemtyp_liste.append(ueberschrift)
layertable = DWG_option.GetDWGExportOptions().GetExportLayerTable()
for item in layertable:
    layinfo = item.Value
    Categoryname = item.Key.CategoryName
    if Categoryname == 'Systemtyp':
        if layinfo.LayerName != 'Systemtyp':
            name = item.Key.SubCategoryName
            layername = layinfo.LayerName
            layerfarbeid = layinfo.ColorNumber
            cutname = layinfo.CutLayerName
            cutfarbeid = layinfo.CutColorNumber
            if layerfarbeid == -1:
                layerfarbeid = ''
            if cutfarbeid == -1:
                cutfarbeid = ''
            mepsystem = MEP_System(name,layername,layerfarbeid,cutname,cutfarbeid)
            systemtyp_liste.append(mepsystem)


excel_path = excel.Adresse.Text
workbook = xlsxwriter.Workbook(excel_path)
worksheet = workbook.add_worksheet('ExportDWGOption')

for row in range(len(systemtyp_liste)):
    item = systemtyp_liste[row]
    worksheet.write(row, 0, item.name)
    worksheet.write(row, 1, item.pro_lay)
    worksheet.write(row, 2, item.pro_farbe)
    worksheet.write(row, 3, item.sch_lay)
    worksheet.write(row, 4, item.sch_farbe)

worksheet.autofilter(0, 0, int(len(systemtyp_liste))-1, 4)
workbook.close()

if t_temp:
    t_temp.Start('temp. Option löschen')
    doc.Delete(DWG_option.Id)
    t_temp.Commit()
    t_temp.Dispose()