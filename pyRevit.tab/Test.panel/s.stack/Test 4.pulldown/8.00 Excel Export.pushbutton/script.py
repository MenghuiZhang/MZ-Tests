# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit import revit, DB
from pyrevit import script
from System.Windows.Forms import SaveFileDialog,DialogResult
import xlsxwriter

__title__ = "8.00 BIM-ID Export"
__doc__ = """exportiertExcel BIM ID"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass


logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc


systeme_luft = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
systeme_rohr = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()


system_luft = {}
system_rohr = {}



for el in systeme_luft:
    systyp = el.GetTypeId().ToString()
    if not systyp in system_luft.Keys:
        system_luft[systyp] = el.Id

for el in systeme_rohr:
    systyp = el.GetTypeId().ToString()
    if not systyp in system_rohr.Keys:
        system_rohr[systyp] = el.Id


def DatenErmitteln(Sys_coll):
    _dict = {}
    for id in Sys_coll.Keys:
        sys = doc.GetElement(Sys_coll[id])
        el = doc.GetElement(DB.ElementId(int(id)))
        sys_Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        KG = el.LookupParameter('IGF_X_Kostengruppe').AsValueString()
        Gewerke = el.LookupParameter('IGF_X_Gewerkkürzel').AsString()
        KN_01 = el.LookupParameter('IGF_X_Kennnummer_1').AsValueString()
        KN_02 = el.LookupParameter('IGF_X_Kennnummer_2').AsValueString()
        ID = el.LookupParameter('IGF_X_BIM-ID').AsString()

        bereich = None
        elements = None
        try:
            elements = sys.PipingNetwork
        except:
            elements = sys.DuctNetwork
        for ele in elements:
            try:
                bereich = ele.LookupParameter('Bearbeitungsbereich').AsValueString()
            except:
                pass
            if bereich:
                break
        _dict[sys_Name] = [Gewerke,KG,KN_01,KN_02,ID,bereich]
    return _dict


def excel_export(path,data_luft,data_rohr):
    workbook = xlsxwriter.Workbook(path)
    cell_format = workbook.add_format()
    cell_format.set_bold(True)
    cell_format1 = workbook.add_format()
    cell_format1.set_italic(True)
    worksheet = workbook.add_worksheet('Luft')
    worksheet.write(0, 0, 'Systemname', cell_format)
    worksheet.write(0, 1, 'Gewerkkürzel', cell_format)
    worksheet.write(0, 2, 'Kostengruppen', cell_format)
    worksheet.write(0, 3, 'Kennnummer_1', cell_format)
    worksheet.write(0, 4, 'Kennnummer_2', cell_format)
    worksheet.write(0, 5, 'BIM-ID', cell_format)
    worksheet.write(0, 6, 'Bearbeitungsbereich', cell_format)
    worksheet.write(1, 0, 'Systemtyp', cell_format1)
    worksheet.write(1, 1, 'IGF_X_Gewerkkürzel', cell_format1)
    worksheet.write(1, 2, 'IGF_X_Kostengruppe', cell_format1)
    worksheet.write(1, 3, 'IGF_X_Kennnummer_1', cell_format1)
    worksheet.write(1, 4, 'IGF_X_Kennnummer_2', cell_format1)
    worksheet.write(1, 5, 'IGF_X_BIM-ID', cell_format1)
    worksheet.write(1, 6, 'Bearbeitungsbereich', cell_format1)

    data_luft_keys = data_luft.keys()[:]
    data_luft_keys.sort()

    for row,key in enumerate(data_luft_keys):
        try:worksheet.write_string(row+2, 0, key)
        except:pass
        try:worksheet.write_string(row+2, 1, data_luft[key][0])
        except:pass
        try:worksheet.write_string(row+2, 2, data_luft[key][1])
        except:pass
        try:worksheet.write_string(row+2, 3, data_luft[key][2])
        except:pass
        try:worksheet.write_string(row+2, 4, data_luft[key][3])
        except:pass
        try:worksheet.write_string(row+2, 5, data_luft[key][4])
        except:pass
        try:worksheet.write_string(row+2, 6, data_luft[key][5])
        except:pass

    worksheet.autofilter(0, 0, len(data_luft_keys)-1, 5)

    worksheet = workbook.add_worksheet('Rohr')
    worksheet.write(0, 0, 'Systemname', cell_format)
    worksheet.write(0, 1, 'Gewerkkürzel', cell_format)
    worksheet.write(0, 2, 'Kostengruppen', cell_format)
    worksheet.write(0, 3, 'Kennnummer_1', cell_format)
    worksheet.write(0, 4, 'Kennnummer_2', cell_format)
    worksheet.write(0, 5, 'BIM-ID', cell_format)
    worksheet.write(0, 6, 'Bearbeitungsbereich', cell_format)
    worksheet.write(1, 0, 'Systemtyp', cell_format1)
    worksheet.write(1, 1, 'IGF_X_Gewerkkürzel', cell_format1)
    worksheet.write(1, 2, 'IGF_X_Kostengruppe', cell_format1)
    worksheet.write(1, 3, 'IGF_X_Kennnummer_1', cell_format1)
    worksheet.write(1, 4, 'IGF_X_Kennnummer_2', cell_format1)
    worksheet.write(1, 5, 'IGF_X_BIM-ID', cell_format1)
    worksheet.write(1, 6, 'Bearbeitungsbereich', cell_format1)

    data_rohr_keys = data_rohr.keys()[:]
    data_rohr_keys.sort()

    for row,key in enumerate(data_rohr_keys):

        try:worksheet.write_string(row+2, 0, key)
        except:pass
        try:worksheet.write_string(row+2, 1, data_rohr[key][0])
        except:pass
        try:worksheet.write_string(row+2, 2, data_rohr[key][1])
        except:pass
        try:worksheet.write_string(row+2, 3, data_rohr[key][2])
        except:pass
        try:worksheet.write_string(row+2, 4, data_rohr[key][3])
        except:pass
        try:worksheet.write_string(row+2, 5, data_rohr[key][4])
        except:pass
        try:worksheet.write_string(row+2, 6, data_rohr[key][5])
        except:passs

    worksheet.autofilter(0, 0, len(data_rohr_keys)-1, 5)
    workbook.close()

def excel_erstellen():
    dialog = SaveFileDialog()
    dialog.Title = "Speichern unter"
    dialog.Filter = "Execl Dateien (*.xlsx)|*.xlsx"
    dialog.FilterIndex = 0
    dialog.RestoreDirectory = True

    if dialog.ShowDialog() == DialogResult.OK:
        workbook = xlsxwriter.Workbook(dialog.FileName)
        workbook.add_worksheet()
        workbook.close()

    return dialog.FileName

Luft = DatenErmitteln(system_luft)
Rohr = DatenErmitteln(system_rohr)
path = excel_erstellen()
excel_export(path,Luft,Rohr)
