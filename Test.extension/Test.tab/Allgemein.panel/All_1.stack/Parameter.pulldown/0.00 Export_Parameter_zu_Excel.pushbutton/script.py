# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_forms import ExcelSuche
from IGF_libKopie import AllProjectParams_2020,AllSharedParams_2020
from IGF_libKopie import AllProjectParams_2022,AllSharedParams_2022
from pyrevit import script, forms
from rpw import revit,DB
import os
from config import ProjektConfig

import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal


__title__ = "0.00 Export Parameter zu Excel"
__doc__ = """exportiert Projekt- und SharedParameter zu Excel 


[2021.10.08]
Version: 2.0
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

revitversion = app.VersionNumber

projectinfo = doc.ProjectInformation.Number + ' - '+ doc.ProjectInformation.Name

config = ProjektConfig.get_config_static(projectinfo,'Parametererstellung',True)

adresse = 'Excel Adresse'

try:
    adresse = config.get_value('adresse')
    if not os.path.exists(config.get_value('adresse')):
        config.set_value('adresse', '')
        adresse = "Excel Adresse"
except:
    pass

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.ShowDialog()

try:
    config.set_value('adresse', ExcelWPF.Adresse.Text)
    config.save_config()
except:
    logger.error('kein Excel gegeben')
    script.exit()

excelPath = config.get_value('adresse')

exapp = Excel.ApplicationClass()

def ExcelLesen(inExcelPath):
    outGroupPara = {}
    outrow = {}
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        groupName = sheet.Name
        rows = sheet.UsedRange.Rows.Count
        mapdaten = []
        for row in range(2, rows + 1):
            rowdaten = []
            if not sheet.Cells[row, 3].Value2:
                continue
            for col in [3,4,5,6,7,8,9,12,13]:
                daten = sheet.Cells[row, col].Value2
                rowdaten.append(daten)
            mapdaten.append(rowdaten)
            outrow[sheet.Cells[row, 3].Value2] = row
        outGroupPara[groupName] = mapdaten

    book.Save()
    book.Close()

    return outGroupPara,outrow


def SheetErstellen(inExcelPath,Projektgroup):
    columns = ['anlegen','angelegt[Ja/Nein]','ParameterName','GUID',
               'Disziplin','Parametertyp','ParameterGroup','Kategorien',
               'Typ oder Exemplar','Medien','Parameter Typ','Gruppe',
               'Hilfe','Berechnungsformel','Anmerkung']
    book = exapp.Workbooks.Open(inExcelPath)
    inExcelGroup = [sheet.Name for sheet in book.Worksheets]
    for i in Projektgroup:
        if len(i) > 31:
            gr = i[:31]
        else:
            gr = i
        if not gr in inExcelGroup:
            sheet = book.Worksheets.Add()
            sheet.Name = gr
            for col in range(1,16):
                sheet.Cells[1,col] = columns[col-1]
            logger.info('Register {} wird erstellt'.format(i))

    book.Save()
    book.Close()

def ExportParaZuExcel(inExcelPath,inGroupParaExcel,ingroupParameterdict):
    if forms.alert('Parameter in Excel exportieren', ok=False, yes=True, no=True):
        book = exapp.Workbooks.Open(inExcelPath)
        for sheetname in ingroupParameterdict.Keys:
            groupforschreiben = sheetname
            if len(sheetname) > 31:
                temp = sheetname[:31]
                sheetname = temp
            sheet = book.Worksheets[sheetname]
            shared = ingroupParameterdict[groupforschreiben][0]
            projek = ingroupParameterdict[groupforschreiben][1]
           
            print(30*'-')
            print(sheetname)
            print(30*'-')
        

            if any(shared):
                keys = shared.keys()
                if any(projek):
                    pro_keys = projek.keys()
                else:
                    pro_keys = []
                title1 = '{value}/{max_value} SharedParameter in Group ' + sheetname
                with forms.ProgressBar(title=title1, cancellable=True, step=2) as pb:
                    for n, item in enumerate(keys):
                        if pb.cancelled:
                            exapp.Quit()
                            script.exit()
                        pb.update_progress(n, len(keys))
                        rows = sheet.UsedRange.Rows.Count
                        if item in pro_keys:
                            continue
                        
                        if not item in inGroupParaExcel.Keys:
                            sheet.Cells[rows+1, 1] = ''
                            sheet.Cells[rows + 1,2] = 'SharedParameter'
                            sheet.Cells[rows + 1,3] = item
                            sheet.Cells[rows + 1,4] = shared[item][0]
                            sheet.Cells[rows + 1,5] = shared[item][1]
                            sheet.Cells[rows + 1,6] = shared[item][2]
                            sheet.Cells[rows + 1,12] = groupforschreiben
                            sheet.Cells[rows + 1,13] = shared[item][3]
                            logger.info('Shared Parameter {} wird erstellt'.format(item))
                        else:
                            row = inGroupParaExcel[item]
                            sheet.Cells[row, 1] = ''
                            sheet.Cells[row,2] = 'SharedParameter'
                            sheet.Cells[row,4] = shared[item][0]
                            sheet.Cells[row,5] = shared[item][1]
                            sheet.Cells[row,6] = shared[item][2]
                            sheet.Cells[row,12] = groupforschreiben
                            sheet.Cells[row,13] = shared[item][3]
                            logger.info('Shared Parameter {} wird aktualisiert'.format(item))
            if any(projek):
                keys = projek.keys()
                title1 = '{value}/{max_value} ProjektParameter in Group ' + sheetname
                with forms.ProgressBar(title=title1, cancellable=True, step=2) as pb:
                    for n, item1 in enumerate(keys):
                        if pb.cancelled:
                            exapp.Quit()
                            script.exit()
                        pb.update_progress(n, len(keys))
                        rows = sheet.UsedRange.Rows.Count
                        if sheetname == 'Keine Gruppe':
                            rows = n + 1
                        if not item1 in inGroupParaExcel.Keys:
                            sheet.Cells[rows +1, 1] = ''
                            sheet.Cells[rows + 1,2] = 'ProjektParameter'
                            sheet.Cells[rows + 1,3] = item1
                            sheet.Cells[rows + 1,4] = projek[item1][0]
                            sheet.Cells[rows + 1,5] = projek[item1][1]
                            sheet.Cells[rows + 1,6] = projek[item1][2]
                            sheet.Cells[rows + 1,7] = projek[item1][3]
                            sheet.Cells[rows + 1,8] = projek[item1][4]
                            sheet.Cells[rows + 1,9] = projek[item1][5]
                            sheet.Cells[rows + 1,12] = projek[item1][6]
                            sheet.Cells[rows + 1,13] = projek[item1][7]
                            logger.info('Projektparameter {} wird erstellt'.format(item1))
                        else:
                            row = inGroupParaExcel[item1]
                            sheet.Cells[row, 1] = ''
                            sheet.Cells[row,2] = 'ProjektParameter'
                            sheet.Cells[row,4] = projek[item1][0]
                            sheet.Cells[row,5] = projek[item1][1]
                            sheet.Cells[row,6] = projek[item1][2]
                            sheet.Cells[row,7] = projek[item1][3]
                            sheet.Cells[row,8] = projek[item1][4]
                            sheet.Cells[row,9] = projek[item1][5]
                            sheet.Cells[row,12] = projek[item1][6]
                            sheet.Cells[row,13] = projek[item1][7]
                            logger.info('Projektparameter {} wird aktualisiert'.format(item1))

        book.Save()
        book.Close()



def main():
    # Excel lesen
    Daten_Excel = None
    try:
        Daten_Excel,Daten_row = ExcelLesen(excelPath)
    except Exception as e:
        logger.error(e)
        script.exit()


    # Parameter ermitteln
    if revitversion == '2020':
        sharedparams,sharedparam_defis = AllSharedParams_2020()
        projectparams,projeparam_defis = AllProjectParams_2020()

    elif revitversion == '2022':
        sharedparams,sharedparam_defis = AllSharedParams_2022()
        projectparams,projeparam_defis = AllProjectParams_2022()

    else:
        script.exit()

    # Daten bearbeiten

    if Daten_Excel:
        if 'Keine Gruppe' in Daten_Excel.keys():
            keinegroup_excel = Daten_Excel['Keine Gruppe']
        else:
            keinegroup_excel = []
        Params_excel = {}
        for key in Daten_Excel.keys():
            for param in Daten_Excel[key]:
                Params_excel[param[0]] = param[7]
    else:
        keinegroup_excel = []
        Params_excel = {}
    
    Param_dict_for_export = {'Keine Gruppe': [{},{}]}
    for dg in sharedparams.keys():
        Param_dict_for_export[dg] = [sharedparams[dg],{}]

    for guid in projectparams.keys():
        if not projectparams[guid][6]:
            Param_dict_for_export['Keine Gruppe'][1][projectparams[guid][0]] = [guid,projectparams[guid][1],projectparams[guid][2],projectparams[guid][3],projectparams[guid][5],projectparams[guid][4],projectparams[guid][6],projectparams[guid][7]]
        elif projectparams[guid][6] in Param_dict_for_export.keys():
            Param_dict_for_export[projectparams[guid][6]][1][projectparams[guid][0]] = [guid,projectparams[guid][1],projectparams[guid][2],projectparams[guid][3],projectparams[guid][5],projectparams[guid][4],projectparams[guid][6],projectparams[guid][7]]

    keinegroup_project = Param_dict_for_export['Keine Gruppe'][1]
    mitgroup_project = []
    for key in Param_dict_for_export.keys():
        for el in Param_dict_for_export[key][0].keys():
            mitgroup_project.append(el)
    for el_excel in keinegroup_excel:
        if not el_excel[0] in keinegroup_project.keys():
            if not el_excel[0] in mitgroup_project:
                Param_dict_for_export['Keine Gruppe'][1][el_excel[0]] = [el_excel[1],el_excel[2],el_excel[3],el_excel[4],el_excel[5],el_excel[6],el_excel[7],el_excel[8]]

    keinegroup_project_neu = {}
    for el_proj in keinegroup_project.keys():
        if el_proj in Params_excel.keys():
            if Params_excel[el_proj] != 'Keine Gruppe':
                if not Params_excel[el_proj] in Param_dict_for_export.keys():
                    Param_dict_for_export[Params_excel[el_proj]] = [{},{}]
                keinegroup_project[el_proj][7] = Params_excel[el_proj]
                Param_dict_for_export[Params_excel[el_proj]][1][el_proj] = keinegroup_project[el_proj]
            else:
                keinegroup_project_neu[el_proj] = keinegroup_project[el_proj]
        
    Param_dict_for_export['Keine Gruppe'][1] = keinegroup_project_neu
    SheetErstellen(excelPath,Param_dict_for_export.keys())

    ExportParaZuExcel(excelPath,Daten_row,Param_dict_for_export)


main()