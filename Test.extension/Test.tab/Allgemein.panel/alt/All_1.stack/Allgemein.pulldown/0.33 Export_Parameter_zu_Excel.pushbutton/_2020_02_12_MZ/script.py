# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import time
from pyrevit import script, forms
from rpw import *
import System


start = time.time()

__title__ = "9.33 Export_Projekt_Shared_Para_zu_Excel"
__doc__ = """die Informationen der Projekt- und SharedParameter zu Excel Exportieren"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()

exapp = Excel.ApplicationClass()

def DeiziplinErmitteln(inParaTypName,inParaTyp):
    commen = ['Text','Ganzzahl','Zahl','Länge','Fläche','Volumen','Winkel','Neigung','Währung','Massendichte',
              'Zeit','Geschwindigkeit','URL','Material','Bild','Ja/Nein','Mehrzeiliger Text']
    Energie = ['Energie','Wärmedurchgangskoeffizient','Thermischer Widerstand','Thermisch wirksame Masse',
               'Wärmeleitfähigkeit','Spezifische Wärme','Spezifische Verdunstungswärme','Permeabilität']
    outDisziplin = None
    if inParaTypName in commen:
        outDisziplin = 'Allgemein'
    elif inParaTypName in Energie:
        outDisziplin = 'Energie'
    else:
        outDisziplin = 'Tragwerk'

    if inParaTyp[:3] == 'HVA':
        outDisziplin = 'Lüftung'
    elif inParaTyp[:3] == 'Pip':
        outDisziplin = 'Rohre'
    elif inParaTyp[:3] == 'Ele':
        outDisziplin = 'Elektro'
    else:
        pass
    if inParaTypName == 'HVACEnergy':
        outDisziplin = 'Energie'

    return outDisziplin

def AktuellSharedPara(inputfile):
    aktuellSharedPara = []
    groupListe = []
    if inputfile.Groups:
        for dg in inputfile.Groups:
            definitionName = dg.Name
            groupListe.append(definitionName)
            definitionInfo = []
            if dg.Definitions:
                for d in dg.Definitions:
                    name = d.Name
                    owner = d.OwnerGroup.Name
                    GUID = d.GUID.ToString()
                    TYPE = d.ParameterType.ToString()
                    type = DB.LabelUtils.GetLabelFor(d.ParameterType)
                    dis = DeiziplinErmitteln(type,TYPE)
                    definitionInfo.append([name,GUID,dis,type])
            definitionInfo.sort()
            aktuellSharedPara.append([definitionName,definitionInfo])
    aktuellSharedPara.sort()

    return aktuellSharedPara,groupListe

def AktuellProjektPara():
    map = doc.ParameterBindings
    dep = map.ForwardIterator()
    Liste = []
    while(dep.MoveNext()):
        definition = dep.Key
        Name = definition.Name
        Paratyp = definition.ParameterType.ToString()
        TypName = DB.LabelUtils.GetLabelFor(definition.ParameterType)
        dis = DeiziplinErmitteln(TypName,Paratyp)
        Group = DB.LabelUtils.GetLabelFor(definition.ParameterGroup)
        cateName = ''
        Binding = dep.Current
        Type = Binding.GetType().ToString()
        typOrex = None
        if Type == 'Autodesk.Revit.DB.InstanceBinding':
            typOrex = 'Exemplar'
        else:
            typOrex = 'Type'
        if Binding:
            cates = Binding.Categories
            for cate in cates:
                cateName = cate.Name + ',' + cateName
        Liste.append([Name,dis,TypName,Group,typOrex,cateName])

    Liste.sort()
    ProjektParaListe = [i[0] for i in Liste]
    return Liste,ProjektParaListe


def GroupPara(SharedListe,ProjektListe,ProjektParaListe):
    OutListe_Shared = []
    OutListe_Projekt = []
    for i in SharedListe:
        Groupname = i[0]
        groupSha = []
        groupPro = []
        for item in i[1]:
            if item[0] in ProjektParaListe:
                for j in ProjektListe:
                    if item[0] == j[0]:
                        groupPro.append([item[0],item[1],item[2],item[3],j[3],j[4],j[5]])
                        break
            else:
                groupSha.append(item)
        groupSha.sort()
        groupPro.sort()
        OutListe_Shared.append([i[0],groupSha])
        OutListe_Projekt.append([i[0],groupPro])
    return OutListe_Shared,OutListe_Projekt


def ExcelLesen(inExcelPath):
    outGroupListe = []
    outGroupPara = []
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        groupParaListe = []
        groupName = sheet.Name
        outGroupListe.append(groupName)
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            definitionName = sheet.UsedRange.Cells[row, 3].Value2
            if definitionName:
                groupParaListe.append(definitionName)
        outGroupPara.append([groupName,groupParaListe])

    book.Save()
    book.Close()
    return outGroupListe,outGroupPara

def GroupErstellen(inExcelPath,inSharedGroup,inExcelGroup):
    columns = ['ProjektParameter[Ja/Nein]','Prüfen','ParameterName','GUID',
               'Disziplin','Parametertyp','ParameterGroup','Kategorien',
               'Typ oder Exemplar','Medien','Parameter Typ','Gruppe',
               'Hilfe','Berechnungsformel','Anmerkung']
    book = exapp.Workbooks.Open(inExcelPath)
    for i in inSharedGroup:
        if not i in inExcelGroup:
            sheet = book.Worksheets.Add()
            sheet.Name = i
            for col in range(1,16):
                sheet.Cells[1,col] = columns[col-1]
            logger.info('Register {} wird erstellt'.format(i))

    book.Save()
    book.Close()

def ExportParaZuExcel(inExcelPath,inGroupParaExcel,inGroupParaShared,inGroupParaProjekt):
    # prüfen, ob ein Parameter ein Projekt- oder Sharedparameter ist
    if forms.alert('nur prüfen, ob ein Parameter ein Projekt- oder Sharedparameter ist?', ok=False, yes=True, no=True):
        book = exapp.Workbooks.Open(inExcelPath)
        for sheet in book.Worksheets:
            sheetName = sheet.Name
            groupShared = []
            groupProjekt = []
            for i in inGroupParaShared:
                if sheetName == i[0]:
                    groupShared = [j[0] for j in i[1]]
                    break
            for i in inGroupParaProjekt:
                if sheetName == i[0]:
                    groupProjekt = [j[0] for j in i[1]]
                    break
            title = '{value}/{max_value} ' + sheetName
            if not sheetName in ['Hinweis', 'Revit (Original)']:
                with forms.ProgressBar(title=title, cancellable=True, step=5) as pb:
                    n = 0
                    rows = sheet.UsedRange.Rows.Count
                    for row in range(2, rows + 1):
                        if pb.cancelled:
                            script.exit()
                        n += 1
                        pb.update_progress(n, rows - 1)
                        if sheet.Cells[row, 3].Value2 in groupProjekt:
                            sheet.Cells[row, 2] = 'ProjektParameter'
                        elif sheet.Cells[row, 3].Value2 in groupShared:
                            sheet.Cells[row, 2] = 'SharedParameter'
                        else:
                            sheet.Cells[row, 2] = None
        book.Save()
        book.Close()
    # ALle Informationen in Excel exportieren
    if forms.alert('Parameter in Excel exportieren', ok=False, yes=True, no=True):
        book = exapp.Workbooks.Open(inExcelPath)
        for sheet in book.Worksheets:
            sheetName = sheet.Name
            groupExcel = []
            groupShared = []
            groupProjekt = []
            for i in inGroupParaExcel:
                if sheetName == i[0]:
                    groupExcel = i[1]
            for j in inGroupParaShared:
                if sheetName == j[0]:
                    groupShared = j[1]
            for k in inGroupParaProjekt:
                if sheetName == k[0]:
                    groupProjekt = k[1]
            rows = sheet.UsedRange.Rows.Count
            allrow = [i for i in range(2, rows + 1)]
            usedRow = []
            logger.info('Group {}, Anzahl {}'.format(sheetName,rows))
            if any(groupProjekt):
                title1 = '{value}/{max_value} ProjektParameter in Group ' + sheetName
                with forms.ProgressBar(title=title1, cancellable=True, step=2) as pb:
                    n = 0

                    for item1 in groupProjekt:
                        if pb.cancelled:
                            exapp.Quit()
                            script.exit()
                        n += 1
                        pb.update_progress(n, len(groupProjekt))
                        rows = sheet.UsedRange.Rows.Count
                        if not item1[0] in groupExcel:
                            sheet.Cells[rows + 1,2] = 'ProjektParameter'
                            sheet.Cells[rows + 1,3] = item1[0]
                            sheet.Cells[rows + 1,4] = item1[1]
                            sheet.Cells[rows + 1,5] = item1[2]
                            sheet.Cells[rows + 1,6] = item1[3]
                            sheet.Cells[rows + 1,7] = item1[4]
                            sheet.Cells[rows + 1,8] = item1[6]
                            sheet.Cells[rows + 1,9] = item1[5]
                            sheet.Cells[rows + 1,12] = sheetName
                        else:
                            notUsedRow = list(set(allrow)-set(usedRow))
                            for row in notUsedRow:
                                if item1[0] == sheet.UsedRange.Cells[row, 3].Value2:
                                    sheet.UsedRange.Cells[row,2] = 'ProjektParameter'
                                    sheet.UsedRange.Cells[row,4] = item1[1]
                                    sheet.UsedRange.Cells[row,5] = item1[2]
                                    sheet.UsedRange.Cells[row,6] = item1[3]
                                    sheet.UsedRange.Cells[row,7] = item1[4]
                                    sheet.UsedRange.Cells[row,8] = item1[6]
                                    sheet.UsedRange.Cells[row,9] = item1[5]
                                    sheet.UsedRange.Cells[row,12] = sheetName
                                    usedRow.append(row)
                                    break
            if any(groupShared):
                title1 = '{value}/{max_value} SharedParameter in Group ' + sheetName
                with forms.ProgressBar(title=title1, cancellable=True, step=2) as pb:
                    n = 0
                    for item in groupShared:
                        if pb.cancelled:
                            exapp.Quit()
                            script.exit()
                        n += 1
                        pb.update_progress(n, len(groupShared))
                        rows = sheet.UsedRange.Rows.Count
                        if not item[0] in groupExcel:
                            sheet.Cells[rows + 1,2] = 'SharedParameter'
                            sheet.Cells[rows + 1,3] = item[0]
                            sheet.Cells[rows + 1,4] = item[1]
                            sheet.Cells[rows + 1,5] = item[2]
                            sheet.Cells[rows + 1,6] = item[3]
                            sheet.Cells[rows + 1,12] = sheetName
                        else:
                            notUsedRow = list(set(allrow)-set(usedRow))
                            for row in notUsedRow:
                                if item[0] == sheet.UsedRange.Cells[row, 3].Value2:
                                    sheet.UsedRange.Cells[row,2] = 'SharedParameter'
                                    sheet.UsedRange.Cells[row,4] = item[1]
                                    sheet.UsedRange.Cells[row,5] = item[2]
                                    sheet.UsedRange.Cells[row,6] = item[3]
                                    sheet.UsedRange.Cells[row,12] = sheetName
                                    usedRow.append(row)
                                    break

        book.Save()
        book.Close()

sharedParaInGroup,sharedParaGroup = AktuellSharedPara(file)
projektPara,projektParaName = AktuellProjektPara()
groupSharedPara,groupProjektPara = GroupPara(sharedParaInGroup,projektPara,projektParaName)
excelPath = ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
defiGroupExcel,groupParaExcel = ExcelLesen(excelPath)
GroupErstellen(excelPath,sharedParaGroup,defiGroupExcel)
ExportParaZuExcel(excelPath,groupParaExcel,groupSharedPara,groupProjektPara)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
