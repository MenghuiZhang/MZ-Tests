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
    for dg in inputfile.Groups:
        for d in dg.Definitions:
            name = d.Name
            owner = d.OwnerGroup.Name
            GUID = d.GUID.ToString()
            TYPE = d.ParameterType.ToString()
            type = DB.LabelUtils.GetLabelFor(d.ParameterType)
            dis = DeiziplinErmitteln(type,TYPE)
            definition = [owner,name,GUID,dis,type]
            aktuellSharedPara.append(definition)
    aktuellSharedPara.sort()
    if any(aktuellSharedPara):
        output.print_table(
            table_data=aktuellSharedPara,
            title="Aktuelle SharedParameters",
            columns=['Group', 'Name', 'GUID','Disziplin', 'ParameterType']
        )

    return aktuellSharedPara
def GroupSharedPara(inSharedPara,inGroup):
    outListe = []
    for i in inGroup:
        outListe.append([i,[]])
    for i in inSharedPara:
        for j in outListe:
            if i[0] == j[0]:
                j[1].append(i[1:])
    return outListe


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
    if any(Liste):
        output.print_table(
            table_data=Liste,
            title="Aktuelle Projectparameter",
            columns=['ParameterName', 'Disziplin','ParameterTyp','ParameterGroup','Typ or Exemplar','Kategorie']
        )

    return Liste

def AktuSharProPara(SharedListe,ProjektListe):
    OutListe = []
    for i in SharedListe:
        for j in ProjektListe:
            if i[1] == j[0]:
                OutListe.append([i[0],i[1],i[2],i[3],i[4],j[3],j[4],j[5]])
                break
    if any(OutListe):
        output.print_table(
            table_data=OutListe,
            title="ProjektParameter aus SharedParameter",
            columns=['DefinitionGroup', 'ParameterName','GUID','Disziplin','ParameterType','ParameterGroup','Typ or Exemplar','Kategorie'])
    return OutListe

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

def ExportSharedParaZuExcel(inExcelPath,inGroupParaExcel,inGroupParaShared):
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        sheetName = sheet.Name
        groupExcel = []
        groupShared = []
        for item in inGroupParaExcel:
            if sheetName == item[0]:
                groupExcel = item[1]
        for item in inGroupParaShared:
            if sheetName == item[0]:
                groupShared = item[1]
        if not any(groupShared):
            continue
        title = '{value}/{max_value} SharedParameter in Group ' + sheetName
        with forms.ProgressBar(title=title, cancellable=True, step=2) as pb:
            n = 0
            usedRow = []
            rows = sheet.UsedRange.Rows.Count
            allrow = [i for i in range(2, rows + 1)]
            print(rows)
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

def ExportProjektParaZuExcel(inExcelPath,inGroupParaExcel,inGroupParaProjekt):
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        sheetName = sheet.Name
        groupExcel = []
        groupProjekt = []
        for item in inGroupParaExcel:
            if sheetName == item[0]:
                groupExcel = item[1]
        for item in inGroupParaProjekt:
            if sheetName == item[0]:
                groupProjekt = item[1]
        if not any(groupProjekt):
            continue
        rows = sheet.UsedRange.Rows.Count
        allrow = [i for i in range(2, rows + 1)]
        print(rows)
        title = '{value}/{max_value} ProjektParameter in Group ' + sheetName
        with forms.ProgressBar(title=title, cancellable=True, step=2) as pb:
            n = 0
            usedRow = []
            for item in groupProjekt:
                if pb.cancelled:
                    script.exit()
                n += 1
                pb.update_progress(n, len(groupProjekt))
                notUsedRow = list(set(allrow)-set(usedRow))
                for row in notUsedRow:
                    if item[0] == sheet.UsedRange.Cells[row, 3].Value2:
                        sheet.UsedRange.Cells[row,2] = 'ProjektParameter'
                        sheet.UsedRange.Cells[row,7] = item[4]
                        sheet.UsedRange.Cells[row,8] = item[6]
                        sheet.UsedRange.Cells[row,9] = item[5]
                        usedRow.append(row)
                        break

    book.Save()
    book.Close()

aktuSharedPara = AktuellSharedPara(file)
aktuProjektPara = AktuellProjektPara()
aktuSharProPara = AktuSharProPara(aktuSharedPara,aktuProjektPara)

sharedDefinitionListe = [i[1] for i in aktuSharedPara]
proDefinitionListe = [i[0] for i in aktuProjektPara]
sharproDefinitionListe = [i[1] for i in aktuSharProPara]
defiGroupShared = [i[0] for i in aktuSharedPara]
defiGroupShared = set(defiGroupShared)
defiGroupShared = list(defiGroupShared)

excelPath = ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
defiGroupExcel,groupParaExcel = ExcelLesen(excelPath)
GroupErstellen(excelPath,defiGroupShared,defiGroupExcel)
groupSharedPara = GroupSharedPara(aktuSharedPara,defiGroupShared)
groupProjektPara = GroupSharedPara(aktuSharProPara,defiGroupShared)
ExportSharedParaZuExcel(excelPath,groupParaExcel,groupSharedPara)
ExportProjektParaZuExcel(excelPath,groupParaExcel,groupProjektPara)

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
