# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import time
from pyrevit import script, forms
from rpw import revit,DB,ui
import System


__title__ = "9.33 Export_Projekt_Shared_Para_zu_Excel"
__doc__ = """die Informationen der Projekt- und SharedParameter zu Excel Exportieren"""
__author__ = "Menghui Zhang"
start = time.time()
logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()
exapp = Excel.ApplicationClass()
exapp.Visible = True
def AktuellSharedPara(inputfile):
    aktuellSharedPara = []
    for dg in inputfile.Groups:
        for d in dg.Definitions:
            name = d.Name
            owner = d.OwnerGroup.Name
            GUID = d.GUID
            TYPE = d.ParameterType.ToString()
            type = DB.LabelUtils.GetLabelFor(d.ParameterType)
            commen = ['Text','Ganzzahl','Zahl','Länge','Fläche','Volumen','Winkel','Neigung','Währung','Massendichte',
                      'Zeit','Geschwindigkeit','URL','Material','Bild','Ja/Nein','Mehrzeiliger Text']
            Energie = ['Energie','Wärmedurchgangskoeffizient','Thermischer Widerstand','Thermisch wirksame Masse',
                       'Wärmeleitfähigkeit','Spezifische Wärme','Spezifische Verdunstungswärme','Permeabilität']
            dis = None
            if type in commen:
                dis = 'Allgemein'
            elif type in Energie:
                dis = 'Energie'
            else:
                dis = 'Tragwerk'

            if TYPE[:3] == 'HVA':
                dis = 'Lüftung'
            elif TYPE[:3] == 'Pip':
                dis = 'Rohre'
            elif TYPE[:3] == 'Ele':
                dis = 'Elektro'
            else:
                pass
            if name == 'HVACEnergy':
                dis = 'Energie'

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

def AktuellProjektPara():
    map = doc.ParameterBindings
    dep = map.ForwardIterator()
    Liste = []
    while(dep.MoveNext()):
        definition = dep.Key
        Name = definition.Name
        Typ = DB.LabelUtils.GetLabelFor(definition.ParameterType)
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
        Liste.append([Name,Typ,cateName,typOrex,Group])
    Liste.sort()
    if any(Liste):
        output.print_table(
            table_data=Liste,
            title="Aktuelle Projectparameter",
            columns=['Name', 'Type','categories','Typ or Exemplar','Group']
        )
    return Liste

def AktuSharProPara(SharedListe,ProjektListe):
    OutListe = []
    for i in SharedListe:
        for j in ProjektListe:
            if i[1] == j[0]:
                OutListe.append([i[0],i[1],i[2],i[3],i,[4],j[4],j[2],j[3]])
    if any(OutListe):
        output.print_table(
            table_data=OutListe,
            title="ProjektParameter aus SharedParameter",
            columns=['Group', 'Name','GUID','Disziplin','ParameterType','ParameterGroup','Kategorie','Typ or Exemplar'])
    return OutListe


aktuSharedPara = AktuellSharedPara(file)
aktuProjektPara = AktuellProjektPara()
aktuSharProPara = AktuSharProPara(aktuSharedPara,aktuProjektPara)

sharedDefinitionListe = [i[1] for i in aktuSharedPara]
sharproDefinitionListe = [i[1] for i in aktuSharProPara]


excelPath = ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
book = exapp.Workbooks.Open(excelPath)

ExcelGroup = []
DefinitionListeInExcel = []
for sheet in book.Worksheets:
    Gname = sheet.Name
    ExcelGroup.append(Gname)
    title = '{value}/{max_value} ' + Gname
    with forms.ProgressBar(title=title, cancellable=True, step=5) as pb:
        n = 0
        if not Gname in ['Hinweis', 'Revit (Original)']:
            rows = sheet.UsedRange.Rows.Count
            for row in range(2, rows + 1):
                if pb.cancelled:
                    script.exit()
                n += 1
                pb.update_progress(n, rows-1)
                DefinitionListeInExcel.append(sheet.Cells[row, 3].Value2)
                if sheet.Cells[row, 3].Value2 in sharproDefinitionListe:
                    sheet.Cells[row, 2] = 'Beide'
                    for item in aktuSharProPara:
                        if item[0] == Gname:
                            if item[1] == sheet.Cells[row, 3].Value2:
                                sheet.Cells[row,4] = item[2]
                                sheet.Cells[row,5] = item[3]
                                sheet.Cells[row,6] = item[4]
                                sheet.Cells[row,7] = item[5]
                                sheet.Cells[row,8] = item[6]
                                sheet.Cells[row,9] = item[7]
                elif sheet.Cells[row, 3].Value2 in sharedDefinitionListe:
                    sheet.Cells[row, 2] = 'Shared'
                    for item in aktuSharedPara:
                        if item[0] == Gname:
                            if item[1] == sheet.Cells[row, 3].Value2:
                                sheet.Cells[row,4] = item[2]
                                sheet.Cells[row,5] = item[3]
                                sheet.Cells[row,6] = item[4]
                else:
                    sheet.Cells[row, 2] = None
with forms.ProgressBar(title='{value}/{max_value} Neuen SharedParameter zu Excel schreiben', cancellable=True, step=5) as pb:
    n = 0
    for Data in sharedDefinitionListe:
        if pb.cancelled:
            script.exit()
        n += 1
        pb.update_progress(n, len(sharedDefinitionListe))
        if not Data[1] in DefinitionListeInExcel:
            if Data[0] in ExcelGroup:
                sheet = book.Sheets["Sheet2"]
                rows = sheet.UsedRange.Rows.Count
                if rows >= 1:
                    sheet.Cells[rows + 1,2] = 'Shared'
                    sheet.Cells[rows + 1,3] = Data[1]
                    sheet.Cells[rows + 1,4] = Data[2]
                    sheet.Cells[rows + 1,5] = Data[3]
                    sheet.Cells[rows + 1,6] = Data[4]
                else:
                    sheet.Cells[1,1] = 'ProjektParameter[Ja/Nein]'
                    sheet.Cells[1,2] = 'Prüfen'
                    sheet.Cells[1,3] = 'ParameterName'
                    sheet.Cells[1,4] = 'GUID'
                    sheet.Cells[1,5] = 'Disziplin'
                    sheet.Cells[1,6] = 'Parametertyp'
                    sheet.Cells[1,7] = 'ParameterGroup'
                    sheet.Cells[1,8] = 'Kategorien'
                    sheet.Cells[1,9] = 'Typ oder Exemplar'
                    sheet.Cells[1,10] = 'Medien'
                    sheet.Cells[1,11] = 'Parameter Typ'
                    sheet.Cells[1,12] = 'Gruppe'
                    sheet.Cells[1,13] = 'Hilfe'
                    sheet.Cells[1,14] = 'Berechnungsformel'
                    sheet.Cells[1,15] = 'Anmerkung'
                    sheet.Cells[2,2] = 'Shared'
                    sheet.Cells[2,3] = Data[1]
                    sheet.Cells[2,4] = Data[2]
                    sheet.Cells[2,5] = Data[3]
                    sheet.Cells[2,6] = Data[4]
            else:
                sheetAnzahl = book.Sheets.Count
                sheet = book.Sheets.Add(After=sheetAnzahl)
                sheet.Name = Data[0]
                sheet.Cells[1,1] = 'ProjektParameter[Ja/Nein]'
                sheet.Cells[1,2] = 'Prüfen'
                sheet.Cells[1,3] = 'ParameterName'
                sheet.Cells[1,4] = 'GUID'
                sheet.Cells[1,5] = 'Disziplin'
                sheet.Cells[1,6] = 'Parametertyp'
                sheet.Cells[1,7] = 'ParameterGroup'
                sheet.Cells[1,8] = 'Kategorien'
                sheet.Cells[1,9] = 'Typ oder Exemplar'
                sheet.Cells[1,10] = 'Medien'
                sheet.Cells[1,11] = 'Parameter Typ'
                sheet.Cells[1,12] = 'Gruppe'
                sheet.Cells[1,13] = 'Hilfe'
                sheet.Cells[1,14] = 'Berechnungsformel'
                sheet.Cells[1,15] = 'Anmerkung'
                sheet.Cells[2,2] = 'Shared'
                sheet.Cells[2,3] = Data[1]
                sheet.Cells[2,4] = Data[2]
                sheet.Cells[2,5] = Data[3]
                sheet.Cells[2,6] = Data[4]



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))