# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import time
from pyrevit import script, forms
from rpw import *
import System
from Autodesk.Revit.DB import *
from pyrevit.forms import WPFWindow
from System.Windows.Forms import *
import os


start = time.time()

__title__ = "0.33 Export_Parameter_zu_Excel"
__doc__ = """die Informationen der Projekt- und SharedParameter zu Excel Exportieren"""
__author__ = "Menghui Zhang"

from pyIGF_logInfo import getlog
getlog(__title__)


logger = script.get_logger()
output = script.get_output()
config = script.get_config()

try:
    if not os.path.exists(config.file):
        config.file = ""
except:
    pass

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()
Coll = FilteredElementCollector(doc).OfClass(clr.GetClrType(SharedParameterElement)).WhereElementIsNotElementType()

# Projektparameter (out_Para[id] = [Name,GUID]) (aus SharedParameterElement)
def SharParaelement(Collector):
    out_Para = {}
    for el in Collector:
        Name = el.Name
        GUID = el.GuidValue.ToString()
        id = el.Id.ToString()
        out_Para[id] = GUID
    return out_Para


# DeiziplinErmitteln, Text--Commen
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

# Sharedparameter aus Datei-- aktuellSharedPara,groupListe
def AktuellSharedPara(inputfile):
    aktuellSharedPara = []
    sharenparamliste = []

    groupListe = []
    if inputfile.Groups:
        for dg in inputfile.Groups:
            definitionName = dg.Name
            groupListe.append(definitionName)
            definitionInfo = []
            if dg.Definitions:
                for d in dg.Definitions:
                    name = d.Name
                    GUID = d.GUID.ToString()
                    Description = d.Description
                    TYPE = d.ParameterType.ToString()
                    type = DB.LabelUtils.GetLabelFor(d.ParameterType)
                    dis = DeiziplinErmitteln(type,TYPE)
                    definitionInfo.append([name,GUID,Description,dis,type])

                    sharenparamliste.append(GUID)
            definitionInfo.sort()
            aktuellSharedPara.append([definitionName,definitionInfo])
    aktuellSharedPara.sort()

    return aktuellSharedPara,groupListe,sharenparamliste

# Projektparameter
# Liste = [[Name,dis,TypName,Group,typOrex,cateName,id],[]...]
def AktuellProjektPara():
    map = doc.ParameterBindings
    dep = map.ForwardIterator()
    Liste = []
    while(dep.MoveNext()):
        definition = dep.Key
        id = definition.Id.ToString()
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
        Liste.append([Name,dis,TypName,Group,typOrex,cateName,id])

    Liste.sort()
    return Liste



def ProjektMitGuid(ProjektListe,guiddict):
    outProjektDict = {}
    for item in ProjektListe:
        if item[6] in guiddict.Keys:
            outProjektDict[guiddict[item[6]]] = item
    return outProjektDict

# OutListe_Shared: [IGF_X,[[IGF_X_Einbauort,...],[]...]]
# OutListe_Projekt: [IGF_X,[[IGF_X_Einbauort,...],[]...]]
def ParameterMitGuidUndGroup(SharedListe,ProjektMitGuid):#'ProjektListe,ProjektParaListe'):
    OutListe_Parameter = {}
    temp = []
    ProjektMitGuid_temp = ProjektMitGuid.copy()
    for defiliste in SharedListe:
        groupSha = []
        groupPro = []
        for item in defiliste[1]:
            temp.append(item[0])
            if item[1] in ProjektMitGuid.Keys:
                proparam = ProjektMitGuid[item[1]]
                groupPro.append([proparam[0],item[1],proparam[1],proparam[2],proparam[3],proparam[4],proparam[5],item[2]])
                del ProjektMitGuid_temp[item[1]]
            else:
                groupSha.append([item[0],item[1],item[3],item[4],item[2]])

        groupSha.sort()
        groupPro.sort()
        OutListe_Parameter[defiliste[0]] = [groupSha,groupPro]
    sonstige = []
    for el in ProjektMitGuid_temp.Keys:
        proparam = ProjektMitGuid_temp[el]
        sonstige.append([proparam[0],el,proparam[1],proparam[2],proparam[3],proparam[4],proparam[5],None])
    if any(sonstige):
        OutListe_Parameter['Sonstige'] = [[],sonstige]

    return OutListe_Parameter

def ExcelLesen(inExcelPath):
    outGroupPara = {}
    out_sonstige = []
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        if sheet.Name == 'Sonstige':
            continue
        groupParaListe = {}
        groupName = sheet.Name
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            definitionName = sheet.UsedRange.Cells[row, 3].Value2
            sheet.Cells[row, 2] = '' 
            if definitionName:
                groupParaListe[definitionName] = row
        outGroupPara[groupName] = groupParaListe
    
    try:
        sonstige = book.Worksheets['Sonstige']
        rows = sonstige.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            rowdaten = []
            if not sonstige.Cells[row, 3].Value2:
                continue
            for col in [3,4,5,6,7,9,8,13]:
                daten = sonstige.Cells[row, col].Value2
                rowdaten.append(daten)
            out_sonstige.append(rowdaten)
            for col1 in range(1,16):
                sonstige.Cells[row, col1] = None
    except:
        pass

    book.Save()
    book.Close()

    return outGroupPara,out_sonstige

def sonstigebearbeiten(in_sonstige,in_ausprojekt,allsharedparam):
    sonstige = in_ausprojekt['Sonstige'][1]
    sonstigeliste = [i[1] for i in sonstige]
    for el in in_sonstige:
        if not el[1] in allsharedparam:
            if not el[1] in sonstigeliste:
                sonstige.append(el)

def GroupErstellen(inExcelPath,inSharedGroup):
    columns = ['anlegen','angelegt[Ja/Nein]','ParameterName','GUID',
               'Disziplin','Parametertyp','ParameterGroup','Kategorien',
               'Typ oder Exemplar','Medien','Parameter Typ','Gruppe',
               'Hilfe','Berechnungsformel','Anmerkung']
    book = exapp.Workbooks.Open(inExcelPath)
    inExcelGroup = [sheet.Name for sheet in book.Worksheets]
    for i in inSharedGroup.Keys:
        if len(i) > 31:
            gr = i[:31]
        else:
            gr = i
        if not gr in inExcelGroup:
            sheet = book.Worksheets.Add()
            print(gr)
            try:
                sheet.Name = gr
            except:
                pass
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
            groupExcel = {}
            if sheetname in inGroupParaExcel.Keys:
                groupExcel = inGroupParaExcel[sheetname]
            print(30*'-')
            print(sheetname)
            print(30*'-')

            if any(shared):
                title1 = '{value}/{max_value} SharedParameter in Group ' + sheetname
                with forms.ProgressBar(title=title1, cancellable=True, step=2) as pb:
                    for n, item in enumerate(shared):
                        if pb.cancelled:
                            exapp.Quit()
                            script.exit()
                        pb.update_progress(n, len(shared))
                        rows = sheet.UsedRange.Rows.Count
                        if not item[0] in groupExcel.Keys:
                            sheet.Cells[rows+1, 1] = ''
                            sheet.Cells[rows + 1,2] = 'X-SharedParameter'
                            sheet.Cells[rows + 1,3] = item[0]
                            sheet.Cells[rows + 1,4] = item[1]
                            sheet.Cells[rows + 1,5] = item[2]
                            sheet.Cells[rows + 1,6] = item[3]
                            sheet.Cells[rows + 1,12] = groupforschreiben
                            sheet.Cells[rows + 1,13] = item[4]
                            logger.info('Shared Parameter {} wird erstellt'.format(item[0]))
                        else:
                            row = groupExcel[item[0]]
                            sheet.Cells[row, 1] = ''
                            sheet.Cells[row,2] = 'X-SharedParameter'
                            sheet.Cells[row,4] = item[1]
                            sheet.Cells[row,5] = item[2]
                            sheet.Cells[row,6] = item[3]
                            sheet.Cells[row,12] = groupforschreiben
                            if sheet.Cells[row, 13].Value2 == None:
                                sheet.Cells[row,13] = item[4]
                            logger.info('Shared Parameter {} wird aktualisiert'.format(item[0]))
            if any(projek):
                title1 = '{value}/{max_value} ProjektParameter in Group ' + sheetname
                with forms.ProgressBar(title=title1, cancellable=True, step=2) as pb:
                    for n, item1 in enumerate(projek):
                        if pb.cancelled:
                            exapp.Quit()
                            script.exit()
                        pb.update_progress(n, len(projek))
                        rows = sheet.UsedRange.Rows.Count
                        if sheetname == 'Sonstige':
                            rows = n + 1
                        if not item1[0] in groupExcel.Keys:
                            sheet.Cells[rows +1, 1] = ''
                            sheet.Cells[rows + 1,2] = 'X-ProjektParameter'
                            sheet.Cells[rows + 1,3] = item1[0]
                            sheet.Cells[rows + 1,4] = item1[1]
                            sheet.Cells[rows + 1,5] = item1[2]
                            sheet.Cells[rows + 1,6] = item1[3]
                            sheet.Cells[rows + 1,7] = item1[4]
                            sheet.Cells[rows + 1,8] = item1[6]
                            sheet.Cells[rows + 1,9] = item1[5]
                            sheet.Cells[rows + 1,13] = item1[7]
                            sheet.Cells[rows + 1,12] = groupforschreiben
                            logger.info('Projektparameter {} wird erstellt'.format(item1[0]))
                        else:
                            row = groupExcel[item1[0]]
                            sheet.Cells[row, 1] = ''
                            sheet.Cells[row,2] = 'X-ProjektParameter'
                            sheet.Cells[row,4] = item1[1]
                            sheet.Cells[row,5] = item1[2]
                            sheet.Cells[row,6] = item1[3]
                            sheet.Cells[row,7] = item1[4]
                            sheet.Cells[row,8] = item1[6]
                            sheet.Cells[row,9] = item1[5]
                            sheet.Cells[row,12] = groupforschreiben
                            if sheet.Cells[row, 13].Value2 == None:
                                sheet.Cells[row,13] = item1[7]
                            logger.info('Projektparameter {} wird aktualisiert'.format(item1[0]))

        book.Save()
        book.Close()

class Suche(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self.read_config()
    def read_config(self):
        try:
            self.Adresse.Text = str(config.adresse)
        except:
            config.adresse = ""

    def write_config(self):
        config.adresse = self.Adresse.Text.encode('utf-8')
        script.save_config()

    def durchsuchen(self,sender,args):
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "BIM-ID Datei suchen"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            self.Adresse.Text = dialog.FileName
        self.write_config()
    def ok(self, sender, args):
        self.Close()


adresse = Suche('window.xaml')
adresse.ShowDialog()


ID_GUID_Dict = SharParaelement(Coll)

sharedParaInGroup,sharedParaGroup,Allsharedparam = AktuellSharedPara(file)
projektPara = AktuellProjektPara()
ProParaMitGuid = ProjektMitGuid(projektPara,ID_GUID_Dict)
exapp = Excel.ApplicationClass()
exapp.Visible = False
groupParameterdict = ParameterMitGuidUndGroup(sharedParaInGroup,ProParaMitGuid)

excelPath = config.adresse
# excelPath = ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
groupParaExcel = None
SonstigeListe = None
try:
    groupParaExcel,SonstigeListe = ExcelLesen(excelPath)
except Exception as e:
    exapp.Visible = True
    logger.error(e)
    script.exit() 
GroupErstellen(excelPath,groupParameterdict)
sonstigebearbeiten(SonstigeListe,groupParameterdict,Allsharedparam)
try:
    ExportParaZuExcel(excelPath,groupParaExcel,groupParameterdict)
except Exception as e:
    exapp.Visible = True
    logger.error(e)
    script.exit() 

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
