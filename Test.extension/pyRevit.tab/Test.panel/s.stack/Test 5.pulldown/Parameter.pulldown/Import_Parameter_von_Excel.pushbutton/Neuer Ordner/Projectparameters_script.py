# coding: utf8
from pyrevit import script, forms, revit
from Autodesk.Revit.DB import Transaction
import os
from System import Guid
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal
from pyrevit.forms import WPFWindow
from System.Windows.Forms import *


__title__ = "0.32 Import_Parameter_von_Excel"
__doc__ = """Erstellung eines Shared- und ProjektParameter excel. Wenn es nicht vorhanden ist, wird aktualisiert"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config = script.get_config()

from pyIGF_logInfo import getlog
getlog(__title__)

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

try:
    if not os.path.exists(config.adresse):
        config.adresse = ""
except:
    pass


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

excelPath = config.adresse

if not excelPath:
    logger.error('Kein Excel')
    script.exit()

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()

def DisziplinErmitteln(inParaTypName,inParaTyp):
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

def ALLParameterTyp():
    allParaType = {}
    alltype = System.Enum.GetValues(DB.ParameterType.Invalid.GetType())
    for i in alltype:
        Type = i.ToString()
        if not Type in ['Invalid','FamilyType']:
            type = DB.LabelUtils.GetLabelFor(i)
            dis = DisziplinErmitteln(type,Type)
            if not dis in allParaType.keys():
                allParaType[dis] = {}
            if not type in allParaType[dis].keys():
                allParaType[dis][type] = i
    return allParaType

def AllKategorien():
    CateListe = {}
    for cat in doc.Settings.Categories:
        name = cat.Name
        CateListe[name] = cat

    return CateListe

def ALLParameterGroup():
    allParaGro = {}
    allGro = System.Enum.GetValues(DB.BuiltInParameterGroup.PG_ROUTE_ANALYSIS.GetType())
    for i in allGro:
        name = DB.LabelUtils.GetLabelFor(i)
        allParaGro[name] = i
    return allParaGro

def AddGroup(inFile,Excel):
    dgs = inFile.Groups
    Groups = [i.Name for i in dgs]
    for excel in Excel:
        if not excel in Groups:
            file.Groups.Create(excel)

def AktuellSharedPara(inputfile):
    aktuellSharedPara_guid = {}
    aktuellSharedPara_name = {}
    groupListe = []
    if inputfile.Groups:
        for dg in inputfile.Groups:
            definitionName = dg.Name
            groupListe.append(definitionName)
            
            if dg.Definitions:
                for d in dg.Definitions:
                    name = d.Name
                    GUID = d.GUID.ToString()
                    TYPE = d.ParameterType.ToString()
                    type = DB.LabelUtils.GetLabelFor(d.ParameterType)
                    dis = DisziplinErmitteln(type,TYPE)
                    aktuellSharedPara_guid[GUID] = [definitionName,name,dis,type,d]
                    aktuellSharedPara_name[name] = [definitionName,GUID,dis,type,d]
    return aktuellSharedPara_guid,aktuellSharedPara_name,groupListe


# Projektparameter (out_Para[id] = GUID) (aus SharedParameterElement)
def SharParaelement():
    Coll = FilteredElementCollector(doc)\
        .OfClass(clr.GetClrType(SharedParameterElement)).WhereElementIsNotElementType()
    out_Para = {}
    for el in Coll:
        GUID = el.GuidValue.ToString()
        id = el.Id.ToString()
        out_Para[id] = GUID
    Coll.Dispose()
    return out_Para

# Projektparameter
# Liste = [[Name,id],[]...]
def AktuellProjektPara():
    map = doc.ParameterBindings
    dep = map.ForwardIterator()
    Liste = []
    while(dep.MoveNext()):
        definition = dep.Key
        id = definition.Id.ToString()
        Name = definition.Name
        Binding = dep.Current
        Liste.append([Name,id,Binding])

    Liste.sort()
    return Liste

def ProjektMitGuid(ProjektListe,guiddict):
    outProjektDict_guid = {}
    outProjektDict_name = {}
    for item in ProjektListe:
        if item[1] in guiddict.Keys:
            outProjektDict_guid[guiddict[item[1]]] = item[0]
            outProjektDict_name[item[0]] = guiddict[item[1]]
    return outProjektDict_guid,outProjektDict_name

def ExcelGroup(filepath):
    ex = Excel.ApplicationClass()
    book = ex.Workbooks.Open(filepath)
    GroupList = []
    for sheet in book.Worksheets:
        Group = sheet.Name
        if not Group in ['Hinweis', 'Revit (Original)','Sonstige']:
            if sheet.Cells[2, 12].Value2:
                GroupList.append(sheet.Cells[2, 12].Value2)
            else:
                GroupList.append(Group)

    book.Save()
    book.Close()
    Marshal.FinalReleaseComObject(sheet)
    Marshal.FinalReleaseComObject(book)
    ex.Quit()
    Marshal.FinalReleaseComObject(ex)
    return GroupList

            # rows = sheet.UsedRange.Rows.Count
            # for row in range(2, rows + 1):
            #     rowlist = []
            #     if sheet.Cells[row, 1].Value2 == 'x':
            #         for col in [3,4,5,6,7,8,9,13]:
            #             Wert = sheet.Cells[row, col].Value2
            #             rowlist.append(Wert)
            #         wert1 = sheet.Cells[row, 3].Value2
            #         wert7 = sheet.Cells[row, 7].Value2
            #         wert8 = sheet.Cells[row, 8].Value2

            #         Parame[wert1] = [wert7,wert8]

            #         ExcelParaListe.append(rowlist)

allType = ALLParameterTyp()
allCate = AllKategorien()
allGroup = ALLParameterGroup()

sharedPara_Guid,sharedPara_Name,sharedParaGroup = AktuellSharedPara(file)

projektPara = AktuellProjektPara()
sharedelementdict = SharParaelement()

project_guid,project_name = ProjektMitGuid(projektPara,sharedelementdict)

vorhandenenPara = Vorhandenen_Para(sharedParaInGroup,projektParaName)
#excelPath = rpw.ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
excelGroup = ExcelGroup(excelPath)

# create SharedParameter Group
AddGroup(file, excelGroup)

# Informationen fehlen

ErrorParameter = []
updatePara = vorhandenenPara[:]
ErrorParameter2 = []

trans0 = Transaction(doc, 'Parameter erstellen')
trans0.Start()
if forms.alert('Parameter erstellen?', ok=False, yes=True, no=True):
    ex = Excel.ApplicationClass()
    book = ex.Workbooks.Open(excelPath)
    GroupList = []
    for sheet in book.Worksheets:
        Group = sheet.Name
        defiGroup = Group
        if Group in ['Hinweis', 'Revit (Original)','Sonstige']:
            continue
        else:
            if sheet.Cells[2, 12].Value2:
                defiGroup = sheet.Cells[2, 12].Value2
            else:
                defiGroup = Group

        definitionGroup = file.Groups[defiGroup]
        
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            if sheet.Cells[row, 1].Value2 == 'x':
                Definame = sheet.Cells[row, 3].Value2
                if not Definame:
                    continue
                DefiGuid = sheet.Cells[row, 4].Value2
                Defidis = sheet.Cells[row, 5].Value2
                if not Defidis in allType.keys():
                    logger.error('falsche Disziplin--' + Definame + '--' + Group)
                    continue
                DefiTyp = sheet.Cells[row, 6].Value2
                if not DefiTyp in allType[Defidis].keys():
                    logger.error('falsche Parametertyp--' + Definame + '--' + Group)
                    continue
                Hilfe = sheet.Cells[row, 13].Value2
                ParaGroup = sheet.Cells[row, 7].Value2
                if not ParaGroup in allGroup.keys():
                    logger.error('falsche Parametergroup--' + Definame + '--' + Group)
                    ParaGroup = ''
                ParaKate = sheet.Cells[row, 8].Value2
                TypExem = sheet.Cells[row, 9].Value2
                if not TypExem in ['Type','Exemplar']:
                    logger.error('falsche Typ/Exemplar--' + Definame + '--' + Group)
                    logger.error('Typ/Exemplar kann nur Type oder Exemplar sein')
                    TypExem = ''
                
                sharedPara = None
                # Shared Parameter erstellen
                if Definame and Defidis and DefiTyp:
                    parameterType = DisziplinErmitteln()
                    if DefiGuid in sharedPara_Guid.Keys:
                        if Definame == sharedPara_Guid[DefiGuid][1] and Defidis == sharedPara_Guid[DefiGuid][2] and DefiTyp == sharedPara_Guid[DefiGuid][3]:
                            sharedPara = sharedPara_Guid[DefiGuid][4]
                            logger.info('SharedParameter bereits vorhanden--' + Definame)
                        else:
                            logger.error('Guid bereits verwendet--'+ Definame)
                    else:
                        if Definame in sharedPara_Name.Keys:
                            if Defidis == sharedPara_Name[Definame][2] and DefiTyp == sharedPara_Name[Definame][3]:
                                sharedPara = sharedPara_Name[Definame][4]
                                logger.info('SharedParameter bereits vorhanden--'+ Definame)
                            else:
                                logger.error('ParameterName bereits verwendet--'+ Definame)
                        else:
                            try:
                                DefiCrea = DB.ExternalDefinitionCreationOptions(Definame, parameterType)
                                if DefiGuid:
                                    DefiCrea.GUID = Guid(DefiGuid)
                                else:
                                    DefiGuid = DefiCrea.GUID
                                if Hilfe:
                                    DefiCrea.Description = Hilfe
                                
                                sharedPara = definitionGroup.Definitions.Create(DefiCrea)
                                logger.info("SharedParameter {} wird erstellt".format(Definame))
                                sheet.Cells[row, 4] = DefiGuid
                                sheet.Cells[row, 2] = 'X-SharedParameter'
                            except Exception as e:
                                logger(e)

                # Projektparameter erstellen
                if DefiGuid in project_guid.Keys:
                    logger.info('ProjektParameter bereits vorhanden--'+ Definame)
                else:
                    if all([ParaGroup,ParaKate,TypExem]):
                        ParCatList = ParaKate.split(',')
                        try:
                            ParaGroup = allGroup[ParaGroup]
                        except:
                            ParaGroup = None
                            
                        ParaCatSet = app.Create.NewCategorySet()
                        for i in ParCatList:
                            if i in allCate.keys():
                                ParaCatSet.Insert(allCate[i])

                        binding = None

                        if TypExem == 'Type':
                            binding = app.Create.NewTypeBinding(ParaCatSet)
                        elif TypExem == 'Exemplar':
                            binding = app.Create.NewInstanceBinding(ParaCatSet)

                        map = uiapp.ActiveUIDocument.Document.ParameterBindings
                        if ParaGroup:
                            map.Insert(sharedPara,binding, ParaGroup)
                        else:
                            map.Insert(sharedPara,binding)
                        logger.info("ProjektParameter {} wird erstellt".format(Definame))
                updatePara.append(Definame)
            else:
                logger.error("Parameter {} konnte nicht erstellt werden".format(Definame))
               

    book.Save()
    book.Close()
    Marshal.FinalReleaseComObject(sheet)
    Marshal.FinalReleaseComObject(book)
    ex.Quit()
    Marshal.FinalReleaseComObject(ex)

trans0.Commit()





    # for neuShaPara in neuParameter:
    #     title = '{value}/{max_value} Parameter in Group '+ neuShaPara[0] +' erstellen'
    #     with forms.ProgressBar(title=title,cancellable=True, step=1) as pb2:
    #         n_1 = 0
    #         definitionGroup = None
    #         Definitions = []

    #         for dg in file.Groups:
    #             if dg.Name == neuShaPara[0]:
    #                 definitionGroup = dg
    #         if definitionGroup:
    #             Definitions = [d.Name for d in definitionGroup.Definitions]

    #         for item in neuShaPara[1]:
    #             if pb2.cancelled:
    #                 script.exit()
    #             n_1 += 1
    #             pb2.update_progress(n_1, len(neuShaPara[1]))

    #             Name = item[0]
    #             GuidExcel = item[1]
    #             disziplin = item[2]
    #             parameterTypeName = item[3]
    #             parameterGroup = item[4]
    #             kategorieListe = item[5]
    #             typeOrExemplar = item[6]
    #             Description = item[7]
    #             parameterType = None
    #             if Name in updatePara:
    #                 neu_item = item[:]
    #                 neu_item.insert(0,neuShaPara[0])
    #                 logger.error("Parameter {} konnte nicht erstellt werden".format(Name))
    #                 ErrorParameter2.append(neu_item)
    #                 continue
    #             for type in allType:
    #                 if disziplin == type[0] and parameterTypeName == type[1]:
    #                     parameterType = type[2]
                # if all([Name,parameterType]):
                #     sharedPara = None
                #     if not Name in Definitions:
                #         DefiCrea = DB.ExternalDefinitionCreationOptions(Name, parameterType)
                #         if GuidExcel:
                #             DefiCrea.GUID = Guid(GuidExcel)
                #         if Description:
                #             DefiCrea.Description = Description
                #         try:
                #             sharedPara = definitionGroup.Definitions.Create(DefiCrea)
                #             logger.info("SharedParameter {} wird erstellt".format(Name))
                #         except:
                #             logger.error(Name)
                        
                #     else:
                #         for defi in definitionGroup.Definitions:
                #             if defi.Name == Name:
                #                 sharedPara = defi
#                     if not Name in projektParaName:
#                         if all([parameterGroup,kategorieListe,typeOrExemplar]):
#                             ParCatList = kategorieListe.split(',')
#                             try:
#                                 ParaGroup = allGroup[parameterGroup]
#                             except:
#                                 ParaGroup = None
                                
#                             ParaCatSet = app.Create.NewCategorySet()
#                             for i in ParCatList:
#                                 if i in allCate.keys():
#                                     ParaCatSet.Insert(allCate[i])

#                             binding = None

#                             if typeOrExemplar == 'Type':
#                                 binding = app.Create.NewTypeBinding(ParaCatSet)
#                             elif typeOrExemplar == 'Exemplar':
#                                 binding = app.Create.NewInstanceBinding(ParaCatSet)

#                             map = uiapp.ActiveUIDocument.Document.ParameterBindings
#                             print(ParaGroup)
#                             print()
#                             if ParaGroup:
#                                 map.Insert(sharedPara,binding, ParaGroup)
#                             else:
#                                 map.Insert(sharedPara,binding)
#                             logger.info("ProjektParameter {} wird erstellt".format(Name))
#                     updatePara.append(Name)
#                 else:
#                     logger.error("Parameter {} konnte nicht erstellt werden".format(item[0]))
#                     ErrorParameter.append([neuShaPara[0],item[0]])
# trans0.Commit()

# else:
#     Task = rpw.ui.forms.Alert
#     Task('Alle Parameter in Excel sind bereit in Revit', title="Parameter", header="Parameter")

# if any(ErrorParameter):
#     output.print_table(
#         table_data=ErrorParameter,
#         title="Die folgenden Parameter konnten aufgrund fehlender Informationen nicht erstellt werden",
#         columns=['DefinitionGroup', 'Name']
#     )

# if any(ErrorParameter2):
#     output.print_table(
#         table_data=ErrorParameter2,
#         title="Die folgenden Parameter konnten nicht erstellt werden",
#         columns=['DefinitionGroup', 'Name', 'GUID','Disziplin', 'ParameterTyp',
#                  'ParameterGroup', 'Kategorie','Typ oder Exemplar']
#     )

# if forms.alert("Parameter aktualisieren?", ok=False, yes=True, no=True):
#     tran = Transaction(doc,'Parameter aktualisieren')
#     tran.Start()
#     with forms.ProgressBar(title="{value}/{max_value} Parameter aktualisieren",cancellable=True, step=10) as pb:
#         n_1 = 0
#         map = doc.ParameterBindings
#         dep = map.ForwardIterator()
#         Liste = []
#         while(dep.MoveNext()):
#             n_1 += 1
#             if pb.cancelled:
#                 script.exit()
#             definition = None
#             Name = None

#             try:
#                 definition = dep.Key
#                 Name = definition.Name
#             except Exception as e:
#                 logger.error(e)
#             if definition:
#                 if Name in Para_Aktu.keys():
#                     print(Para_Aktu[Name][0])
#                     Group = allGroup[Para_Aktu[Name][0]]
#                     Cates = Para_Aktu[Name][1]
#                     cate_list = Cates.split(',')
#                     Binding = dep.Current
#                     Cate = app.Create.NewCategorySet()
#                     for i in cate_list:
#                         if i in allCate.keys():
#                             Cate.Insert(allCate[i])

#                     binding = None
#                     Type = Binding.GetType().ToString()
#                     if Type == 'Autodesk.Revit.DB.InstanceBinding':
#                         binding = app.Create.NewInstanceBinding(Cate)
#                     else:
#                         binding = app.Create.NewTypeBinding(Cate)
#                     definition.ParameterGroup = Group
#                     print(definition.Name)
#                     doc.ParameterBindings.ReInsert(definition, binding)


#     tran.Commit()
