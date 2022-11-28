# coding: utf8
from pyrevit import script, forms
from rpw import *
import time
from Autodesk.Revit.DB import Transaction

import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

start = time.time()


__title__ = "0.32 Import_Projekt_Shared_Para_von_Excel"
__doc__ = """Erstellung eines Shared- und ProjektParameter excel. Wenn es nicht vorhanden ist, wird aktualisiert"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

from pyIGF_logInfo import getlog
getlog(__title__)


uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

ex = Excel.ApplicationClass()

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()


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
            aktuellSharedPara.append([definitionName,definitionInfo])
    aktuellSharedPara.sort()
    for i in aktuellSharedPara:
        if any(i[1]):
            output.print_table(
                table_data=i[1],
                title="Aktuelle SharedParameters in Group "+i[0],
                columns=['Name', 'GUID','Disziplin', 'ParameterType']
            )

    return aktuellSharedPara,groupListe

def AktuellProjektPara():
    map = doc.ParameterBindings
    dep = map.ForwardIterator()
    Liste = []
    while(dep.MoveNext()):
        try:
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
        except:
            pass

    Liste.sort()
    ProjektParaListe = [i[0] for i in Liste]
    if any(Liste):
        output.print_table(
            table_data=Liste,
            title="Aktuelle Projectparameter",
            columns=['ParameterName', 'Disziplin','ParameterTyp','ParameterGroup',
                     'Typ or Exemplar','Kategorie']
        )

    return Liste,ProjektParaListe

def Vorhandenen_Para(SharedListe,ProjektListe):
    OutListe = []
    for item in SharedListe:
        for i in item[1]:
            if i[0] in ProjektListe:
                OutListe.append(i[0])
    return OutListe

def ExcelPara(filepath):
    book = ex.Workbooks.Open(filepath)
    GroupList = []
    paraInGroup = []
    Parame = {}
    for sheet in book.Worksheets:
        ExcelParaListe = []
        Group = sheet.Name
        if not Group in ['Hinweis', 'Revit (Original)']:
            GroupList.append(Group)
            rows = sheet.UsedRange.Rows.Count
            for row in range(2, rows + 1):
                rowlist = []
                if sheet.Cells[row, 1].Value2 == 'x':
                    for col in [3,4,5,6,7,8,9,13]:
                        Wert = sheet.Cells[row, col].Value2
                        rowlist.append(Wert)
                    wert1 = sheet.Cells[row, 3].Value2
                    wert7 = sheet.Cells[row, 7].Value2
                    wert8 = sheet.Cells[row, 8].Value2

                    Parame[wert1] = [wert7,wert8]

                    ExcelParaListe.append(rowlist)

        paraInGroup.append([Group,ExcelParaListe])

    for i in paraInGroup:
        if any(i[1]):
            output.print_table(
                table_data=i[1],
                title="Aktuelle SharedParameters in Group "+i[0]+" aus Excel",
                columns=['Name', 'GUID', 'Disziplin', 'ParameterTyp',
                         'ParameterGroup','Kategorie','Typ oder Exemplar','Hilfe']
            )

    book.Save()
    book.Close()
    return paraInGroup,GroupList,Parame

def NeuPara(inExcelData,inProjektSharedPara):
    felSharedPara = []
    for booklData in inExcelData:
        paraInGroup = []
        for Para in booklData[1]:
            if not Para[0] in inProjektSharedPara:
                paraInGroup.append(Para)
        if any(paraInGroup):
            felSharedPara.append([booklData[0],paraInGroup])

    for i in felSharedPara:
        if any(i[1]):
            output.print_table(
                table_data=i[1],
                title="neue Parameter aus Excel____"+i[0],
                columns=['Name', 'GUID','Disziplin', 'ParameterTyp', 'ParameterGroup',
                         'Kategorie','Typ oder Exemplar','Hilfe']
            )
    return felSharedPara

def ALLParameterTyp():
    allParaType = []
    alltype = System.Enum.GetValues(DB.ParameterType.Invalid.GetType())
    for i in alltype:
        Type = i.ToString()
        if not Type in ['Invalid','FamilyType']:
            type = DB.LabelUtils.GetLabelFor(i)
            dis = DeiziplinErmitteln(type,Type)
            allParaType.append([dis,type,i])

    allParaType.sort()
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

def AddGroup(inFile,InGroup):
    dgs = inFile.Groups
    Groups = [i.Name for i in dgs]
    if not InGroup in Groups:
        file.Groups.Create(InGroup)

sharedParaInGroup,sharedParaGroup = AktuellSharedPara(file)
projektPara,projektParaName = AktuellProjektPara()
vorhandenenPara = Vorhandenen_Para(sharedParaInGroup,projektParaName)
excelPath = rpw.ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
excelParaInGroup,excelGroup, Para_Aktu = ExcelPara(excelPath)
neuParameter = NeuPara(excelParaInGroup,vorhandenenPara)
allType = ALLParameterTyp()
allCate = AllKategorien()
allGroup = ALLParameterGroup()
# create SharedParameter Group
for i in excelGroup:
    AddGroup(file, i)
# Informationen fehlen
ErrorParameter = []
updatePara = vorhandenenPara[:]
ErrorParameter2 = []
if any(neuParameter):
    trans0 = Transaction(doc, 'Parameter erstellen')
    trans0.Start()
    if forms.alert('neue Parameter erstellen?', ok=False, yes=True, no=True):
        for neuShaPara in neuParameter:
            title = '{value}/{max_value} Parameter in Group '+ neuShaPara[0] +' erstellen'
            with forms.ProgressBar(title=title,cancellable=True, step=1) as pb2:
                n_1 = 0
                definitionGroup = None
                Definitions = []

                for dg in file.Groups:
                    if dg.Name == neuShaPara[0]:
                        definitionGroup = dg
                if definitionGroup:
                    Definitions = [d.Name for d in definitionGroup.Definitions]

                for item in neuShaPara[1]:
                    if pb2.cancelled:
                        script.exit()
                    n_1 += 1
                    pb2.update_progress(n_1, len(neuShaPara[1]))

                    Name = item[0]
                    GuidExcel = item[1]
                    disziplin = item[2]
                    parameterTypeName = item[3]
                    parameterGroup = item[4]
                    kategorieListe = item[5]
                    typeOrExemplar = item[6]
                    Description = item[7]
                    parameterType = None
                    if Name in updatePara:
                        neu_item = item[:]
                        neu_item.insert(0,neuShaPara[0])
                        logger.error("Parameter {} konnte nicht erstellt werden".format(Name))
                        ErrorParameter2.append(neu_item)
                        continue
                    for type in allType:
                        if disziplin == type[0] and parameterTypeName == type[1]:
                            parameterType = type[2]
                    if all([Name,parameterType]):
                        sharedPara = None
                        if not Name in Definitions:
                            DefiCrea = DB.ExternalDefinitionCreationOptions(Name, parameterType)
                            if GuidExcel:
                                DefiCrea.GUID = System.Guid(GuidExcel)
                            if Description:
                                DefiCrea.Description = Description
                            try:
                                sharedPara = definitionGroup.Definitions.Create(DefiCrea)
                                logger.info("SharedParameter {} wird erstellt".format(Name))
                            except:
                                logger.error(Name)
                            
                        else:
                            for defi in definitionGroup.Definitions:
                                if defi.Name == Name:
                                    sharedPara = defi
                        if not Name in projektParaName:
                            if all([parameterGroup,kategorieListe,typeOrExemplar]):
                                ParCatList = kategorieListe.split(',')
                                ParaGroup = None
                                for i in allGroup:
                                    try:
                                        ParaGroup = allGroup[i]
                                    except:
                                        pass
                                ParaCatSet = app.Create.NewCategorySet()
                                for i in ParCatList:
                                    if i in allCate.keys():
                                        ParaCatSet.Insert(allCate[i])

                                binding = None

                                if typeOrExemplar == 'Type':
                                    binding = app.Create.NewTypeBinding(ParaCatSet)
                                elif typeOrExemplar == 'Exemplar':
                                    binding = app.Create.NewInstanceBinding(ParaCatSet)

                                map = uiapp.ActiveUIDocument.Document.ParameterBindings
                                map.Insert(sharedPara,binding, ParaGroup)
                                logger.info("ProjektParameter {} wird erstellt".format(Name))
                        updatePara.append(Name)
                    else:
                        logger.error("Parameter {} konnte nicht erstellt werden".format(item[0]))
                        ErrorParameter.append([neuShaPara[0],item[0]])
    trans0.Commit()

else:
    Task = rpw.ui.forms.Alert
    Task('Alle Parameter in Excel sind bereit in Revit', title="Parameter", header="Parameter")

if any(ErrorParameter):
    output.print_table(
        table_data=ErrorParameter,
        title="Die folgenden Parameter konnten aufgrund fehlender Informationen nicht erstellt werden",
        columns=['DefinitionGroup', 'Name']
    )

if any(ErrorParameter2):
    output.print_table(
        table_data=ErrorParameter2,
        title="Die folgenden Parameter konnten nicht erstellt werden",
        columns=['DefinitionGroup', 'Name', 'GUID','Disziplin', 'ParameterTyp',
                 'ParameterGroup', 'Kategorie','Typ oder Exemplar']
    )

if forms.alert("Parameter aktualisieren?", ok=False, yes=True, no=True):
    tran = Transaction(doc,'Parameter aktualisieren')
    tran.Start()
    with forms.ProgressBar(title="{value}/{max_value} Parameter aktualisieren",cancellable=True, step=10) as pb:
        n_1 = 0
        map = doc.ParameterBindings
        dep = map.ForwardIterator()
        Liste = []
        while(dep.MoveNext()):
            n_1 += 1
            if pb.cancelled:
                script.exit()
            definition = None
            Name = None

            try:
                definition = dep.Key
                Name = definition.Name
            except Exception as e:
                logger.error(e)
            if definition:
                if Name in Para_Aktu.keys():
                    print(Para_Aktu[Name][0])
                    Group = allGroup[Para_Aktu[Name][0]]
                    Cates = Para_Aktu[Name][1]
                    cate_list = Cates.split(',')
                    Binding = dep.Current
                    Cate = app.Create.NewCategorySet()
                    for i in cate_list:
                        if i in allCate.keys():
                            Cate.Insert(allCate[i])

                    binding = None
                    Type = Binding.GetType().ToString()
                    if Type == 'Autodesk.Revit.DB.InstanceBinding':
                        binding = app.Create.NewInstanceBinding(Cate)
                    else:
                        binding = app.Create.NewTypeBinding(Cate)
                    definition.ParameterGroup = Group
                    print(definition.Name)
                    doc.ParameterBindings.ReInsert(definition, binding)


    tran.Commit()



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
