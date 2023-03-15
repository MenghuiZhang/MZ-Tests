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


__title__ = "9.32 Projekt- und SharedParameter aus Excel erstellen"
__doc__ = """Erstellung eines Shared- und ProjektParameter excel, wenn es in den SharedParametern nicht vorhanden ist"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

ex = Excel.ApplicationClass()

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()


def AktuellSharedPara(inputfile):
    aktuellSharedPara = []
    for dg in inputfile.Groups:
        for d in dg.Definitions:
            name = d.Name
            guid = d.GUID.ToString()
            owner = d.OwnerGroup.Name
            TYPE = d.ParameterType.ToString()
            type = DB.LabelUtils.GetLabelFor(d.ParameterType)
            commen = ['Text', 'Ganzzahl', 'Zahl', 'Länge', 'Fläche', 'Volumen', 'Winkel', 'Neigung', 'Währung',
                      'Massendichte','Zeit', 'Geschwindigkeit', 'URL', 'Material', 'Bild', 'Ja/Nein', 'Mehrzeiliger Text']
            Energie = ['Energie', 'Wärmedurchgangskoeffizient', 'Thermischer Widerstand', 'Thermisch wirksame Masse',
                       'Wärmeleitfähigkeit', 'Spezifische Wärme', 'Spezifische Verdunstungswärme', 'Permeabilität']
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
            definition = [owner,name,guid,dis,type]
            aktuellSharedPara.append(definition)
    aktuellSharedPara.sort()
    if any(aktuellSharedPara):
        output.print_table(
            table_data=aktuellSharedPara,
            title="Aktuelle SharedParameters",
            columns=['DefinitionGroup', 'Name', 'guid', 'Disziplin','ParameterType']
        )

    return aktuellSharedPara

def AllKategorien():
    CateListe = []
    for cat in doc.Settings.Categories:
        name = cat.Name
        CateListe.append([name, cat])
    CateListe.sort()

    return CateListe

def ExcelPara(filepath):
    ExcelParaListe = []
    book = ex.Workbooks.Open(filepath)
    GroupList = []
    for sheet in book.Worksheets:
        Group = sheet.Name
        if not Group in ['Hinweis', 'Revit (Original)']:
            GroupList.append(Group)
            rows = sheet.UsedRange.Rows.Count
            for row in range(2, rows + 1):
                rowlist = [Group]
                if sheet.Cells[row, 1].Value2ToString() == 'x':
                    for col in range(3, 9):
                        Wert = sheet.Cells[row, col].Value2
                        rowlist.append(Wert)
                    ExcelParaListe.append(rowlist)
                else:
                    pass


    if any(ExcelParaListe):
        output.print_table(
            table_data=ExcelParaListe,
            title="Parameter aus Excel",
            columns=['DefinitionGroup', 'Name', 'Disziplin', 'ParameterTyp',
                     'ParameterGroup','Kategorie','Typ oder Exemplar']
        )
    return ExcelParaListe,GroupList


def AktuSharedParaSort(AktuSharPara):
    Group = []
    GroupPara = []
    for i in AktuSharPara:
        Group.append(i[0])
    Group = set(Group)
    Group = list(Group)
    for j in Group:
        Liste = []
        for i in AktuSharPara:
            if i[0] == j:
                Liste.append(i[1])
        GroupPara.append([j,Liste])
    GroupPara.sort()
    return GroupPara


def FelendeSharPara(ExcelData,AktuSharedParaInGroup):
    felSharedPara = []
    groupListe = [i[0] for i in AktuSharedParaInGroup]
    for excelData in ExcelData:
        if excelData[0] in groupListe:
            for aktuP in AktuSharedParaInGroup:
                if aktuP[0] == excelData[0]:
                    if excelData[1] in aktuP[1]:
                        pass
                    else:
                        felSharedPara.append(excelData)
        else:
            felSharedPara.append(excelData)

    if any(felSharedPara):
        output.print_table(
            table_data=felSharedPara,
            title="neue Parameter aus Excel",
            columns=['DefinitionGroup', 'Name', 'Disziplin', 'ParameterTyp',
                     'ParameterGroup','Kategorie','Typ oder Exemplar']
        )
    return felSharedPara

def ALLParameterTyp():
    allParaType = []
    alltype = System.Enum.GetValues(DB.ParameterType.Invalid.GetType())
    for i in alltype:
        typename = i.ToString()
        merkmale = typename[:3]

        if typename != 'Invalid' and typename != 'FamilyType':
            dis = None
            name = DB.LabelUtils.GetLabelFor(i)
            if merkmale == 'HVA':
                dis = 'Lüftung'
            elif merkmale == 'Pip':
                dis = 'Rohre'
            elif merkmale == 'Ele':
                dis = 'Elektro'
            else:
                pass
            if typename == 'HVACEnergy':
                dis = 'Energie'

            allParaType.append([dis,name,i])

    allParaType.sort()

    return allParaType

def ALLParameterGroup():
    allParaGro = []
    allGro = System.Enum.GetValues(DB.BuiltInParameterGroup.PG_ROUTE_ANALYSIS.GetType())
    for i in allGro:
        name = DB.LabelUtils.GetLabelFor(i)
        allParaGro.append([name, i])
    allParaGro.sort()

    return allParaGro

def AddGroup(inFile,InGroup):
    dgs = inFile.Groups
    Groups = [i.Name for i in dgs]
    if not InGroup in Groups:
        file.Groups.Create(InGroup)


def AddPara(inFile,InName,InSharedParaTyp,InDefinitionGroup,InParameterGroup,InKategorie,InTypOderExem):
    ParCatList = InKategorie.split(',')
    ParaGroup = None
    for i in allGroup:
        if i[0] == InParameterGroup:
            ParaGroup = i[1]
    ParaCatSet = app.Create.NewCategorySet()
    for i in ParCatList:
        for j in allCate:
            if i == j[0]:
                ParaCatSet.Insert(j[1])

    for dg in inFile.Groups:
        if dg.Name == InDefinitionGroup:
            addNeuSharedPara = dg.Definitions.Create(DB.ExternalDefinitionCreationOptions(InName, InSharedParaTyp))
            logger.info("SharedParameter {} wird erstellt".format(InName))
            binding = None
            if InTypOderExem == 'Type':
                binding = app.Create.NewTypeBinding(ParaCatSet)
            elif InTypOderExem == 'Exemplar':
                binding = app.Create.NewInstanceBinding(ParaCatSet)

            map = uiapp.ActiveUIDocument.Document.ParameterBindings;
            map.Insert(addNeuSharedPara,binding, ParaGroup);
            logger.info("ProjektParameter {} wird erstellt".format(InName))


aktuSharedPara = AktuellSharedPara(file)
aktuSharedParaInGroup = AktuSharedParaSort(aktuSharedPara)
excelPath = rpw.ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
excelData,ExGroupList = ExcelPara(excelPath)
neuSharParaListe = FelendeSharPara(excelData,aktuSharedParaInGroup)
allType = ALLParameterTyp()
allCate = AllKategorien()
allGroup = ALLParameterGroup()
# create SharedParameter Group
for i in ExGroupList:
    AddGroup(file, i)
# Informationen fehlen
ErrorParameter = []
if any(neuSharParaListe):
    if forms.alert('neue Parameter erstellen?', ok=False, yes=True, no=True):
        with forms.ProgressBar(title='{value}/{max_value} Parameter erstellen',
                               cancellable=True, step=1) as pb2:
            n_1 = 0
            trans0 = Transaction(doc, 'Add parameters')
            trans0.Start()
            for neuShaPara in neuSharParaListe:
                if pb2.cancelled:
                    script.exit()
                n_1 += 1
                pb2.update_progress(n_1, len(neuSharParaListe))
                if all(neuShaPara):
                    Name = neuShaPara[1]
                    definitionGroup = neuShaPara[0]
                    disziplin = neuShaPara[2]
                    parameterGroup = neuShaPara[4]
                    kategorieListe = neuShaPara[5]
                    typeOrExemplar = neuShaPara[6]
                    parameterType = DB.ParameterType.HVACEnergy
                    for ty in allType:
                        if Name == 'Energie':
                            if disziplin == 'Lüftung':
                                parameterType = DB.ParameterType.HVACPower
                            elif disziplin == 'Energie':
                                parameterType = DB.ParameterType.HVACEnergy
                            else:
                                parameterType = DB.ParameterType.Energy
                        else:
                            if disziplin in ['Lüftung', 'HVAC', 'Rohre']:
                                if disziplin == ty[0] and neuShaPara[3] == ty[1]:
                                    parameterType = ty[2]
                            else:
                                if neuShaPara[3] == ty[1]:
                                    parameterType = ty[2]

                    AddPara(file, Name, ParameterType, definitionGroup, parameterGroup, kategorieListe, typeOrExemplar)
                else:
                    ErrorParameter.append(neuShaPara)



            trans0.Commit()

else:
    Task = rpw.ui.forms.Alert
    Task('Alle Parameter in Excel sind bereit in Revit', title="Parameter", header="Parameter")

if any(ErrorParameter):
    output.print_table(
        table_data=ErrorParameter,
        title="Die folgenden Parameter konnten aufgrund fehlender Informationen nicht erstellt werden",
        columns=['DefinitionGroup', 'Name', 'Disziplin', 'ParameterTyp',
                 'ParameterGroup', 'Kategorie', 'Typ oder Exemplar']
    )


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
