# coding: utf8
from pyrevit import script, forms
from rpw import *
import time
from Autodesk.Revit.DB import *

import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

start = time.time()


__title__ = "0.32 Import_Projekt_Shared_Para_von_Excel"
__doc__ = """Erstellung eines Shared- und ProjektParameter excel, wenn es nicht vorhanden ist"""
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

class ProParaAusShaPara:
    def __init__(self,Name,GUID,Disziplin,Type,Group,Kategorien,TypExemplar):
        self.Name = Name
        self.GUID = GUID
        self.Disziplin = Disziplin
        self.Type = Type
        self.Group = Group
        self.Kategorien = Kategorien
        self.TypExemplar = TypExemplar
    def table_row(self):
        return [self.Name, self.GUID, self.Disziplin, self.Type,
        self.Group, self.Kategorien, self.TypExemplar]


class ProPara:
    def __init__(self,Name,Disziplin,Type,Group,Kategorien,TypExemplar):
        self.Name = Name
        self.Disziplin = Disziplin
        self.Type = Type
        self.Group = Group
        self.Kategorien = Kategorien
        self.TypExemplar = TypExemplar
    def table_row(self):
        return [self.Name, self.Disziplin, self.Type,
        self.Group, self.Kategorien, self.TypExemplar]

class ShaPara:
    def __init__(self,Name, GUID, Disziplin,Type,Group,Info):
        self.Name = Name
        self.GUID = GUID
        self.Disziplin = Disziplin
        self.Type = Type
        self.Group = Group
        self.Info = Info

    def table_row(self):
        return [self.Group, self.Name, self.GUID, self.Disziplin, self.Type, self.Info,]

class ExcPara:
    def __init__(self,Name,GUID,Disziplin,Type,Group,Kategorien,TypExemplar,Gruppe,Hilfe):
        self.Name = Name
        self.GUID = GUID
        self.Disziplin = Disziplin
        self.Type = Type
        self.Group = Group
        self.Kategorien = Kategorien
        self.TypExemplar = TypExemplar
        self.Gruppe = Gruppe
        self.Hilfe = Hilfe

    def table_row(self):
        return [self.Name, self.GUID, self.Disziplin, self.Type,
        self.Group, self.Kategorien, self.TypExemplar, self.Gruppe,
        self.Hilfe]

def ALLParameterTyp():
    allParaType = {}
    alltype = System.Enum.GetValues(DB.ParameterType.Invalid.GetType())
    for i in alltype:
        Type = i.ToString()
        if not Type in ['Invalid','FamilyType']:
            type = DB.LabelUtils.GetLabelFor(i)
            dis = DeiziplinErmitteln(type,Type)
            if dis in allParaType.keys():
                allParaType[dis][type] = i
            else:
                allParaType[dis] = {dis:i}

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

def ProSharPara(Collector):
    out_Para = {}
    for el in Collector:
        Name = el.Name
        GUID = el.GuidValue.ToString()
        out_Para[Name] = GUID
    return out_Para

def AktuellSharedPara(inputfile):
    aktuellSharedPara = []
    groupListe = []
    Anzeige = []
    if inputfile.Groups:
        for dg in inputfile.Groups:
            definitionName = dg.Name
            groupListe.append(definitionName)
            if dg.Definitions:
                for d in dg.Definitions:
                    name = d.Name
                    owner = d.OwnerGroup.Name
                    GUID = d.GUID.ToString()
                    type = LabelUtils.GetLabelFor(d.ParameterType)
                    info = d.Description.ToString()
                    dis = DeiziplinErmitteln(LabelUtils.GetLabelFor(d.ParameterType),d.ParameterType.ToString())

                    shaPara = ShaPara(name,GUID,dis,type,owner,info)

                    aktuellSharedPara.append(shaPara)
                    Anzeige.append([owner,name,GUID,dis,type,info])

    if any(Anzeige):
        output.print_table(
            table_data=Anzeige,
            title="Aktuelle SharedParameters in Group "+i[0],
            columns=['Group','Name', 'GUID','Disziplin', 'ParameterType', 'Hilfe']
        )

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
    # if any(Liste):
    #     output.print_table(
    #         table_data=Liste,
    #         title="Aktuelle Projectparameter",
    #         columns=['ParameterName', 'Disziplin','ParameterTyp','ParameterGroup',
    #                  'Typ or Exemplar','Kategorie']
    #     )

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

                    ExcelParaListe.append(rowlist)
                    print(rowlist)
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
    return paraInGroup,GroupList

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



def AddGroup(inFile,InGroup):
    dgs = inFile.Groups
    Groups = [i.Name for i in dgs]
    if not InGroup in Groups:
        file.Groups.Create(InGroup)

sharedParaInGroup,sharedParaGroup = AktuellSharedPara(file)
projektPara,projektParaName = AktuellProjektPara()
vorhandenenPara = Vorhandenen_Para(sharedParaInGroup,projektParaName)
excelPath = rpw.ui.forms.TextInput('Excel: ', default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\IGF_SharedParameter.xlsx")
excelParaInGroup,excelGroup = ExcelPara(excelPath)
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
                            sharedPara = definitionGroup.Definitions.Create(DefiCrea)
                            logger.info("SharedParameter {} wird erstellt".format(Name))
                        else:
                            for defi in definitionGroup.Definitions:
                                if defi.Name == Name:
                                    sharedPara = defi
                        if not Name in projektParaName:
                            if all([parameterGroup,kategorieListe,typeOrExemplar]):
                                ParCatList = kategorieListe.split(',')
                                ParaGroup = None
                                for i in allGroup:
                                    if i[0] == parameterGroup:
                                        ParaGroup = i[1]
                                ParaCatSet = app.Create.NewCategorySet()
                                for i in ParCatList:
                                    if i in allCate.keys():
                                        ParaCatSet.Insert(allCate[i])
                                binding = None

                                if typeOrExemplar == 'Type':
                                    binding = app.Create.NewTypeBinding(ParaCatSet)
                                elif typeOrExemplar == 'Exemplar':
                                    binding = app.Create.NewInstanceBinding(ParaCatSet)

                                map = uiapp.ActiveUIDocument.Document.ParameterBindings;
                                map.Insert(sharedPara,binding, ParaGroup);
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

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
