# coding: utf8
from IGF_log import getlog
from IGF_forms import ExcelSuche
from IGF_lib import AllProjectParams_2020,AllSharedParams_2020,CreateSharedParam_2020,CreateProjFromSharedParam_2020,AllParameterType_2020,AllCategories,AllParameterGroup_2020,\
                    AllProjectParams_2022,AllSharedParams_2022,CreateSharedParam_2022,CreateProjFromSharedParam_2022,AllParameterType_2022,AllParameterGroup_2022
from pyrevit import script, forms
from rpw import revit,DB
import os
import time
import System
from config import ProjektConfig
from IGF_Klasse._parameter import SharedParameter,ProjektParameter


__title__ = "0.01 Import Parameter von Excel" 
__doc__ = """erstellt Shared- und ProjektParameter von excel. 

+: erstellen
-: löschen

update: Performance verbessert

[2022.10.17]
Version: 2.2"""
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

config = ProjektConfig.get_config_static(projectinfo,'Parameterimport')
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

# Excel lesen
# Parameter Klass aus Excel
class E_Parameter:
    def __init__(self,name = '',Disziplin = '', paramtype = '',paramgroup = '',kate = '',exemplar = False,Group = '',info = '',GUID = ''):
        self.Name = name
        self.GUID = GUID
        self.Disziplin = Disziplin        
        self.ParamType = paramtype
        self.ParamGroup = paramgroup
        self.Kategorien = kate
        self.Exemplar = exemplar
        self.Group = Group
        self.info = info
        self.row = None

excelPath = config.get_value('adresse')
if revitversion == '2020':
    import excel._NPOI_2020 as _NPOI

else:
    import excel._NPOI_2022 as _NPOI
# excelPath = r'C:\Users\Zhang\Desktop\IGF_Parameter_RUB-NA_2020.xlsx'
fs = _NPOI.FileStream(excelPath,_NPOI.FileMode.Open,_NPOI.FileAccess.Read)
book1 = _NPOI.np.XSSF.UserModel.XSSFWorkbook(fs)

def get_cell_Daten(sheet,r,c):
    row = sheet.GetRow(r)
    if row:
        cell = row.GetCell(c)
        if cell:
            return cell.StringCellValue
        else:
            return ''
    else:
        return ''

def get_cell(row,column):
    cell = row.GetCell(column)
    if cell:return cell
    else:
        cell = row.CreateCell(column)
        return cell

Excel_Param_PLUS = []
Excel_Param_MINUS = []
GroupList = []

sheetscount = book1.NumberOfSheets
for n in range(sheetscount):
    sheet = book1.GetSheetAt(n)
    SheetName = sheet.SheetName
    if SheetName in ['Hinweis','Keine Gruppe','Revit (Original)']:
        continue

    rows = sheet.LastRowNum
    for row in range(1,rows+1):
        row_daten = sheet.GetRow(row)
        if row_daten:
            if get_cell_Daten(sheet,row,0) == '+':
                param = E_Parameter(
                    name=get_cell_Daten(sheet,row,2),\
                    GUID=get_cell_Daten(sheet,row,3),\
                    Disziplin=get_cell_Daten(sheet,row,4),\
                    paramtype=get_cell_Daten(sheet,row,5),\
                    paramgroup=get_cell_Daten(sheet,row,6),\
                    kate=get_cell_Daten(sheet,row,7),\
                    exemplar=get_cell_Daten(sheet,row,8),\
                    Group=get_cell_Daten(sheet,row,11),\
                    info=get_cell_Daten(sheet,row,12))
                param.row = row_daten
                Excel_Param_PLUS.append(param)
                if param.Group not in GroupList:
                    GroupList.append(param.Group)
                

            elif get_cell_Daten(sheet,row,0) == '-':
                param = E_Parameter(
                    name=get_cell_Daten(sheet,row,2),\
                    GUID=get_cell_Daten(sheet,row,3),\
                    Disziplin=get_cell_Daten(sheet,row,4),\
                    paramtype=get_cell_Daten(sheet,row,5),\
                    paramgroup=get_cell_Daten(sheet,row,6),\
                    kate=get_cell_Daten(sheet,row,7),\
                    exemplar=get_cell_Daten(sheet,row,8),\
                    Group=get_cell_Daten(sheet,row,11),\
                    info=get_cell_Daten(sheet,row,12))
                param.row = row_daten
                Excel_Param_MINUS.append(param)

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()

# Group erstellen
Groups = [i.Name for i in file.Groups]
for i in GroupList:
    if not i in Groups:
        file.Groups.Create(i)

t = DB.Transaction(doc)
# SharedParameter, ProjektParameter
if Excel_Param_PLUS:
    t.Start('Parameter erstellen')
    if revitversion == '2020':
        sharedparams,sharedparam_defis = AllSharedParams_2020()
        projectparams,projeparam_defis = AllProjectParams_2020()
        ParamGroup_dict = AllParameterGroup_2020()
        ParamType_dict = AllParameterType_2020()
        Categories_dict = AllCategories()
        dmap = doc.ParameterBindings

        with forms.ProgressBar(title="{value}/{max_value} Parameter erstellt",cancellable=True, step=1) as pb1:
            for n,e_param in enumerate(Excel_Param_PLUS):
                if pb1.cancelled:
                    t.RollBack()
                    break
                pb1.update_progress(n+1,len(Excel_Param_PLUS))

                if e_param.GUID in projectparams.keys():
                    if e_param.Disziplin == projectparams[e_param.GUID][1] and e_param.ParamType == projectparams[e_param.GUID][2]:
                        logger.info("ProjektParameter {} vorhanden".format(e_param.Name))
                        try:
                            cell_0 = get_cell(e_param.row,0)
                            cell_1 = get_cell(e_param.row,1)
                            cell_0.SetCellValue('')
                            cell_1.SetCellValue('Projektparameter')
                        except:
                            pass
  
                    else:
                        logger.error('Parameter {} konnte nicht erstellt werden wegen GUID-Konflikt'.format(e_param.Name))
                        continue

                    kate_excel_temp = e_param.Kategorien.split(',')
                    kate_excel = []
                    kate_revit = projectparams[e_param.GUID][5].split(',')
                    for i in kate_excel_temp:
                        if i:
                            while(i[0] == ' '):
                                i = i[1:]
                            kate_excel.append(i)
                    
                    if e_param.Exemplar != projectparams[e_param.GUID][4]:
                        logger.warning('Parameter {} vorhanden, aber Typ (Exemplar) passt nicht.'.format(e_param.Name))

                    if e_param.ParamGroup == projectparams[e_param.GUID][3] and sorted(kate_excel) == sorted(kate_revit):
                        continue
                    else:
                        if e_param.ParamGroup != projectparams[e_param.GUID][3] and sorted(kate_excel) != sorted(kate_revit):
                            text = 'Kategorien and Parametergroup'
                        elif e_param.ParamGroup != projectparams[e_param.GUID][3]:
                            text = 'Parametergroup'
                        else:
                            text = 'Kategorien'

                        if forms.alert("Parameter {} vorhanden, aber {} passt nicht. Möchten Sie der Parameter aktualisieren?".format(e_param.Name,text), ok=False, yes=True, no=True):
                            ParaCatSet = revit.app.Create.NewCategorySet()
                            for i in kate_excel:
                                if i in Categories_dict.keys():
                                    ParaCatSet.Insert(Categories_dict[i])
                                else:
                                    logger.error("{} ist keine Revit-Kategorie, bitte überprüfen.".format(i))
                            if ParaCatSet.IsEmpty:
                                continue
                            binding = app.Create.NewTypeBinding(ParaCatSet)
                            if projectparams[e_param.GUID][4] == 'Exemplar':
                                binding = app.Create.NewInstanceBinding(ParaCatSet)
                            if e_param.ParamGroup in ParamGroup_dict.keys():
                                dmap.ReInsert(projeparam_defis[e_param.GUID], binding, ParamGroup_dict[e_param.ParamGroup])
                                logger.info('Parameter {} wird aktualisiert'.format(e_param.Name))
                                try:
                                    cell_0 = get_cell(e_param.row,0)
                                    cell_1 = get_cell(e_param.row,1)
                                    cell_0.SetCellValue('')
                                    cell_1.SetCellValue('Projektparameter')

                                except:
                                    pass
                            else:
                                logger.error("{} ist kein Revit-Parametergroup, bitte überprüfen.".format(e_param.ParamGroup))
                            continue
                try:
                    paramtype = ParamType_dict[e_param.Disziplin][e_param.ParamType]
                except:
                    logger.error('Parameter {} konnte nicht erstellt werden wegen ungültigen Disziplin und Parametertype'.format(e_param.Name))
                    continue
                if e_param.Name in sharedparam_defis.keys():
                    exdefinition = sharedparam_defis[e_param.Name]
                    if e_param.GUID:
                        if e_param.GUID != exdefinition.GUID.ToString():
                            logger.error('Parameter {} konnte nicht erstellt werden wegen Namen-Konflikt'.format(e_param.Name))
                            continue
                    
                    if paramtype.ToString() != exdefinition.ParameterType.ToString():
                        logger.error('Parameter {} bereits vorhanden, aber der Parametertyp vorhandernes Parameter passt nicht zum Parametertyp von Excel'.format(e_param.Name))
                        continue
                            
                    logger.info('Shared Parameter {} vorhanden'.format(e_param.Name))
                    try:
                        cell_0 = get_cell(e_param.row,0)
                        cell_1 = get_cell(e_param.row,1)
                        cell_3 = get_cell(e_param.row,3)
                        cell_0.SetCellValue('')
                        cell_1.SetCellValue('Sharedparameter')
                        cell_3.SetCellValue(exdefinition.GUID.ToString())
                    except:
                        pass
                else:
                    file = app.OpenSharedParameterFile()
                    try:
                        parameterType = ParamType_dict[e_param.Disziplin][e_param.ParamType]
                    except:
                        parameterType = DB.ParameterType.Text
                    DefiCrea = DB.ExternalDefinitionCreationOptions(e_param.Name, parameterType)
                    if e_param.info:
                        DefiCrea.Description = e_param.info
                    if e_param.GUID:
                        DefiCrea.GUID = System.Guid(e_param.GUID)

                    exdefinition = file.Groups[e_param.Group].Definitions.Create(DefiCrea)
                    logger.info('Shared Parameter {} erstellt'.format(e_param.Name))
                    try:
                        cell_1 = get_cell(e_param.row,1)
                        cell_3 = get_cell(e_param.row,3)
                        cell_1.SetCellValue('Sharedparameter')
                        cell_3.SetCellValue(exdefinition.GUID.ToString())
                    except:
                        pass  
                try:
                    if e_param.Exemplar and e_param.Kategorien:
                        ParaCatSet = app.Create.NewCategorySet()
                        ParCatList = e_param.Kategorien.split(',')
                        for i in ParCatList:
                            if i:
                                while(i[0] == ' '):
                                    i = i[1:]
                                if i in Categories_dict.keys():
                                    ParaCatSet.Insert(Categories_dict[i])
                                else:
                                    logger.error('{} ist kein Revit-Kategorie'.format(i))
  
                    if e_param.Exemplar == 'Exemplar':
                        binding = app.Create.NewInstanceBinding(ParaCatSet)
                    else:
                        binding = app.Create.NewTypeBinding(ParaCatSet)
                    ParaGroup = DB.BuiltInParameterGroup.INVALID
                    if e_param.ParamGroup in ParamGroup_dict.keys():
                        ParaGroup = ParamGroup_dict[e_param.ParamGroup]

                    map = uiapp.ActiveUIDocument.Document.ParameterBindings
                    try:
                        map.Insert(exdefinition, binding, ParaGroup)
                        try:
                            cell_0 = get_cell(e_param.row,0)
                            cell_1 = get_cell(e_param.row,1)
                            cell_0.SetCellValue('')
                            cell_1.SetCellValue('Projektparameter')
                        except:
                            pass 
                        logger.info("ProjektParameter {} wird erstellt".format(e_param.Name))
                    except Exception as e:
                        logger.error(e)
                        logger.error("ProjektParameter {} konnte nicht erstellt werden".format(e_param.Name))
                except Exception as e:
                    logger.error(e)
                    logger.error("ProjektParameter {} konnte nicht erstellt werden".format(e_param.Name))

      
    else:
        sharedparams,sharedparam_defis = AllSharedParams_2022()
        projectparams,projeparam_defis = AllProjectParams_2022()
        ParamGroup_dict = AllParameterGroup_2022()
        Categories_dict = AllCategories()
        ParamType_dict = AllParameterType_2022()
        dmap = doc.ParameterBindings


        with forms.ProgressBar(title="{value}/{max_value} Parameter erstellt",cancellable=True, step=1) as pb1:

            for n,e_param in enumerate(Excel_Param_PLUS):
                if pb1.cancelled:
                    t.RollBack()
                    break
                pb1.update_progress(n+1,len(Excel_Param_PLUS))


                if e_param.GUID in projectparams.keys():
                    if e_param.Disziplin == projectparams[e_param.GUID][1] and e_param.ParamType == projectparams[e_param.GUID][2]:
                        logger.info("ProjektParameter {} vorhanden".format(e_param.Name))
                        try:
                            cell_0 = get_cell(e_param.row,0)
                            cell_1 = get_cell(e_param.row,1)
                            cell_0.SetCellValue('')
                            cell_1.SetCellValue('Projektparameter')
                        except:
                            pass
                        
                    else:
                        logger.error('Parameter {} konnte nicht erstellt werden wegen GUID-Konflikt'.format(e_param.Name))
                        continue
                    kate_excel_temp = e_param.Kategorien.split(',')
                    kate_excel = []
                    kate_revit = projectparams[e_param.GUID][5].split(',')
                    for i in kate_excel_temp:
                        if i:
                            while(i[0] == ' '):
                                i = i[1:]
                            kate_excel.append(i)
                    
                    if e_param.Exemplar != projectparams[e_param.GUID][4]:
                        logger.warning('Parameter {} vorhanden, aber Typ (Exemplar) passt nicht.'.format(e_param.Name))

                    if e_param.ParamGroup == projectparams[e_param.GUID][3] and sorted(kate_excel) == sorted(kate_revit):
                        continue
                    else:
                        if forms.alert("Parameter {} vorhanden, aber Kategorien/ Parametergroup passt nicht. Möchten Sie der Parameter aktualisieren?".format(e_param.Name), ok=False, yes=True, no=True):
                            ParaCatSet = revit.app.Create.NewCategorySet()
                            for i in kate_excel:
                                if i in Categories_dict.keys():
                                    ParaCatSet.Insert(Categories_dict[i])
                                else:
                                    logger.error("{} ist keine Revit-Kategorie, bitte überprüfen.".format(i))
                            if ParaCatSet.IsEmpty:
                                continue
                            binding = app.Create.NewTypeBinding(ParaCatSet)
                            if projectparams[e_param.GUID][4] == 'Exemplar':
                                binding = app.Create.NewInstanceBinding(ParaCatSet)
                            if e_param.ParamGroup in ParamGroup_dict.keys():
                                dmap.ReInsert(projeparam_defis[e_param.GUID], binding, ParamGroup_dict[e_param.ParamGroup])
                                logger.info('Parameter {} wird aktualisiert'.format(e_param.Name))
                                try:
                                    cell_0 = get_cell(e_param.row,0)
                                    cell_1 = get_cell(e_param.row,1)
                                    cell_0.SetCellValue('')
                                    cell_1.SetCellValue('Projektparameter')

                                except:
                                    pass
                            else:
                                logger.error("{} ist kein Revit-Parametergroup, bitte überprüfen.".format(e_param.ParamGroup))
                            continue

                try:
                    paramtype = AllParameterType_2022()[e_param.Disziplin][e_param.ParamType]
                except:
                    logger.error('Parameter {} konnte nicht erstellt werden wegen ungültigen Disziplin und Parametertype'.format(e_param.Name))
                    continue
                if e_param.Name in sharedparam_defis.keys():
                    exdefinition = sharedparam_defis[e_param.Name]
                    if e_param.GUID:
                        if e_param.GUID != exdefinition.GUID.ToString():
                            logger.error('Parameter {} konnte nicht erstellt werden wegen Namen-Konflikt'.format(e_param.Name))
                            continue
                    if paramtype.TypeId != exdefinition.GetDataType().TypeId:
                        logger.error('Parameter {} bereits vorhanden, aber der Parametertyp vorhandenes Parameter passt nicht zum Parametertyp von Excel'.format(e_param.Name))
                        continue
                            
                    logger.info('Shared Parameter {} vorhanden'.format(e_param.Name))
                    try:
                        cell_0 = get_cell(e_param.row,0)
                        cell_1 = get_cell(e_param.row,1)
                        cell_3 = get_cell(e_param.row,3)
                        cell_0.SetCellValue('')
                        cell_1.SetCellValue('Sharedparameter')
                        cell_3.SetCellValue(exdefinition.GUID.ToString())
                    except:
                        pass
                else:
                    file = app.OpenSharedParameterFile()
                    try:
                        parameterType = ParamType_dict[e_param.Disziplin][e_param.ParamType]
                    except:
                        parameterType = DB.SpecTypeId.String.Text
                    DefiCrea = DB.ExternalDefinitionCreationOptions(e_param.Name, parameterType)
                    if e_param.info:
                        DefiCrea.Description = e_param.info
                    if e_param.GUID:
                        DefiCrea.GUID = System.Guid(e_param.GUID)

                    exdefinition = file.Groups[e_param.Group].Definitions.Create(DefiCrea)
                    logger.info('Shared Parameter {} erstellt'.format(e_param.Name))
                    try:
                        cell_1 = get_cell(e_param.row,1)
                        cell_3 = get_cell(e_param.row,3)
                        cell_1.SetCellValue('Sharedparameter')
                        cell_3.SetCellValue(exdefinition.GUID.ToString())
                    except:
                        pass  
                try:
                    if e_param.Exemplar and e_param.Kategorien:
                        ParaCatSet = app.Create.NewCategorySet()
                        ParCatList = e_param.Kategorien.split(',')
                        for i in ParCatList:
                            if i:
                                while(i[0] == ' '):
                                    i = i[1:]
                                if i in Categories_dict.keys():
                                    ParaCatSet.Insert(Categories_dict[i])
                                else:
                                    logger.error('{} ist kein Revit-Kategorie'.format(i))
  
                    if e_param.Exemplar == 'Exemplar':
                        binding = app.Create.NewInstanceBinding(ParaCatSet)
                    else:
                        binding = app.Create.NewTypeBinding(ParaCatSet)
                    ParaGroup = DB.BuiltInParameterGroup.INVALID
                    if e_param.ParamGroup in ParamGroup_dict.keys():
                        ParaGroup = ParamGroup_dict[e_param.ParamGroup]

                    map = uiapp.ActiveUIDocument.Document.ParameterBindings
                    try:
                        map.Insert(exdefinition, binding, ParaGroup)
                        try:
                            cell_0 = get_cell(e_param.row,0)
                            cell_1 = get_cell(e_param.row,1)
                            cell_0.SetCellValue('')
                            cell_1.SetCellValue('Projektparameter')
                        except:
                            pass 
                        logger.info("ProjektParameter {} wird erstellt".format(e_param.Name))
                    except Exception as e:
                        logger.error(e)
                        logger.error("ProjektParameter {} konnte nicht erstellt werden".format(e_param.Name))
                except Exception as e:
                    logger.error(e)
                    logger.error("ProjektParameter {} konnte nicht erstellt werden".format(e_param.Name))

                    
            
if t.HasStarted():
    doc.Regenerate()
    t.Commit()

if Excel_Param_MINUS:  
    t.Start('Parameter löschen')   
    with forms.ProgressBar(title="{value}/{max_value} Parameter löschen",cancellable=True, step=1) as pb1:
        _params = DB.FilteredElementCollector(doc).OfClass(DB.ParameterElement).WhereElementIsNotElementType()
        _paramids = _params.ToElementIds()
        _params.Dispose()
        _ParamDict = {}
        for el in _paramids:
            _ParamDict[doc.GetElement(el).GetDefinition().Name] = el

        for n1,e_param in enumerate(Excel_Param_MINUS):
            if pb1.cancelled:
                t.RollBack()
                script.exit()
            pb1.update_progress(n1+1,len(Excel_Param_MINUS))
            if e_param.Name in _ParamDict.keys():
                if forms.alert("Parameter {} löschen?".format(e_param.Name), ok=False, yes=True, no=True):
                    doc.Delete(_ParamDict[e_param.Name])
                    logger.info("ProjektParameter {} wird gelöscht".format(e_param.Name))
                    try:
                        cell_0 = get_cell(e_param.row,0)
                        cell_1 = get_cell(e_param.row,1)
                        cell_0.SetCellValue('')
                        cell_1.SetCellValue('Sharedparameter')
                    except:
                        pass
            else:
                logger.info("ProjektParameter {} nicht vorhanden".format(e_param.Name))
                try:
                    cell_0 = get_cell(e_param.row,0)
                    cell_1 = get_cell(e_param.row,1)
                    cell_0.SetCellValue('')
                    cell_1.SetCellValue('Sharedparameter')
                except:
                    pass

if t.HasStarted():
    t.Commit()

fs = _NPOI.FileStream(excelPath, _NPOI.FileMode.Create, _NPOI.FileAccess.Write)
book1.Write(fs)
book1.Close()
fs.Close()
