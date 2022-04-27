# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_forms import ExcelSuche
from IGF_libKopie import AllProjectParams_2020,AllSharedParams_2020,CreateSharedParam_2020,CreateProjFromSharedParam_2020,AllParameterType_2020,AllCategories,AllParameterGroup_2020
from IGF_libKopie import AllProjectParams_2022,AllSharedParams_2022,CreateSharedParam_2022,CreateProjFromSharedParam_2022,AllParameterType_2022,AllParameterGroup_2022
from pyrevit import script, forms
from rpw import revit,DB
import os

import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal


__title__ = "Import Parameter von Excel" 
__doc__ = """erstellt Shared- und ProjektParameter von excel. 

+: erstellen
-: löschen

[2021.10.08]
Version: 2.0"""
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

projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number

config = script.get_config('Parametererstellung-' + projectinfo)
adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.ShowDialog()
try:
    config.adresse = ExcelWPF.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel gegeben')
    script.exit()

# Excel lesen
# Parameter Klass aus Excel
class E_Parameter:
    def __init__(self):
        self.Name = ''
        self.GUID = ''
        self.Disziplin = ''
        self.ParamType = ''
        self.ParamGroup = ''
        self.Kategorien = ''
        self.Exemplar = ''
        self.Group = ''
        self.info = ''
    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self,value):
        self._Name = value
    
    @property
    def GUID(self):
        return self._GUID
    @GUID.setter
    def GUID(self,value):
        self._GUID = value
    
    @property
    def Disziplin(self):
        return self._Disziplin
    @Disziplin.setter
    def Disziplin(self,value):
        self._Disziplin = value
    
    @property
    def ParamType(self):
        return self._ParamType
    @ParamType.setter
    def ParamType(self,value):
        self._ParamType = value
    
    @property
    def ParamGroup(self):
        return self._ParamGroup
    @ParamGroup.setter
    def ParamGroup(self,value):
        self._ParamGroup = value
    
    @property
    def Kategorien(self):
        return self._Kategorien
    @Kategorien.setter
    def Kategorien(self,value):
        self._Kategorien = value
    
    @property
    def Exemplar(self):
        return self._Exemplar
    @Exemplar.setter
    def Exemplar(self,value):
        self._Exemplar = value
    
    @property
    def Group(self):
        return self._Group
    @Group.setter
    def Group(self,value):
        self._Group = value
    
    @property
    def info(self):
        return self._info
    @info.setter
    def info(self,value):
        self._info = value

excelPath = config.adresse
exapplication = ex.ApplicationClass()

book1 = exapplication.Workbooks.Open(excelPath)
E_Parameter_Dict = {}
E_Parameter_Dict_del = {}
GroupList = []

sheetscount = book1.Worksheets.Count
n = 0
for sheet in book1.Worksheets:
    n += 1
    Group = sheet.Name
    if Group in ['Hinweis','Keine Gruppe','Revit (Original)']:
        continue
    with forms.ProgressBar(title="Excel lesen",cancellable=True, step=1) as pb:
        pb.title = str(n) + '/' + str(sheetscount) + " Excel lesen -- " + Group + " -- {value}/{max_value} Parameter"
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            if pb.cancelled:
                Marshal.FinalReleaseComObject(sheet)
                Marshal.FinalReleaseComObject(book1)
                exapplication.Quit()
                Marshal.FinalReleaseComObject(exapplication)
                script.exit()
            pb.update_progress(row-1, rows)
            if sheet.Cells[row, 1].Value2 == '+':
                param = E_Parameter()
                param.Name = sheet.Cells[row, 3].Value2
                param.GUID = sheet.Cells[row, 4].Value2
                param.Disziplin = sheet.Cells[row, 5].Value2
                param.ParamType = sheet.Cells[row, 6].Value2
                param.ParamGroup = sheet.Cells[row, 7].Value2
                param.Kategorien = sheet.Cells[row, 8].Value2
                param.Exemplar = sheet.Cells[row, 9].Value2
                param.Group = sheet.Cells[row, 12].Value2
                param.info = sheet.Cells[row, 13].Value2
                E_Parameter_Dict[param] = [Group,row]
                if sheet.Cells[row, 12].Value2:
                    if not sheet.Cells[row, 12].Value2 in GroupList:
                        GroupList.append(sheet.Cells[row, 12].Value2)
            elif sheet.Cells[row, 1].Value2 == '-':
                param = E_Parameter()
                param.Name = sheet.Cells[row, 3].Value2
                param.GUID = sheet.Cells[row, 4].Value2
                param.Disziplin = sheet.Cells[row, 5].Value2
                param.ParamType = sheet.Cells[row, 6].Value2
                param.ParamGroup = sheet.Cells[row, 7].Value2
                param.Kategorien = sheet.Cells[row, 8].Value2
                param.Exemplar = sheet.Cells[row, 9].Value2
                param.Group = sheet.Cells[row, 12].Value2
                param.info = sheet.Cells[row, 13].Value2
                E_Parameter_Dict_del[param] = [Group,row]


book1.Save()
book1.Close()


filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()

Groups = [i.Name for i in file.Groups]
for i in GroupList:
    if not i in Groups:
        file.Groups.Create(i)

E_Parameter_Liste = []
E_Parameter_Liste_del = []

try:
    E_Parameter_Liste = E_Parameter_Dict.keys()[:]
except Exception as e:
    logger.error(e)

try:
    E_Parameter_Liste_del = E_Parameter_Dict_del.keys()[:]
except Exception as e:
    logger.error(e)

book2 = exapplication.Workbooks.Open(excelPath)   
t = DB.Transaction(doc)
# SharedParameter, ProjektParameter
if E_Parameter_Liste:
    t.Start('Parameter erstellen')
    if revitversion == '2020':
        sharedparams,sharedparam_defis = AllSharedParams_2020()
        projectparams,projeparam_defis = AllProjectParams_2020()
        ParamGroup_dict = AllParameterGroup_2020()
        ParamType_dict = AllParameterType_2020()
        Categories_dict = AllCategories()
        dmap = doc.ParameterBindings

        with forms.ProgressBar(title="{value}/{max_value} Parameter erstellt",cancellable=True, step=1) as pb1:
            for n1,e_param in enumerate(E_Parameter_Liste):
                if pb1.cancelled:
                    t.RollBack()
                    book2.Close(SaveChanges = False)
                    break
                pb1.update_progress(n1+1,len(E_Parameter_Liste))

                if e_param.GUID in projectparams.keys():
                    if e_param.Disziplin == projectparams[e_param.GUID][1] and e_param.ParamType == projectparams[e_param.GUID][2]:
                        logger.info("ProjektParameter {} vorhanden".format(e_param.Name))
                        try:
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Projektparameter'
                        except:
                            pass
  
                    else:
                        logger.error('Parameter {} konnte nicht erstellt werden wegen GUID-Konflikt'.format(e_param.Name))
                        continue
                    kate_excel_temp = e_param.Kategorien.Split(',')
                    kate_excel = []
                    kate_revit = projectparams[e_param.GUID][5].Split(',')
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
                            binding = app.Create.NewTypeBinding(ParaCatSet)
                            if projectparams[e_param.GUID][4] == 'Exemplar':
                                binding = app.Create.NewInstanceBinding(ParaCatSet)
                            dmap.ReInsert(projeparam_defis[e_param.GUID], binding, ParamGroup_dict[e_param.ParamGroup])
                            logger.info('Parameter {} wird aktualisiert'.format(e_param.Name))
                            try:
                                book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                                book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Projektparameter'
                            except:
                                pass
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
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Sharedparameter'
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],4] = exdefinition.GUID.ToString()
                    except:
                        pass
                else:
                    exdefinition = CreateSharedParam_2020(Name=e_param.Name, Disziplin=e_param.Disziplin,Typ=e_param.ParamType,GUID=e_param.GUID,Info=e_param.info,Gruppe=e_param.Group)
                    logger.info('Shared Parameter {} erstellt'.format(e_param.Name))
                    try:
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Sharedparameter'
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],4] = exdefinition.GUID.ToString()
                    except:
                        pass           
                try:
                    if e_param.Exemplar and e_param.Kategorien:
                        CreateProjFromSharedParam_2020(ExternalDefinition=exdefinition,Gruppe=e_param.ParamGroup,Kategorien=e_param.Kategorien,Typ_Exemplar=e_param.Exemplar)
                        try:
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Projektparameter'
                        except:
                            pass 
                        logger.info("ProjektParameter {} wird erstellt".format(e_param.Name))
                    else:
                        logger.error("ProjektParameter {} konnte nicht erstellt werden wegen unvollständigen Eingaben".format(e_param.Name))
                except Exception as e:
                    logger.error(e)
                    logger.error("ProjektParameter {} konnte nicht erstellt werden".format(e_param.Name))

        
    elif revitversion == '2022':
        sharedparams,sharedparam_defis = AllSharedParams_2022()
        projectparams,projeparam_defis = AllProjectParams_2022()
        ParamGroup_dict = AllParameterGroup_2022()
        Categories_dict = AllCategories()
        dmap = doc.ParameterBindings


        with forms.ProgressBar(title="{value}/{max_value} Parameter erstellt",cancellable=True, step=1) as pb1:

            for n1,e_param in enumerate(E_Parameter_Liste):
                if pb1.cancelled:
                    t.RollBack()
                    book2.Close(SaveChanges = False)
                    break
                pb1.update_progress(n1+1,len(E_Parameter_Liste))


                if e_param.GUID in projectparams.keys():
                    if e_param.Disziplin == projectparams[e_param.GUID][1] and e_param.ParamType == projectparams[e_param.GUID][2]:
                        logger.info("ProjektParameter {} vorhanden".format(e_param.Name))
                        try:
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Projektparameter'
                        except:
                            pass
                        
                    else:
                        logger.error('Parameter {} konnte nicht erstellt werden wegen GUID-Konflikt'.format(e_param.Name))
                        continue
                    kate_excel_temp = e_param.Kategorien.Split(',')
                    kate_excel = []
                    kate_revit = projectparams[e_param.GUID][5].Split(',')
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
                            binding = app.Create.NewTypeBinding(ParaCatSet)
                            if projectparams[e_param.GUID][4] == 'Exemplar':
                                binding = app.Create.NewInstanceBinding(ParaCatSet)
                            dmap.ReInsert(projeparam_defis[e_param.GUID], binding, ParamGroup_dict[e_param.ParamGroup])
                            logger.info('Parameter {} wird aktualisiert'.format(e_param.Name))
                            try:
                                book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                                book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Projektparameter'
                            except:
                                pass
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
                    
                    if paramtype.TypeId.ToString() != exdefinition.GetDataType().TypeId.ToString():
                        logger.error('Parameter {} bereits vorhanden, aber der Parametertyp vorhandenes Parameter passt nicht zum Parametertyp von Excel'.format(e_param.Name))
                        continue
                            
                    logger.info('Shared Parameter {} vorhanden'.format(e_param.Name))
                    try:
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Sharedparameter'
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],4] = exdefinition.GUID.ToString()
                    except:
                        pass
                else:
                    exdefinition = CreateSharedParam_2022(Name=e_param.Name, Disziplin=e_param.Disziplin,Typ=e_param.ParamType,GUID=e_param.GUID,Info=e_param.info,Gruppe=e_param.Group)
                    logger.info('Shared Parameter {} erstellt'.format(e_param.Name))
                    doc.Regenerate()
                    try:
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Sharedparameter'
                        book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],4] = exdefinition.GUID.ToString()
                    except:
                        pass          
                try:
                    if e_param.Exemplar and e_param.Kategorien:
                        CreateProjFromSharedParam_2022(ExternalDefinition=exdefinition,Gruppe=e_param.ParamGroup,Kategorien=e_param.Kategorien,Typ_Exemplar=e_param.Exemplar)
                        
                        logger.info("ProjektParameter {} wird erstellt".format(e_param.Name))
                        try:
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],1] = ''
                            book2.Worksheets[E_Parameter_Dict[e_param][0]].Cells[E_Parameter_Dict[e_param][1],2] = 'Projektparameter'
                        except:
                            pass
                    else:
                        logger.error("ProjektParameter {} konnte nicht erstellt werden wegen unvollständigen Eingaben".format(e_param.Name))
                except Exception as e:
                    logger.error(e)
                    logger.error("ProjektParameter {} konnte nicht erstellt werden".format(e_param.Name))
            

if t.HasStarted():
    doc.Regenerate()
    t.Commit()
book2.Save()
book2.Close()


book3 = exapplication.Workbooks.Open(excelPath)   
if E_Parameter_Liste_del:  
    t.Start('Parameter löschen')   
    with forms.ProgressBar(title="{value}/{max_value} Parameter löschen",cancellable=True, step=1) as pb1:
        _params = DB.FilteredElementCollector(doc).OfClass(DB.ParameterElement).WhereElementIsNotElementType()
        _paramids = _params.ToElementIds()
        _params.Dispose()
        _ParamDict = {}
        for el in _paramids:
            _ParamDict[doc.GetElement(el).GetDefinition().Name] = el

        for n1,e_param in enumerate(E_Parameter_Liste_del):
            if pb1.cancelled:
                t.RollBack()
                book3.Close(SaveChanges = False)
                script.exit()
            pb1.update_progress(n1+1,len(E_Parameter_Liste_del))
            if e_param.Name in _ParamDict.keys():
                if forms.alert("Parameter {} löschen?".format(e_param.Name), ok=False, yes=True, no=True):
                    doc.Delete(_ParamDict[e_param.Name])
                    logger.info("ProjektParameter {} wird gelöscht".format(e_param.Name))
                    try:
                        book3.Worksheets[E_Parameter_Dict_del[e_param][0]].Cells[E_Parameter_Dict_del[e_param][1],1] = ''
                        book3.Worksheets[E_Parameter_Dict_del[e_param][0]].Cells[E_Parameter_Dict_del[e_param][1],2] = 'Sharedparameter'
                    except:
                        pass
            else:
                logger.info("ProjektParameter {} nicht vorhanden".format(e_param.Name))
                try:
                    book3.Worksheets[E_Parameter_Dict_del[e_param][0]].Cells[E_Parameter_Dict_del[e_param][1],1] = ''
                    book3.Worksheets[E_Parameter_Dict_del[e_param][0]].Cells[E_Parameter_Dict_del[e_param][1],2] = 'Sharedparameter'
                except:
                    pass

if t.HasStarted():
    t.Commit()

book3.Save()
book3.Close()