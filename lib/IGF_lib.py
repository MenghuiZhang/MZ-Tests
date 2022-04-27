# coding: utf8
from logging import exception
import System
from System.IO import Path, File
from rpw import revit,DB,UI

######################################################################################
def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if param.StorageType.ToString() == 'ElementId':
        return param.AsValueString()
    elif param.StorageType.ToString() == 'Integer':
        value = param.AsInteger()
    elif param.StorageType.ToString() == 'Double':
        value = param.AsDouble()
    elif param.StorageType.ToString() == 'String':
        value = param.AsString()
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value
######################################################################################


def AllBuiltinCategory():
    _builtInCategory = {}
    allcate = System.Enum.GetValues(DB.BuiltInCategory.OST_PointClouds.GetType())
    for cate in allcate:
        try:
            name = revit.doc.Settings.Categories.get_Item(cate).Name
            if name:
                _builtInCategory[name] = cate
        except:
            pass
    return _builtInCategory


def AllCategories():
    _Categories = {}
    for cat in revit.doc.Settings.Categories:
        name = cat.Name
        _Categories[name] = cat
    return _Categories


######################################################################################
# Parameter
# Revit 2020

def DeiziplinErmitteln_2020(inParaTypName,inParaTyp):
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


def AllParameterType_2020():
    _ParameterType = {}
    _ParameterType['Allgemein'] = {}
    _ParameterType['Energie'] = {}
    _ParameterType['Tragwerk'] = {}
    _ParameterType['Lüftung'] = {}
    _ParameterType['Rohre'] = {}
    _ParameterType['Elektro'] = {}

    alltype = System.Enum.GetValues(DB.ParameterType().GetType())
    for i in alltype:
        Type = i.ToString()
        if Type in ['Invalid','FamilyType']:
            continue
        type_user = DB.LabelUtils.GetLabelFor(i)
        dis = DeiziplinErmitteln_2020(type_user,Type)
        _ParameterType[dis][type_user] = i

    return _ParameterType

def AllParameterGroup_2020():
    _parametergroup = {}
    allGro = System.Enum.GetValues(DB.BuiltInParameterGroup().GetType())
    for i in allGro:
        name = DB.LabelUtils.GetLabelFor(i)
        _parametergroup[name] = i
    return _parametergroup


def AllSharedParams_2020():
    aktuellSharedPara = {}
    aktuellSharedPara_defi = {}
    file = revit.app.OpenSharedParameterFile()
    if file.Groups:
        for dg in file.Groups:
            dgName = dg.Name
            aktuellSharedPara[dgName] = {}
            if dg.Definitions:
                for d in dg.Definitions:
                    name = d.Name
                    info = d.Description
                    GUID = d.GUID.ToString()
                    Type = d.ParameterType.ToString()
                    type_user = DB.LabelUtils.GetLabelFor(d.ParameterType)
                    dis = DeiziplinErmitteln_2020(type_user,Type)
                    aktuellSharedPara[dgName][name] = [GUID,dis,type_user,info]
                    aktuellSharedPara_defi[name] = d

    sorted(aktuellSharedPara)
    return aktuellSharedPara,aktuellSharedPara_defi

def AllSharedParamsmitGUID_2020():
    aktuellSharedPara = {}
    file = revit.app.OpenSharedParameterFile()
    if file.Groups:
        for dg in file.Groups:
            dgName = dg.Name
            if dg.Definitions:
                for d in dg.Definitions:
                    info = d.Description
                    GUID = d.GUID.ToString()
                    aktuellSharedPara[GUID] = [info,dgName]
    return aktuellSharedPara

def AllProjectParams_2020():
    _SharedParams = {}
    param_coll = DB.FilteredElementCollector(revit.doc).OfClass(DB.SharedParameterElement).WhereElementIsNotElementType()
    paramids = param_coll.ToElementIds()
    param_coll.Dispose()
    for paramid in paramids:
        param = revit.doc.GetElement(paramid)
        definition = param.GetDefinition()
        _SharedParams[definition.Id.ToString()] = param.GuidValue.ToString()

    map = revit.doc.ParameterBindings
    dep = map.ForwardIterator()
    _ProjectParams = {}
    _ProjectParam_defis = {}
    GUID_Daten = AllSharedParamsmitGUID_2020()
    while(dep.MoveNext()):
        try:
            neu_definition = dep.Key
            Name = neu_definition.Name
            Paratyp = neu_definition.ParameterType.ToString()
            TypName = DB.LabelUtils.GetLabelFor(neu_definition.ParameterType)
            dis = DeiziplinErmitteln_2020(TypName,Paratyp)
            Group = DB.LabelUtils.GetLabelFor(neu_definition.ParameterGroup)
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
                    cateName = cateName + cate.Name + ','
            try:
                guid = _SharedParams[neu_definition.Id.ToString()]
            except:
                continue
            info = ''
            dgGroup = ''
            if guid in GUID_Daten.keys():
                info = GUID_Daten[guid][0]
                dgGroup = GUID_Daten[guid][1]
            
            _ProjectParams[guid] = [Name,dis,TypName,Group,typOrex,cateName[:-1],dgGroup,info]
            _ProjectParam_defis[guid] = neu_definition
            
        except:
            pass
    return _ProjectParams,_ProjectParam_defis

def CreateDefinition_2020(Name = None, Disziplin = None, Typ = None, Info = None, GUID = None):
    try:
        parameterType = AllParameterType_2020()[Disziplin][Typ]
    except:
        parameterType = DB.ParameterType.Text
    DefiCrea = DB.ExternalDefinitionCreationOptions(Name, parameterType)
    if Info:
        DefiCrea.Description = Info
    if GUID:
        DefiCrea.GUID = System.Guid(GUID)
    return DefiCrea

def CreateSharedParam_2020(Gruppe = 'TemporaryDefintionGroup', Name = None, Disziplin = None, Typ = None, Info = None, GUID = None):
    app = revit.app
    file = app.OpenSharedParameterFile()
    DefiCrea = CreateDefinition_2020(Name = Name, Disziplin = Disziplin, Typ = Typ, Info = Info, GUID = GUID)
    try:
        sharedPara = file.Groups[Gruppe].Definitions.Create(DefiCrea)
        return sharedPara
    except:
        try:
            sharedPara = file.Groups.Create(Gruppe).Definitions.Create(DefiCrea)
            return sharedPara
        except:
            return


def CreateProjParam_2020(Name = None, Disziplin = None, Parametertyp = None, Info = None, Gruppe = None, Kategorien = 'Allgemeines Modell', Typ_Exemplar = 'Typ'):
    app = revit.app
    filename = app.SharedParametersFilename
    tempFile = Path.GetTempFileName() + ".txt"
    File.Create(tempFile)
    app.SharedParametersFilename = tempFile
    ExternalDefinition = CreateSharedParam_2020(Name = Name, Disziplin = Disziplin, Typ = Parametertyp, Info = Info)

    app.SharedParametersFilename = filename
    File.Delete(tempFile)

    ParaCatSet = revit.app.Create.NewCategorySet()
    ParCatList = Kategorien.Split(',')

    for i in ParCatList:
        if i:

            while(i[0] == ' '):
                i = i[1:]
            if i in AllCategories().keys():
                ParaCatSet.Insert(AllCategories()[i])

    binding = app.Create.NewTypeBinding(ParaCatSet)
    if Typ_Exemplar == 'Exemplar':
        binding = app.Create.NewInstanceBinding(ParaCatSet)
    else:
        binding = app.Create.NewTypeBinding(ParaCatSet)

    ParaGroup = DB.BuiltInParameterGroup.INVALID
    if Gruppe in AllParameterGroup_2020().keys():
        ParaGroup = AllParameterGroup_2020()[Gruppe]

    map = revit.uiapp.ActiveUIDocument.Document.ParameterBindings
    try:
        map.Insert(ExternalDefinition, binding, ParaGroup)
    except Exception as e:
        print(e)


def CreateProjFromSharedParam_2020(ExternalDefinition = None, Gruppe = None, Kategorien = None, Typ_Exemplar = None):
    uiapp = revit.uiapp
    app = revit.app

    if not ExternalDefinition:
        return
    ParaCatSet = app.Create.NewCategorySet()
    ParCatList = Kategorien.Split(',')

    for i in ParCatList:
        if i:
            while(i[0] == ' '):
                i = i[1:]
            if i in AllCategories().keys():
                ParaCatSet.Insert(AllCategories()[i])

    binding = app.Create.NewTypeBinding(ParaCatSet)
    if Typ_Exemplar == 'Exemplar':
        binding = app.Create.NewInstanceBinding(ParaCatSet)
    else:
        binding = app.Create.NewTypeBinding(ParaCatSet)

    ParaGroup = DB.BuiltInParameterGroup.INVALID
    if Gruppe in AllParameterGroup_2020().keys():
        ParaGroup = AllParameterGroup_2020()[Gruppe]

    map = uiapp.ActiveUIDocument.Document.ParameterBindings
    try:
        map.Insert(ExternalDefinition, binding, ParaGroup)
    except:
        return



#############################################################################################
# Revit 2021/2022
def AllParameterGroup_2022():
    _parametergroup = {}
    for i in DB.ParameterUtils.GetAllBuiltInGroups():
        name = DB.LabelUtils.GetLabelForGroup(i)
        _parametergroup[name] = i
    return _parametergroup

def AllParameterType_2022():
    # Allgemien: "Kosten pro Fläche", "Abstand", 'Drehwinkel'
    _ParameterType = {}
    _ParameterType['Elektro'] = {}
    _ParameterType['Energie'] = {}
    _ParameterType['HLK'] = {}
    _ParameterType['Infrastruktur'] = {}
    _ParameterType['Rohre'] = {}
    _ParameterType['Tragwerk'] = {}
    _ParameterType['Allgemein'] = {}
    for i in DB.SpecUtils.GetAllSpecs():
        name = DB.LabelUtils.GetLabelForSpec(i)

        if i.TypeId.ToString().find('autodesk.spec.aec:numberOfPoles') != -1:
            _ParameterType['Elektro'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec.electrical') != -1:
            _ParameterType['Elektro'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec.energy') != -1:
            _ParameterType['Energie'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec.hvac') != -1:
            _ParameterType['HLK'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec.infrastructure') != -1:
            _ParameterType['Infrastruktur'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec.piping') != -1:
            _ParameterType['Rohre'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec.structural') != -1:
            _ParameterType['Tragwerk'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.aec:') != -1:
            _ParameterType['Allgemein'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.measurable:') != -1:
            _ParameterType['Allgemein'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.string') != -1:
            _ParameterType['Allgemein'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec:') != -1:
            _ParameterType['Allgemein'][name] = i
        elif i.TypeId.ToString().find('autodesk.spec.reference') != -1:
            _ParameterType['Allgemein'][name] = i
    
    return _ParameterType

def AllParameterTypeMitDis_2022():
    # Allgemien: "Kosten pro Fläche", "Abstand", 'Drehwinkel'
    _ParameterType = {}

    for i in DB.SpecUtils.GetAllSpecs():
        name = DB.LabelUtils.GetLabelForSpec(i)
        typeid = i.TypeId.ToString()
        if typeid.find('-') != -1:
            typeid = typeid[:typeid.find('-')]

        
        if typeid.find('autodesk.spec.aec:numberOfPoles') != -1:
            _ParameterType[typeid] = ['Elektro',name]
        elif typeid.find('autodesk.spec.aec.electrical') != -1:
            _ParameterType[typeid] = ['Elektro',name]
        elif typeid.find('autodesk.spec.aec.energy') != -1:
            _ParameterType[typeid] = ['Energie',name] 
        elif typeid.find('autodesk.spec.aec.hvac') != -1:
            _ParameterType[typeid] = ['HLK',name]
        elif typeid.find('autodesk.spec.aec.infrastructure') != -1:
            _ParameterType[typeid] = ['Infrastruktur',name]
        elif typeid.find('autodesk.spec.aec.piping') != -1:
            _ParameterType[typeid] = ['Rohre',name]
        elif typeid.find('autodesk.spec.aec.structural') != -1:
            _ParameterType[typeid] = ['Tragwerk',name]
        else:
            _ParameterType[typeid] = ['Allgemein',name]

    return _ParameterType

def AllSharedParams_2022():
    aktuellSharedPara = {}
    aktuellSharedPara_defi = {}
    typedict = AllParameterTypeMitDis_2022()

    file = revit.app.OpenSharedParameterFile()
    if file.Groups:
        for dg in file.Groups:
            dgName = dg.Name
            aktuellSharedPara[dgName] = {}
            if dg.Definitions:
                for d in dg.Definitions:
                    name = d.Name
                    info = d.Description
                    guid = d.GUID.ToString()
                    Type = d.GetDataType().TypeId.ToString()
                    if Type.find('-') != -1:
                        Type = Type[:Type.find('-')]
                    type_user = typedict[Type][1]
                    dis = typedict[Type][0]
                    aktuellSharedPara[dgName][name] = [guid,dis,type_user,info]
                    aktuellSharedPara_defi[name] = d

    sorted(aktuellSharedPara)
    return aktuellSharedPara,aktuellSharedPara_defi

def AllSharedParamsmitGUID_2022():
    aktuellSharedPara = {}
    file = revit.app.OpenSharedParameterFile()
    if file.Groups:
        for dg in file.Groups:
            dgName = dg.Name
            if dg.Definitions:
                for d in dg.Definitions:
                    info = d.Description
                    guid = d.GUID.ToString()
                    aktuellSharedPara[guid] = [info,dgName]
    return aktuellSharedPara

def AllProjectParams_2022():
    _SharedParams = {}
    param_coll = DB.FilteredElementCollector(revit.doc).OfClass(DB.SharedParameterElement).WhereElementIsNotElementType()
    paramids = param_coll.ToElementIds()
    param_coll.Dispose()
    for paramid in paramids:
        param = revit.doc.GetElement(paramid)
        definition = param.GetDefinition()
        _SharedParams[definition.Id.ToString()] = param
    map = revit.doc.ParameterBindings
    dep = map.ForwardIterator()
    _ProjectParams = {}
    _ProjectParam_defis = {}
    typedict = AllParameterTypeMitDis_2022()
    GUID_Daten = AllSharedParamsmitGUID_2022()

    while(dep.MoveNext()):
        try:
            neu_definition = dep.Key
            Def_id = neu_definition.Id.ToString()
            if Def_id in _SharedParams.keys():
                SharedPara = _SharedParams[Def_id]
                Definition = SharedPara.GetDefinition()
                Name = Definition.Name
                Paratyp = Definition.GetDataType().TypeId.ToString()
                if Paratyp.find('-') != -1:
                    Paratyp = Paratyp[:Paratyp.find('-')]
                TypName = typedict[Paratyp][1]
                dis = typedict[Paratyp][0]
                
                Group = DB.LabelUtils.GetLabelForGroup(Definition.GetGroupTypeId())
                guid = SharedPara.GuidValue.ToString()
                
                Binding = dep.Current
                Type = Binding.GetType().ToString()
                cateName = ''
                info = ''
                dgGroup = ''
                if guid in GUID_Daten.keys():
                    info = GUID_Daten[guid][0]
                    dgGroup = GUID_Daten[guid][1]
                typOrex = None
                if Type == 'Autodesk.Revit.DB.InstanceBinding':
                    typOrex = 'Exemplar'
                else:
                    typOrex = 'Type'
                if Binding:
                    cates = Binding.Categories
                    for cate in cates:
                        cateName = cateName + cate.Name + ','
                _ProjectParams[guid] = [Name,dis,TypName,Group,typOrex,cateName[:-1],dgGroup,info]
                _ProjectParam_defis[guid] = neu_definition
            else:
                continue
            
        except Exception as e:
            pass
    return _ProjectParams,_ProjectParam_defis


def CreateDefinition_2022(Name = None, Disziplin = None, Typ = None, Info = None, GUID = None):
    try:
        ForgeTypeId = AllParameterType_2022()[Disziplin][Typ]
    except:
        ForgeTypeId = DB.SpecTypeId.String.Text

    DefiCrea = DB.ExternalDefinitionCreationOptions(Name, ForgeTypeId)
    if Info:
        DefiCrea.Description = Info
    if GUID:
        DefiCrea.GUID = System.Guid(GUID)
    return DefiCrea

def CreateSharedParam_2022(Gruppe = 'TemporaryDefintionGroup', Name = None, Disziplin = None, Typ = None, Info = None, GUID = None):
    app = revit.app
    file = app.OpenSharedParameterFile()
    DefiCrea = CreateDefinition_2022(Name = Name, Disziplin = Disziplin, Typ = Typ, Info = Info, GUID = GUID)
    print(DefiCrea.GUID)
    print('True')
    try:
        sharedPara = file.Groups[Gruppe].Definitions.Create(DefiCrea)
        return sharedPara
    except:
        try:
            sharedPara = file.Groups.Create(Gruppe).Definitions.Create(DefiCrea)
            return sharedPara
        except:
            return 

def CreateProjParam_2022(Name = None, Disziplin = None, Parametertyp = None, Info = None, Gruppe = None, Kategorien = 'Allgemeines Modell', Typ_Exemplar = 'Typ'):
    app = revit.app
    filename = app.SharedParametersFilename
    tempFile = Path.GetTempFileName() + ".txt"
    File.Create(tempFile)
    app.SharedParametersFilename = tempFile
    ExternalDefinition = CreateSharedParam_2022(Name = Name, Disziplin = Disziplin, Typ = Parametertyp, Info = Info)
    app.SharedParametersFilename = filename
    File.Delete(tempFile)

    ParaCatSet = app.Create.NewCategorySet()
    ParCatList = Kategorien.Split(',')

    for i in ParCatList:
        if i:

            while(i[0] == ' '):
                i = i[1:]
            if i in AllCategories().keys():
                ParaCatSet.Insert(AllCategories()[i])

    binding = app.Create.NewTypeBinding(ParaCatSet)
    if Typ_Exemplar == 'Exemplar':
        binding = app.Create.NewInstanceBinding(ParaCatSet)
    else:
        binding = app.Create.NewTypeBinding(ParaCatSet)

    ParaGroup = None
    if Gruppe in AllParameterGroup_2022().keys():
        ParaGroup = AllParameterGroup_2022()[Gruppe]

    map = revit.uiapp.ActiveUIDocument.Document.ParameterBindings
    if ParaGroup:
        try:
            map.Insert(ExternalDefinition, binding, ParaGroup)
        except:
            return
    else:
        try:
            map.Insert(ExternalDefinition, binding)
        except:
            return


def CreateProjFromSharedParam_2022(ExternalDefinition = None, Gruppe = None, Kategorien = None, Typ_Exemplar = None):
    app = revit.app
    if not ExternalDefinition:
        return
    ParaCatSet = app.Create.NewCategorySet()
    ParCatList = Kategorien.Split(',')
    

    for i in ParCatList:
        if i:
            while(i[0] == ' '):
                i = i[1:]
            if i in AllCategories().keys():
                ParaCatSet.Insert(AllCategories()[i])
    

    binding = app.Create.NewTypeBinding(ParaCatSet)
    if Typ_Exemplar == 'Exemplar':
        binding = app.Create.NewInstanceBinding(ParaCatSet)
    else:
        binding = app.Create.NewTypeBinding(ParaCatSet)

    ParaGroup = None
    if Gruppe in AllParameterGroup_2022().keys():
        ParaGroup = AllParameterGroup_2022()[Gruppe]

    map = revit.doc.ParameterBindings
    if ParaGroup:
        map.Insert(ExternalDefinition, binding, ParaGroup)
       
    else:
        map.Insert(ExternalDefinition, binding)



#######################################################################
# Solid

def getSolids(elem):
    def GetSolid(GeoEle):
        lstSolid = []
        for el in GeoEle:
            if el.GetType().ToString() == 'Autodesk.Revit.DB.Solid':
                if el.SurfaceArea > 0 and el.Volume > 0 and el.Faces.Size > 1 and el.Edges.Size > 1:
                    lstSolid.append(el)
            elif el.GetType().ToString() == 'Autodesk.Revit.DB.GeometryInstance':
                ge = el.GetInstanceGeometry()
                lstSolid.extend(GetSolid(ge))
        return lstSolid

    lstSolid = []
    opt = DB.Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    ge = elem.get_Geometry(opt)
    if ge != None:
        lstSolid.extend(GetSolid(ge))
    return lstSolid

def TransformSolid(elem,revitlink):
    m_lstModels = []
    listSolids = getSolids(elem)
    for solid in listSolids:
        tempSolid = solid
        tempSolid = DB.SolidUtils.CreateTransformed(solid,revitlink.GetTransform())
        m_lstModels.append(tempSolid)
    return m_lstModels

#######################################################################

def Leitungslinie(elem):
    pipecurve = None
    conns = list(elem.ConnectorManager.Connectors)        
    pipecurve = DB.Line.CreateBound(conns[0].Origin, conns[1].Origin)
    return pipecurve

def pickElements(elementids):
    sel = revit.uidoc.Selection.GetElementIds()
    for el in elementids:
        sel.Add(el)
    revit.uidoc.Selection.SetElementIds(sel)

def Pickelem():
    try:
        re = revit.uidoc.Selection.PickObject(UI.Selection.ObjectType.Element)
        elem = revit.doc.GetElement(re)
        return elem
    except:
        pass

def AlleElementInAnsicht():
    '''Wählt alle Elmente des angeklickten Typs im gesamten Projekt aus und gibt den Familien- und Typnamen zurück'''

    element = Pickelem()
    idListe = []
    familie = None
    typ = None
    try:
        familie = element.Parameter[DB.BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
        typ = element.Parameter[DB.BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
        coll = Element_TypFilter(familie, typ, True)
        idListe = list(coll.ToElementIds())
        if any(idListe):
            pickElements(idListe)
    except Exception as e:
        pass
    return idListe, familie, typ


def AlleElementInDocument():
    '''Wählt alle Elmente des angeklickten Typs im gesamten Projekt aus und gibt den Familien- und Typnamen zurück'''

    element = Pickelem()
    idListe = []
    familie =None
    typ = None
    try:
        familie = element.Parameter[DB.BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
        typ = element.Parameter[DB.BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
        coll = Element_TypFilter(familie, typ)
        idListe = list(coll.ToElementIds())
        if any(idListe):
            pickElements(idListe)
    except:
        pass
    return idListe,familie, typ



def Element_FamilyFilter(Fam_name = None, ActiveviewOn = None):

    param_equality=DB.FilterStringEquals()

    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)

    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)

    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,Fam_name,True)

    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    if ActiveviewOn:
        coll = DB.FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(revit.doc).WherePasses(Fam_name_filter).WhereElementIsNotElementType()

    return coll

def Element_TypFilter(Fam_name = None, Typname = None, ActiveviewOn = None):

    param_equality=DB.FilterStringEquals()

    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)

    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)

    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,Fam_name,True)

    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)


    Fam_typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)

    Fam_typ_prov = DB.ParameterValueProvider(Fam_typ_id)

    Fam_typ_value_rule = DB.FilterStringRule(Fam_typ_prov, param_equality, Typname, True)

    Fam_typ_filter = DB.ElementParameterFilter(Fam_typ_value_rule)

    if ActiveviewOn:
        coll = DB.FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType().WherePasses(
            Fam_typ_filter)
    else:

        coll = DB.FilteredElementCollector(revit.doc).WherePasses(Fam_name_filter).WhereElementIsNotElementType().WherePasses(Fam_typ_filter)

    return coll

def Element_FamilyContainsFilter(Fam_name = None, ActiveviewOn = None):

    param_equality=DB.FilterStringContains()

    Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)

    Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)

    Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,Fam_name,True)

    Fam_name_filter = DB.ElementParameterFilter(Fam_name_value_rule)

    if ActiveviewOn:

        coll = DB.FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id).WherePasses(Fam_name_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(revit.doc).WherePasses(
            Fam_name_filter).WhereElementIsNotElementType()

    return coll

def Muster_Pruefen(el):
    '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
    try:
        bb = el.LookupParameter('Bearbeitungsbereich').AsValueString()
        if bb == 'KG4xx_Musterbereich':return True
        else:return False
    except:return False