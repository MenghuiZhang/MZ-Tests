# coding: utf8
from pyrevit import script, forms
import rpw
import time
import Autodesk.Revit.DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import System
from System.IO import Path, File
from pyrevit import revit


"""
Funktion:
get_value(param);
ALLParameterTyp();
AllKategorien();
ALLParameterGroup();
RowAddSharedParam(Gruppe = 'neue Gruppe', Name = None, Disziplin = 'Allgemein',
 Typ = 'Text', Info = None, GUID = None);
RowAddProjParam(Name = None, Disziplin = 'Allgemein', Parametertyp = 'Text',
Info = None, Gruppe = 'Sonstige', Kategorien = 'Allgemeines Modell',
Typ_Exemplar = 'Typ');
AddProjParamfromShared(SharedGruppe = 'neue Gruppe', GUID = None, Name = None,
Disziplin = 'Allgemein', Parametertyp = 'Text', Info = None,
Gruppe = 'Sonstige', Kategorien = 'Allgemeines Modell', Typ_Exemplar = 'Typ');
getSolids(elem): get Solid of element;
TransformSolid(elem,revitlink): get Solid of element in Linkmodell;
ProKurve(elem): get pipe/ductcurve in modell;
set_value(elem,paramName,value): write value in parameter;
pickElements(elementids):pick elements, Input: Elementid_list
AlleElementInAnsicht(): Element mit gleiche Familie und Typ in ein Ansicht auswählen
AlleElementInDocument(): Element mit gleiche Familie und Typ in Projekt auswählen
Pickelem(): pick element während des Durchlaufes der Skript
"""

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp
activeview = uidoc.ActiveView

filename = app.SharedParametersFilename
file = app.OpenSharedParameterFile()

def get_value(param):
    """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""
    value = revit.query.get_param_value(param)
    try:
        unit = param.DisplayUnitType
        value = UnitUtils.ConvertFromInternalUnits(value,unit)
    except Exception as e:
        pass
    return value

def ALLParameterTyp():
    """ get all Parametertype.
    Bsp.  _ParameterType['Allgemein']['Text'] = DB.ParameterType.Text"""
    _ParameterType = {}
    _ParameterType['Allgemein'] = {}
    _ParameterType['Energie'] = {}
    _ParameterType['Tragwerk'] = {}
    _ParameterType['Lüftung'] = {}
    _ParameterType['Rohre'] = {}
    _ParameterType['Elektro'] = {}
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

    alltype = System.Enum.GetValues(ParameterType.Invalid.GetType())
    for i in alltype:
        Type = i.ToString()
        if not Type in ['Invalid','FamilyType']:
            type = LabelUtils.GetLabelFor(i)
            dis = DeiziplinErmitteln(type,Type)
            _ParameterType[dis][type] = i

    return _ParameterType

def AllKategorien():
    """ get all categories.
    Bsp.  _Categories['Leuchten'] = DB.Category"""
    _Categories = {}
    for cat in doc.Settings.Categories:
        name = cat.Name
        _Categories[name] = cat
    return _Categories

def ALLParameterGroup():
    """ get all categories.
    Bsp.  _Categories['HLS'] = DB.BuiltInParameterGroup.PG_MECHANICAL"""
    _parametergroup = {}
    allGro = System.Enum.GetValues(BuiltInParameterGroup.PG_ROUTE_ANALYSIS.GetType())
    for i in allGro:
        name = LabelUtils.GetLabelFor(i)
        _parametergroup[name] = i
    return _parametergroup

def ALLBuiltinCategory():
    """ get all builtincategory.
    Bsp.  _Categories['MEP-Räume'] = DB.BuiltInCategory.OST_MEPSpaces"""
    _builtInCategory = {}
    allcate = System.Enum.GetValues(BuiltInCategory.OST_PointClouds.GetType())
    for cate in allcate:
        try:
            name = doc.Settings.Categories.get_Item(cate).Name
            if name:
                _builtInCategory[name] = cate
        except:
            pass
    return _builtInCategory

def RowAddSharedParam(Gruppe = 'neue Gruppe', Name = None, Disziplin = 'Allgemein', Typ = 'Text', Info = None, GUID = None):
    """ create sharedparameter.
    for example: RowAddSharedParam(Gruppe = 'Testgruppe', Name = 'TestName',
    Disziplin = 'Allgemein', Typ = 'Text')
    """
    def AddGroup(inFile,InGroup):
        dgs = inFile.Groups
        Groups = [i.Name for i in dgs]
        if not InGroup in Groups:
            file.Groups.Create(InGroup)

    AddGroup(file,Gruppe)
    _AllDifis = []

    try:
        _AllDifis = [i.Name for i in file.Groups[Gruppe].Definitions]
    except Exception as e:
        pass

    if not Name in _AllDifis:
        parameterType = ALLParameterTyp()[Disziplin][Typ]
        DefiCrea = ExternalDefinitionCreationOptions(Name, parameterType)
        if Info:
            DefiCrea.Description = Info
        if GUID:
            DefiCrea.GUID = System.Guid(GUID)
        sharedPara = file.Groups[InGruppe].Definitions.Create(DefiCrea)

def RowAddProjParam(Name = None, Disziplin = 'Allgemein', Parametertyp = 'Text', Info = None, Gruppe = 'Sonstige', Kategorien = 'Allgemeines Modell', Typ_Exemplar = 'Typ'):
    """ create projectparameter.
    for example: RowAddProjParam(Name = 'TestName',  Disziplin = 'Allgemein',
    Typ = 'Text', Gruppe = 'Sonstige', Kategorien = 'Allgemeines Modell',
    Typ_Exemplar = 'Typ')
    """
    tempFile = Path.GetTempFileName() + ".txt"
    File.Create(tempFile)
    app.SharedParametersFilename = tempFile
    parameterType = ALLParameterTyp()[Disziplin][Parametertyp]
    DefiCrea = ExternalDefinitionCreationOptions(Name, parameterType)
    if Info:
        DefiCrea.Description = Info
    ExternalDefinition = app.OpenSharedParameterFile().Groups.Create("TemporaryDefintionGroup").Definitions.Create(DefiCrea)

    app.SharedParametersFilename = filename
    File.Delete(tempFile)

    ParaCatSet = app.Create.NewCategorySet()
    ParCatList = Kategorien.Split(',')

    for i in ParCatList:
        if i in AllKategorien().keys():
            ParaCatSet.Insert(AllKategorien()[i])

    binding = None
    if Typ_Exemplar == 'Exemplar':
        binding = app.Create.NewInstanceBinding(ParaCatSet)
    else:
        binding = app.Create.NewTypeBinding(ParaCatSet)

    ParaGroup = ALLParameterGroup()['Sonstige']
    if Gruppe in ALLParameterGroup().keys():
        ParaGroup = ALLParameterGroup()[Gruppe]

    map = uiapp.ActiveUIDocument.Document.ParameterBindings
    map.Insert(ExternalDefinition, binding, ParaGroup)

def AddProjParamfromShared(SharedGruppe = 'neue Gruppe', GUID = None, Name = None, Disziplin = 'Allgemein', Parametertyp = 'Text', Info = None, Gruppe = 'Sonstige', Kategorien = 'Allgemeines Modell', Typ_Exemplar = 'Typ'):
    """ create sharedparameter.
    for example: AddProjParamfromShared(SharedGruppe = 'neue Gruppe',
    Name = 'TestName', Disziplin = 'Allgemein', Typ = 'Text',
    Gruppe = 'Sonstige', Kategorien = 'Allgemeines Modell', Typ_Exemplar = 'Typ')
    """

    parameterType = ALLParameterTyp()[Disziplin][Parametertyp]
    DefiCrea = ExternalDefinitionCreationOptions(Name, parameterType)
    if Info:
        DefiCrea.Description = Info
    if GUID:
        DefiCrea.GUID = System.Guid(GUID)

    group = None
    difinition = None
    try:
        group = app.OpenSharedParameterFile().Groups.Create(SharedGruppe)
    except:
        group = app.OpenSharedParameterFile().Groups[SharedGruppe]
    try:
        difinition = group.Definitions.Create(DefiCrea)
    except:
        difinitions = {i.Name: i for i in group.Definitions}
        difinition = difinitions[Name]

    ParaCatSet = app.Create.NewCategorySet()
    ParCatList = Kategorien.Split(',')

    for i in ParCatList:
        if i in AllKategorien().keys():
            ParaCatSet.Insert(AllKategorien()[i])

    binding = None
    if Typ_Exemplar == 'Exemplar':
        binding = app.Create.NewInstanceBinding(ParaCatSet)
    else:
        binding = app.Create.NewTypeBinding(ParaCatSet)

    ParaGroup = ALLParameterGroup()['Sonstige']
    if Gruppe in ALLParameterGroup().keys():
        ParaGroup = ALLParameterGroup()[Gruppe]

    map = uiapp.ActiveUIDocument.Document.ParameterBindings
    map.Insert(difinition, binding, ParaGroup)

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
    opt = Options()
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
        tempSolid = SolidUtils.CreateTransformed(solid,revitlink.GetTransform())
        m_lstModels.append(tempSolid)
    return m_lstModels

def ProKurve(elem):
    pipecurve = None
    csi = elem.ConnectorManager.Connectors.ForwardIterator()
    list = []
    while csi.MoveNext():
        conn = csi.Current
        list.append(conn.Origin)
    pipecurve = Line.CreateBound(list[0], list[1])
    return pipecurve

def set_value(elem,paramName,value):
    param = elem.LookupParameter(paramName)
    StorageType = param.StorageType.ToString()
    if StorageType == 'Double':
        try:
            param.SetValueString(str(value))
        except:
            pass
    elif StorageType == 'String':
        try:
            param.Set(str(value))
        except:
            pass
    else:
        try:
            param.Set(str(value))
        except:
            pass

def pickElements(elementids):
    sel = uidoc.Selection.GetElementIds()
    for el in elementids:
        sel.Add(el)
    uidoc.Selection.SetElementIds(sel)

def Pickelem():
    try:
        re = uidoc.Selection.PickObject(Selection.ObjectType.Element)
        elem = doc.GetElement(re)
        return elem
    except:
        pass

def AlleElementInAnsicht():
    element = Pickelem()
    idListe = []
    familie = None
    try:
        cateName = element.Category.Name
        Category = ALLBuiltinCategory()[cateName]
        familie = element.Parameter[BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
        typ = element.Parameter[BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
        coll = FilteredElementCollector(doc,activeview.Id).OfCategory(Category).WhereElementIsNotElementType()

        for elem in coll:
            fa_name = elem.Parameter[BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
            ty_name = elem.Parameter[BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
            if typ == ty_name and fa_name == familie:
                idListe.append(elem.Id)
        if any(idListe):
            pickElements(idListe)
    except Exception as e:
        pass

    return idListe,familie
def AlleElementInDocument():
    element = Pickelem()
    idListe = []
    familie =None
    try:
        cateName = element.Category.Name
        Category = ALLBuiltinCategory()[cateName]
        familie = element.Parameter[BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
        typ = element.Parameter[BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
        coll = FilteredElementCollector(doc).OfCategory(Category).WhereElementIsNotElementType()

        for elem in coll:
            fa_name = elem.Parameter[BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
            ty_name = elem.Parameter[BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
            if typ == ty_name and fa_name == familie:
                idListe.append(elem.Id)

        if any(idListe):
            pickElements(idListe)
    except Exception as e:
        pass
    return idListe,familie
