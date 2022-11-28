# coding: utf8
from pyrevit import revit, UI, DB,script
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,UIView,ExternalEvent

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'HK-Bauteile')
logger = script.get_logger()
DICT_MEP_NUMMER_NAME = {}
Liste_Params = []

def get_MEP_NUMMER_NAMR():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        DICT_MEP_NUMMER_NAME[el.Number] = [el.Number + ' - ' + el.LookupParameter('Name').AsString(),el.Id.ToString()]
    for elem in el.Parameters:
        if elem.StorageType.ToString() == 'Double':
            try:
                if elem.Definition.Name not in Liste_Params:
                    Liste_Params.append(elem.Definition.Name)
            except:pass
    
get_MEP_NUMMER_NAMR()

class GrundInfo(object):
    def __init__(self,name,index):
        self.name = name
        self.index = index

DICT_NEW = {}
FAM_IS = ObservableCollection[GrundInfo]()

def get_Heizkoeper_IS():
    Dict = {}
    HLSs = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
        .WhereElementIsNotElementType().ToElements()
    for el in HLSs:
        FamilyName = el.Symbol.FamilyName
        if FamilyName not in Dict.keys():
            Dict[FamilyName] = [[el.Id.ToString()],[]]
            params = el.Parameters
            for param in params:
                try:
                    if param.StorageType.ToString() == 'Double' and param.IsReadOnly == False:
                        Dict[FamilyName][1].append(param.Definition.Name)
                except:pass
        else:Dict[FamilyName][0].append(el.Id)
    
    for famindex,fam in enumerate(sorted(Dict.keys())):
        FAM_IS.Add(GrundInfo(fam,famindex))
        DICT_NEW[famindex] = [fam,ObservableCollection[GrundInfo](),Dict[fam][0],{},{}]
        for paramindex,para in enumerate(sorted(Dict[fam][1])):
            DICT_NEW[famindex][1].Add(GrundInfo(para,paramindex))
            DICT_NEW[famindex][3][paramindex] = para
            DICT_NEW[famindex][4][para] = paramindex
    
get_Heizkoeper_IS()

class HK_Familie(object):
    def __init__(self):
        self.Alle_Info = DICT_NEW
        self.Familien = FAM_IS
        self.familieindex = -1
        self.nennleistungindex = -1
        self.heizleistungindex = -1
        self.nennleistungName = ''
        self.heizleistungName = ''


    @property
    def Parameter_Dict(self):
        try:return self.Alle_Info[self.familieindex][3]
        except:return {}
    
    @property
    def Parameterindex_Dict(self):
        try:return self.Alle_Info[self.familieindex][4]
        except:return {}

    @property
    def elemids(self):
        try:return self.Alle_Info[self.familieindex][2]
        except:return []
    @property
    def Parameter_NL(self):
        try:return self.Alle_Info[self.familieindex][1]
        except:return ObservableCollection[GrundInfo]()
    @property
    def FamilienName(self):
        try:
            for x in self.Familien: 
                if x.index == self.familieindex:return x.name 
        except:return ''
    
    def get_nennleistungName(self):
        try:self.nennleistungName = self.Parameter_Dict[self.nennleistungindex]
        except:self.nennleistungName = ''

    def get_heizleistungName(self):
        try:self.heizleistungName = self.Parameter_Dict[self.heizleistungindex]
        except:self.heizleistungName = ''
    
    def get_nennleistungindex(self):
        try:self.nennleistungindex = self.Parameterindex_Dict[self.nennleistungName]
        except:self.nennleistungindex = -1
    
    def get_heizleistungindex(self):
        try:self.heizleistungindex = self.Parameterindex_Dict[self.heizleistungName]
        except:self.heizleistungindex = -1

AUSWAHL_HEIZKOERPER_IS = ObservableCollection[HK_Familie]()

def Get_HK_Einstellung():
    try:
        if len(config.HK_Familie):
            for el in config.HK_Familie:
                temp = HK_Familie()
                AUSWAHL_HEIZKOERPER_IS.Add(temp)
                for fam in temp.Familien:
                    if el[0]== fam.name:
                        temp.familieindex = fam.index
                        temp.nennleistungName = el[1]
                        temp.heizleistungName = el[2]
                        temp.get_heizleistungindex()
                        temp.get_nennleistungindex()
                        temp.get_heizleistungName()
                        temp.get_nennleistungName()
    except:pass

Get_HK_Einstellung()

class HeizkoerperInstance(object):
    def __init__(self,elemid,Nl_Param,HL_Param):
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.Param_NL = self.elem.LookupParameter(Nl_Param)
        self.Param_HL = self.elem.LookupParameter(HL_Param)
        self.phase = doc.GetElement(self.elem.CreatedPhaseId)
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.ismuster =  self.Muster_Pruefen(self.elem)
        self.raum = ''
        self.raumnummer = ''
        self.raumid = ''
        self.GetRaum()
        self.Nennleistung = round(self.get_value(self.Param_NL))
        self.Heizleistung = round(self.get_value(self.Param_HL))
    
    def get_value(self,param):
        """gibt den gesuchten Wert ohne Einheit zurück"""
        if not param:return
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

    def Muster_Pruefen(self,el):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = el.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False

    def GetRaum(self):
        try:
            self.raumid = self.elem.Space[self.phase].Id.ToString()
            self.raumnummer = self.elem.Space[self.phase].Number
            self.raum = self.elem.Space[self.phase].Number + ' - ' + self.elem.Space[self.phase].LookupParameter('Name').AsString()
        except:
            if not self.ismuster:
                try:
                    param_einbauort = self.get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
                    if param_einbauort not in DICT_MEP_NUMMER_NAME.keys():
                        logger.error('Einbauort konnte nicht ermittelt werden, FamilieName: {}, TypName: {}, ElementId: {}'.format(self.familyname,self.typname,self.elem.Id.ToString()))
                    else:
                        self.raumid = DICT_MEP_NUMMER_NAME[param_einbauort][1]
                        self.raumnummer = param_einbauort
                        self.raum = DICT_MEP_NUMMER_NAME[param_einbauort][0]
                except:pass
            return
    
    def Wert_schreiben(self):
        try:self.Param_HL.SetValueString(str(round(self.Heizleistung,2)))
        except:pass

class MEP_Raum:
    def __init__(self,elemid):
        self.elem = doc.GetElement(elemid)
        self.Param_HK = '' 
        self.Leistung = ''
        self.LeistungInDouble = 0
        self.HK_Bauteile = ObservableCollection[HeizkoerperInstance]()

    def get_HKL(self):
        try:
            unit = self.elem.LookupParameter(self.Param_HK).DisplayUnitType
            value = DB.UnitUtils.ConvertFromInternalUnits(self.elem.LookupParameter(self.Param_HK).AsDouble(),unit)
        except:
            try:
                # in Revit 2021/2022
                unit = self.elem.LookupParameter(self.Param_HK).GetUnitTypeId()
                value = DB.UnitUtils.ConvertFromInternalUnits(self.elem.LookupParameter(self.Param_HK).AsDouble(),unit)
            except:
                value = 0

        self.LeistungInDouble = round(value)
        self.Leistung = str(int(round(value))) + ' w'

class MEPRaumFilter(ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2003600':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class PickMEP(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None
            
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        re = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,MEPRaumFilter())
        elem = doc.GetElement(re)
        Nummer = elem.Number + ' - ' +elem.LookupParameter('Name').AsString()
        self.GUI_MEPRaum.Nummer.Text = Nummer
        self.GUI_MEPRaum.mepraum = self.GUI_MEPRaum.MEP_IS[Nummer]
        self.GUI_MEPRaum.set_up()
        
    def GetName(self):
        return "Show Element"

class ChangeParameterWert(IExternalEventHandler):
    def __init__(self):
        self.bauteile = None
        self.raum = ''

    def Execute(self,uiapp):
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,'HK-Leistung-Verteilung für Raum {}'.format(self.raum))
        t.Start()
        for bauteil in self.bauteile:
            bauteil.Wert_schreiben()
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "wert schreiben"

class ChangeParameterWertForAll(IExternalEventHandler):
    def __init__(self):
        self.GUI_MEPRaum = None

    def Execute(self,uiapp):
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,'HK-Leistung-Verteilung für alle Räume')
        t.Start()
        for nummer in sorted(self.GUI_MEPRaum.MEP_IS.keys()):
            el = self.GUI_MEPRaum.MEP_IS[nummer]
            el.Param_HK = self.GUI_MEPRaum.Param_HK
            el.get_HKL()
            if el.HK_Bauteile.Count == 0:continue
            summe = sum([x.Nennleistung for x in el.HK_Bauteile])
            for bauteil in el.HK_Bauteile:
                bauteil.Heizleistung = round(el.LeistungInDouble / summe * bauteil.Nennleistung,2)
                bauteil.Wert_schreiben()

        t.Commit()
        t.Dispose()
        try:
            self.GUI_MEPRaum.lv_HK.Items.Refresh()
        except:pass

    def GetName(self):
        return "HK-Leistung-Verteilung für alle Räume"