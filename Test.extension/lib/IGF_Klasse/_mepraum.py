# coding: utf8
from IGF_Klasse import DB,ItemTemplate,TemplateItemBase
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from IGF_Funktionen._Parameter import get_value
from System import Guid
from IGF_Allgemein import Parameter_Dict,Luftberechnung_NachName,Luftberechnung_NachNummer,Luftberechnung_Einheit
from rpw import db
from IGF_Funktionen._Parameter import get_value,get_Parameter,wert_schreibenbase
doc = __revit__.ActiveUIDocument.Document

class MEPRaumBase(TemplateItemBase):
    def __init__(self,elem):
        TemplateItemBase.__init__(self)
        self.elem = elem
        self.doc = self.elem.Document
        self.Nummer = self.elem.Number
        self.Name = get_value(self.elem.LookupParameter('Name'))

    @property
    def NummerName(self):
        return self.Nummer + ' - ' + self.Name
    
    def wert_schreiben(self,paramname,wert):
        wert_schreibenbase(self.elem.get_Parameter(Parameter_Dict[paramname]), wert)

class SchachtBase(MEPRaumBase):
    def __init__(self,elem):
        MEPRaumBase.__init__(self, elem)
        self.Param_IsSchacht = self.elem.get_Parameter(Parameter_Dict['IGF_RLT_InstallationsSchacht'])
        self.Param_SchachtName = self.elem.get_Parameter(Parameter_Dict['IGF_RLT_InstallationsSchachtName'])
        self.IsSchacht = get_value(self.Param_IsSchacht) 
        if self.IsSchacht == 1:
            self._schacht = True
        else:
            self._schacht = False
        
        self._Schachtname = get_value(self.Param_SchachtName) 
    
    @property
    def schacht(self):
        return self._schacht
    @schacht.setter
    def schacht(self,value):
        if self._schacht != value:
            self._schacht = value
            self.RaisePropertyChanged('schacht')

    @property
    def Schachtname(self):
        return self._Schachtname
    @Schachtname.setter
    def Schachtname(self,value):
        if self._Schachtname != value:
            self._Schachtname = value
            self.RaisePropertyChanged('Schachtname')

    def werte_schreiben(self):
        if self.schacht:wert_schreibenbase(self.Param_IsSchacht, 1)
        else:wert_schreibenbase(self.Param_IsSchacht, 0)
        wert_schreibenbase(self.Param_SchachtName, self.Schachtname)

class ReduzierFaktor(TemplateItemBase):
    def __init__(self,nummer,faktor = 1):
        TemplateItemBase.__init__(self)
        self._nummer = nummer
        self._faktor = faktor
    @property
    def nummer(self):
        return self._nummer
    @nummer.setter
    def nummer(self,value):
        if value != self._nummer:
            self._nummer = value
            self.RaisePropertyChanged('nummer')
    @property
    def faktor(self):
        return self._faktor
    @faktor.setter
    def faktor(self,value):
        if value != self._faktor:
            self._faktor = value
            self.RaisePropertyChanged('faktor')

class Schacht(SchachtBase):
    """Labaranschluss Parameter"""
    def __init__(self,elem):
        SchachtBase.__init__(self,elem)
        
        self._checked = False
        self.reduzierfaktor = ObservableCollection[ReduzierFaktor]()
        self.Param_Reduzierfaktor = self.elem.get_Parameter(Parameter_Dict['IGF_RLT_Schacht-LuftmengenReduzierung'])
        self.value = get_value(self.Param_Reduzierfaktor)
        self.daten = {}
        self.get_detail_daten()
        if not (self.elem.Area or self.elem.Volume):
            return
        try:self.get_AnlagenInSchacht()
        except:pass
    
    def get_wert(self):
        wert = ''
        for el in self.reduzierfaktor:
            wert += 'Anl_' + str(el.nummer) + '=' + str(el.faktor) + ', '
        if wert:
            wert = wert[:-2]
        self.value = wert
    
    def get_detail_daten(self):
        self.daten = {}
        if not self.value:
            return
        Liste = self.value.split(', ')
        for el in Liste:
            if el:
                if el.find('Anl_') != -1 and el.find('=') != -1:
                    try:
                        anlnummer = el[4:el.find('=')]
                        faktor = el[el.find('=')+1:]
                        self.daten[anlnummer] = faktor
                    except:
                        pass
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    
    def get_items(self):
        for el in self.reduzierfaktor:
            if el.nummer in self.daten.keys():
                el.faktor = self.daten[el.nummer]
    
    def werte_schreiben1(self):
        self.elem.get_Parameter(self.Param_Reduzierfaktor).Set(self.value)

    def get_AnlagenInSchacht(self):
        outline = DB.Outline(self.elem.get_BoundingBox(None).Min,self.elem.get_BoundingBox(None).Max)
        filter1 = DB.BoundingBoxIntersectsFilter(outline)
        filter2 = DB.BoundingBoxIsInsideFilter(outline)
        filter3 = DB.LogicalOrFilter(List[DB.ElementFilter]([filter1,filter2]))
        coll = DB.FilteredElementCollector(self.doc).OfCategory(DB.BuiltInCategory.OST_DuctCurves).WherePasses(filter3).WhereElementIsNotElementType().ContainedInDesignOption(DB.ElementId(-1)).ToElements()

        filter1.Dispose()
        filter2.Dispose()
        filter3.Dispose()
        outline.Dispose()

        Liste = []
        for elem in coll:
            system = elem.LookupParameter('Systemtyp').AsElementId()
            if system.IntegerValue != -1:
                param = self.doc.GetElement(system).get_Parameter(Parameter_Dict['IGF_X_AnlagenNr'])
                if param:
                    anlnr = param.AsValueString()
                    if anlnr not in Liste:Liste.append(anlnr)
        Liste.sort()
        Liste1 = []
        for anlagenr in Liste:
            if anlagenr in self.daten.keys():
                Liste1.append(ReduzierFaktor(anlagenr,self.daten[anlagenr]))
            else:
                Liste1.append(ReduzierFaktor(anlagenr))
        Items = ObservableCollection[ReduzierFaktor](Liste1)
        self.reduzierfaktor = Items

class MEPRaumFaytory:
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    colls = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    
    @staticmethod
    def GetAllSchacht():
        liste = []
        for el in MEPRaumFaytory.colls:
            tmp = Schacht(el)
            if tmp.schacht:
                liste.append(tmp)
        liste.sort(key=lambda x:x.Schachtname)
        return ObservableCollection[Schacht](liste)
    
    @staticmethod
    def GetAllMEPRaum():
        liste = [MEPRaumBase(elem) for elem in MEPRaumFaytory.colls]
        liste.sort(key=lambda x:x.NummerName)
        return ObservableCollection[MEPRaumBase](liste)


