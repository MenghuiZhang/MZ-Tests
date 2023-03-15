# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, UI, DB
from pyrevit import script, forms
from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection

__title__ = "zählt nötige Brandschotts (für Rohre)"
__doc__ = """zählt nötige Brandschotts (für Rohre)
Paremeter: IGF_HLS_Brandschott

[2021.11.15]
Version: 1.1
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass


revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

system_rohr = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()
system_rohr_dict = {}

rohesys = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()
rohr = rohesys.ToElementIds()
rohesys.Dispose()

def coll2dict(coll,dict):
    for el in coll:
        name = el.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
        type = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        if type in dict.Keys:
            dict[type].append(el.Id)
        else:
            dict[type] = [el.Id]

coll2dict(system_rohr,system_rohr_dict)
system_rohr.Dispose()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el

rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()
rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()

walls = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
wallsName = []
for el in walls:
    if not el.Name in wallsName:
        wallsName.append(el.Name)
BrandWalls = forms.SelectFromList.show(wallsName,multiselect=True, button_name='Select Walls')

BrandWallEles = []
for el in walls:
    if el.Name in BrandWalls:
        BrandWallEles.append(el)

class System(object):
    def __init__(self):
        self.checked = False
        self.SystemName = ''
        self.TypName = ''

    @property
    def TypName(self):
        return self._TypName
    @TypName.setter
    def TypName(self, value):
        self._TypName = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def ElementId(self):
        return self._ElementId
    @ElementId.setter
    def ElementId(self, value):
        self._ElementId = value

Liste_Rohr = ObservableCollection[System]()

for key in system_rohr_dict.Keys:
    temp_system = System()
    temp_system.TypName = key
    temp_system.ElementId = system_rohr_dict[key]
    Liste_Rohr.Add(temp_system)

# GUI Systemauswahl
class Systemauswahl(WPFWindow):
    def __init__(self, xaml_file_name,liste_Rohr):
        self.liste_Rohr = liste_Rohr
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[System]()
        self.altdatagrid = None

        try:
            self.dataGrid.ItemsSource = liste_Rohr
            self.altdatagrid = liste_Rohr
        except Exception as e:
            logger.error(e)

        self.suche.TextChanged += self.search_txt_changed

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.TypName.find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = True
        self.dataGrid.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = False
        self.dataGrid.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.checked
            item.checked = not value
        self.dataGrid.Items.Refresh()

    def auswahl(self,sender,args):
        self.Close()
    
Systemwindows = Systemauswahl("System.xaml",Liste_Rohr)
Systemwindows.ShowDialog()

SystemListe = {}
for el in Liste_Rohr:
    if el.checked == True:
        for it in el.ElementId:
            elem = doc.GetElement(it)
            sysname = elem.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
            systype = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            if not systype in SystemListe.Keys:
                SystemListe[systype] = [elem]
            else:
                SystemListe[systype].append(elem)

def getSolids(elem):
    lstSolid = []
    opt = DB.Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    ge = elem.get_Geometry(opt)
    if ge != None:
        lstSolid.extend(GetSolid(ge))
    return lstSolid
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

def TransformSolid(elem):
    m_lstModels = []
    listSolids = getSolids(elem)
    for solid in listSolids:
        tempSolid = solid
        tempSolid = DB.SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
        m_lstModels.append(tempSolid)
    return m_lstModels

def ProKurve(elem):
    pipecurve = None
    csi = elem.ConnectorManager.Connectors.ForwardIterator()
    list = []
    while csi.MoveNext():
        conn = csi.Current
        list.append(conn.Origin)
    pipecurve = DB.Line.CreateBound(list[0], list[1])
    return pipecurve

# RvtLinkElem
RvtLinkElemSolids = []
# ProElemCurve = []
step = int(len(BrandWallEles)/200)
with forms.ProgressBar(title='{value}/{max_value} Wände in Revitverknüpfung',cancellable=True, step=step) as pb:
    n_1 = 0
    for ele in BrandWallEles:
        if pb.cancelled:
            script.exit()
        n_1 += 1
        pb.update_progress(n_1, len(BrandWallEles))
        models = TransformSolid(ele)
        RvtLinkElemSolids.append([ele,models])

rvtdoc.Dispose()
Datenbank = []

for systyp in SystemListe.Keys:
    sysliste = SystemListe[systyp]
    elements_liste = []
    for sys_ele in sysliste:
        elements = sys_ele.PipingNetwork
        for ele in elements:
            if ele.Category.Name == 'Rohre':
                if not ele.Id.ToString() in elements_liste:
                    elements_liste.append(ele.Id.ToString())
 
    title = '{value}/{max_value} Rohre in System ' + systyp
    with forms.ProgressBar(title=title,cancellable=True, step=5) as pb2: 
        for n_1,elemid in enumerate(elements_liste):
            if pb2.cancelled:
                script.exit()
            pb2.update_progress(n_1+1, len(elements_liste))
            elem = doc.GetElement(DB.ElementId(int(elemid)))

            n = 0
            pipecurve = ProKurve(elem)
            opt1 = DB.SolidCurveIntersectionOptions()
            opt1.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsOutside
            opt2 = DB.SolidCurveIntersectionOptions()
            opt2.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside
            for item in RvtLinkElemSolids:
                n1 = 0
                elelink = item[1]
                for solid in elelink:
                    result1 = solid.IntersectWithCurve(pipecurve,opt1)
                    result2 = solid.IntersectWithCurve(pipecurve,opt2)
                    if result1.SegmentCount > 0 and result2.SegmentCount > 0:
                        n1 = n1 + 1
                        break
                n = n + n1
            if n> 0:
                Datenbank.append([elem,n])
            else:
                Datenbank.append([elem,0])

# Daten schreiben
t = DB.Transaction(doc,'Brandschott zählen')
t.Start()
step3 = int(len(Datenbank)/200)
with forms.ProgressBar(title='{value}/{max_value} Rohre',cancellable=True, step=step3) as pb3:
    n_1 = 0
    for el in Datenbank:
        if pb3.cancelled:
            script.exit()
        n_1 += 1
        pb3.update_progress(n_1, len(Datenbank))
        el[0].LookupParameter('IGF_HLS_Brandschott').Set(int(el[1]))
t.Commit()