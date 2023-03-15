# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from pyrevit import script, forms
from rpw import *
import System
from Autodesk.Revit.DB import *
import xlsxwriter


__title__ = "0.82 Rohre"
__doc__ = """Brandschott"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)


revitLinks_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

revitLinks_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()
system_rohr = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()
system_rohr_dict = {}

def coll2dict(coll,dict):
    for el in coll:
        name = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
        type = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
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

walls = FilteredElementCollector(rvtdoc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
wallsName = []
for el in walls:
    if not el.Name in wallsName:
        wallsName.append(el.Name)
BrandWalls = forms.SelectFromList.show(wallsName,
multiselect=True, button_name='Select Walls')

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
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.TypName.upper().find(text_typ) != -1:
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
            sysname = elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
            systype = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            if not systype in SystemListe.Keys:
                SystemListe[systype] = [elem]
            else:
                SystemListe[systype].append(elem)
def getSolids(elem):
    lstSolid = []
    opt = Options()
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
        tempSolid = SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
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
Datenbank = {}

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

            n = 0
            elem = doc.GetElement(ElementId(int(elemid)))
            pipecurve = ProKurve(elem)
            elid = elem.Id.ToString()
            opt1 = SolidCurveIntersectionOptions()
            opt1.ResultType = SolidCurveIntersectionMode.CurveSegmentsOutside
            opt2 = SolidCurveIntersectionOptions()
            opt2.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
            
            for item in RvtLinkElemSolids:
                if elid in Datenbank.keys():
                    break
                elelink = item[1]
                for solid in elelink:
                    result1 = solid.IntersectWithCurve(pipecurve,opt1)
                    result2 = solid.IntersectWithCurve(pipecurve,opt2)
                    if result1.SegmentCount == 0 and result2.SegmentCount > 0:
                        Systemtyp = elem.MEPSystem.LookupParameter('Typ').AsValueString()
                        DN = elem.LookupParameter('Größe').AsString()
                        Length = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                        Datenbank[elid] = [Systemtyp,Length,DN]
                        break



if any(Datenbank.Keys):
    Liste_Export = [['ElementId','System','Länge','Größe']]
    for el in Datenbank.Keys:
        Liste_Export.append([el,Datenbank[el][0],Datenbank[el][1],Datenbank[el][2]])
        
    path = r'C:\Users\Zhang\Desktop\test_neu.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for col in range(len(Liste_Export[0])):
        for row in range(len(Liste_Export)):
            try:
                float(Liste_Export[row][col])
                worksheet.write_number(row, col, float(Liste_Export[row][col]))
            except:
                worksheet.write(row, col, Liste_Export[row][col])
    worksheet.autofilter(0, 0, int(len(Liste_Export))-1, int(len(Liste_Export[0])-1))
    workbook.close()

