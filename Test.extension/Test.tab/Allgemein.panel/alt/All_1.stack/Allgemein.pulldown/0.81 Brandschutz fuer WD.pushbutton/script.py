# coding: utf8
import clr
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from pyrevit import script, forms,revit
from Autodesk.Revit.DB import *
from System import Guid

__title__ = "0.81 ergänzt Brandschutzklasse für WD (MFC)"
__doc__ = """Brandschutz für WD (MFC)
Parameter: IGF_HLS_Brandschutz (276f654e-b5ea-4ead-ac28-fdc326b12e12)
Kategorie: HLS-Bauteile"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config = script.get_config()

uidoc = revit.uidoc
doc = revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)

# Wände
revitLinks_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

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

# Wanddurchbrüche

HLS_Bauteile = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
HLS_Ids = HLS_Bauteile.ToElementIds()
HLS_Bauteile.Dispose()

HLS_Dict = {}
HLS_Para_Dict = {}
for el in HLS_Ids:
    elem = doc.GetElement(el)
    Family = elem.Symbol.FamilyName
    if not Family in HLS_Dict.keys():
        HLS_Dict[Family] = [elem]
        Paraliste = []
        for para in elem.Parameters:
            if not para.Definition.Name in Paraliste:
                Paraliste.append(para.Definition.Name)
        HLS_Para_Dict[Family]=Paraliste
    else:
        HLS_Dict[Family].append(elem)


class HLS(object):
    def __init__(self):
        self.checked = False
        self.Name = ''
        self.SelectedOption = 0
        self.Paras = []

    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self, value):
        self._Name = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def SelectedOption(self):
        return self._SelectedOption
    @SelectedOption.setter
    def SelectedOption(self, value):
        self._SelectedOption = value
    @property
    def Paras(self):
        return self._Paras
    @Paras.setter
    def Paras(self, value):
        self._Paras = value

class select(object):
    def __init__(self):
        self.selectId = 0
        self.ParaName = []

    @property
    def ParaName(self):
        return self._ParaName
    @ParaName.setter
    def ParaName(self, value):
        self._ParaName = value
    @property
    def selectId(self):
        return self._selectId
    @selectId.setter
    def selectId(self, value):
        self._selectId = value

Liste_WD = ObservableCollection[HLS]()

hlsname = HLS_Para_Dict.keys()[:]
hlsname.sort()
for elem in hlsname:
    temp = HLS()
    temp.Name = elem
    params = HLS_Para_Dict[elem][:]
    params.sort()
    para_Liste = []
    for n,para in enumerate(params):
        temp_para = select()
        temp_para.selectId = n
        temp_para.ParaName = para
        para_Liste.append(temp_para)
    temp.Paras = para_Liste
    Liste_WD.Add(temp)

# GUI Systemauswahl
class Wanddurchbrueche(WPFWindow):
    def __init__(self, xaml_file_name,Liste_WandD):
        self.Liste_WandD = Liste_WandD
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[HLS]()
        self.altdatagrid = Liste_WandD
        self.dataGrid.ItemsSource = Liste_WandD
        self.read_config()
        
        self.suche.TextChanged += self.search_txt_changed
    
    def read_config(self):
        try:
            wddict = config.Wd_dict
            for item in self.Liste_WandD:
                if item.Name in wddict.keys():
                    item.checked = wddict[item.Name][0]
                    item.SelectedOption = wddict[item.Name][1]
            self.dataGrid.Items.Refresh()
        except:
            pass
        try:
            if config.Hide:
                tempcoll = ObservableCollection[HLS]()
                for item in self.dataGrid.Items:
                    if item.checked:
                        tempcoll.Add(item)
                self.dataGrid.ItemsSource = tempcoll
                self.dataGrid.Items.Refresh()
                   
        except Exception as e:
            pass
        
        
    def write_comfig(self):
        dict_ = {}
        for item in self.Liste_WandD:
            dict_[item.Name] = [item.checked,item.SelectedOption]

        config.Wd_dict = dict_
        script.save_config()

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
                if item.Name.upper().find(text_typ) != -1:
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

    def select(self,sender,args):
        self.write_comfig()
        self.Close()

    def hide(self,sender,args):
        tempcoll = ObservableCollection[HLS]()
        for item in self.dataGrid.Items:
            if item.checked:
                tempcoll.Add(item)
        self.dataGrid.ItemsSource = tempcoll
        config.Hide = True
        script.save_config()

    def show(self,sender,args):
        self.dataGrid.ItemsSource = self.altdatagrid
        config.Hide = False
        script.save_config()
    
Systemwindows = Wanddurchbrueche("window.xaml",Liste_WD)
Systemwindows.ShowDialog()

FamiliemitTiefe = {}
for ele in Liste_WD:
    if ele.checked:
        paras = ele.Paras
        selected = ele.SelectedOption
        for el in paras:
            if el.selectId == selected:
                FamiliemitTiefe[ele.Name] = el.ParaName


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

def HLSkurve(elem,Tiefe):
    hlscurve = None
    BB = elem.get_BoundingBox(None)
    list = []
    Len_X = int((BB.Max.X - BB.Min.X) * 304.8)
    Len_Y = int((BB.Max.Y - BB.Min.Y) * 304.8)
    Len_Z = int((BB.Max.Z - BB.Min.Z) * 304.8)
    Cen_X = (BB.Max.X + BB.Min.X) / 2
    Cen_Y = (BB.Max.Y + BB.Min.Y) / 2
    Cen_Z = (BB.Max.Z + BB.Min.Z) / 2
    if abs(int(Tiefe) - Len_X) < 30:
        list = [XYZ(BB.Max.X + 0.2,Cen_Y,Cen_Z),XYZ(BB.Min.X - 0.2,Cen_Y,Cen_Z)]
    elif abs(int(Tiefe) - Len_Y) < 30:
        list = [XYZ(Cen_X,BB.Max.Y + 0.2,Cen_Z),XYZ(Cen_X,BB.Min.Y - 0.2,Cen_Z)]
    elif abs(int(Tiefe) - Len_Z) < 30:
        list = [XYZ(Cen_X,Cen_Y,BB.Max.Z + 0.2),XYZ(Cen_X,Cen_Y,BB.Min.Y - 0.2)]

    if any(list):
        hlscurve = Line.CreateBound(list[0], list[1])
        return hlscurve
    else:
        return None

def EbenenUmbenennen(ebene):
    out = ''
    if not ebene:
        out = ''
    if ebene.find('EG') != -1:
        out = 'EG'
    elif ebene.find('SAN') != -1:
        out = 'EG'
    elif ebene.find('1.OG') != -1:
        out = '1.OG'
    elif ebene.find('2.OG') != -1:
        out = '2.OG'
    elif ebene.find('3.OG') != -1:
        out = '3.OG'
    elif ebene.find('1.UG') != -1:
        out = '1.UG'
    elif ebene.find('2.UG') != -1:
        out = '2.UG'
    elif ebene.find('3.UG') != -1:
        out = '3.UG'
    else:
        out = '4.OG'
    return out

def NameUmbenennen(Name):
    out = ''
    if Name.find('F30') != -1:
        out = 'F30'
    elif Name.find('F90') != -1:
        out = 'F90'
    elif Name.find('F120') != -1:
        out = 'F120'
    else:
        out = 'F0'
    return out

# RvtLinkElem
RvtLinkElemSolids = {}
# ProElemCurve = []
step = int(len(BrandWallEles)/200)
with forms.ProgressBar(title='{value}/{max_value} Wände in RVT-Link Model',cancellable=True, step=step) as pb:
    for n_1, ele in enumerate(BrandWallEles):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n_1, len(BrandWallEles))
        models = TransformSolid(ele)
        ebenename = rvtdoc.GetElement(ele.LevelId).Name
        WallName = ele.Name
        Klasse = NameUmbenennen(WallName)
        Ebenen = EbenenUmbenennen(ebenename)
        if not Ebenen in RvtLinkElemSolids.keys():
            RvtLinkElemSolids[Ebenen] = {}
        if not Klasse in RvtLinkElemSolids[Ebenen].keys():
            RvtLinkElemSolids[Ebenen][Klasse] = []
        if not models in RvtLinkElemSolids[Ebenen][Klasse]:
            RvtLinkElemSolids[Ebenen][Klasse].append(models)

rvtdoc.Dispose()
Datenbank = {}
# print(RvtLinkElemSolids.keys())
# for key in RvtLinkElemSolids.keys():
#     print(RvtLinkElemSolids[key].keys())

for key in FamiliemitTiefe.keys():
    Elements = HLS_Dict[key]
    title = '{value}/{max_value} Bauteile in Familie (' + key+')'
    with forms.ProgressBar(title=title,cancellable=True, step=10) as pb2:
        for n_1,elem in enumerate(Elements):
            if pb2.cancelled:
                script.exit()
            pb2.update_progress(n_1+1, len(Elements))
            Klass = ''
            ebene = doc.GetElement(elem.LevelId).Name
            neu_ebene = EbenenUmbenennen(ebene)
            para = FamiliemitTiefe[key]
            tiefe = elem.GetParameters(para)[0].AsValueString()
            curve = HLSkurve(elem,tiefe)
            if not curve:
                logger.error('Linie kann nicht erstellt werden. ID: ' + elem.Id.ToString())
                continue
            if neu_ebene in RvtLinkElemSolids.keys():
                models_dict = RvtLinkElemSolids[neu_ebene]
                for klasse in ['F120','F90','F30','F0']:
                    if not Klass:
                        if klasse in models_dict.keys():
                            models = models_dict[klasse]
                            opt = SolidCurveIntersectionOptions()
                            for item in models:
                                if not Klass:
                                    for solid in item:
                                        result = solid.IntersectWithCurve(curve,opt)
                                        if result.SegmentCount > 0:
                                            Klass = klasse
                                            Datenbank[elem.Id.ToString()] = Klass
            if not elem.Id.ToString() in Datenbank.keys():
                Datenbank[elem.Id.ToString()] = ''
# Daten schreiben
if forms.alert("Daten schreiben?", ok=False, yes=True, no=True):
    t = Transaction(doc,'Brandschutz für WD')
    t.Start()
    eles = Datenbank.keys()[:]
    step3 = int(len(eles)/1000)+10
    with forms.ProgressBar(title='{value}/{max_value} Wanddurchbrüche',cancellable=True, step=step3) as pb3:
        for n,el in enumerate(eles):
            if pb3.cancelled:
                t.RollBack()
                script.exit()
            pb3.update_progress(n, len(eles))
            doc.GetElement(ElementId(int(el))).get_Parameter(Guid('276f654e-b5ea-4ead-ac28-fdc326b12e12')).Set(Datenbank[el])
    t.Commit()