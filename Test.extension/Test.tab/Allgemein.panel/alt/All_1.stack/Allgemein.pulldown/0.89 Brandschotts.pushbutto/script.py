# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, UI, DB
from pyrevit import script, forms


import clr
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
import time
from pyrevit import script, forms
from rpw import *
import System
from Autodesk.Revit.DB import *

start = time.time()

__title__ = "0.89 Brandschott"
__doc__ = """Brandschott"""
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

rohesys = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()
rohr = rohesys.ToElementIds()
rohesys.Dispose()

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
Datenbank = []

 
title = '{value}/{max_value} Rohre in Projekt '
with forms.ProgressBar(title=title,cancellable=True, step=5) as pb2: 
    for n_1,elemid in enumerate(rohr):
        if pb2.cancelled:
            script.exit()
        pb2.update_progress(n_1+1, len(elements_liste))
        elem = doc.GetElement(elemid)

        n = 0
        pipecurve = ProKurve(elem)
        opt1 = SolidCurveIntersectionOptions()
        opt1.ResultType = SolidCurveIntersectionMode.CurveSegmentsOutside
        opt2 = SolidCurveIntersectionOptions()
        opt2.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
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
t = Transaction(doc,'Brandschott')
t.Start()
step3 = int(len(Datenbank)/1000)+5
with forms.ProgressBar(title='{value}/{max_value} Rohre',cancellable=True, step=step3) as pb3:
    n_1 = 0
    for el in Datenbank:
        if pb3.cancelled:
            script.exit()
        n_1 += 1
        pb3.update_progress(n_1, len(Datenbank))
        el[0].LookupParameter('IGF_HLS_Brandschott').Set(int(el[1]))
t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))