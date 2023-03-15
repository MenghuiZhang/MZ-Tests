# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import time
from pyrevit import script, forms
from rpw import *
import System
from Autodesk.Revit.DB import *


start = time.time()

__title__ = "9.80 Brandschutz"
__doc__ = """Brandschutz"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

revitLinks_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

pipes_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()
pipes = pipes_collector.ToElementIds()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el

rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
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
    m_lstModels = [[],[],[]]
    listSolids = getSolids(elem)
    for solid in listSolids:
        tempSolid = solid
        tempSolid = SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
        m_lstModels[0].append(tempSolid)
        for i in tempSolid.Faces:
            m_lstModels[1].append(i)
        for j in tempSolid.Edges:
            m_lstModels[2].append(j.AsCurve())
    return m_lstModels

def TransformSolidPro(elem):
    m_lstModels = [[],[],[]]
    listSolids = getSolids(elem)
    for solid in listSolids:
        tempSolid = solid
        m_lstModels[0].append(tempSolid)
        for i in tempSolid.Faces:
            m_lstModels[1].append(i)
        for j in tempSolid.Edges:
            m_lstModels[2].append(j.AsCurve())
    return m_lstModels

# RvtLinkElem
RvtLinkElemSolids = []
ProElemSolids = []
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


def ProKurve(elem):
    pipecurve = None
    csi = elem.ConnectorManager.Connectors.ForwardIterator()
    list = []
    while csi.MoveNext():
        conn = csi.Current
        list.append(conn.Origin)
    pipecurve = Line.CreateBound(list[0], list[1])
    return pipecurve

# Projektelem
step1 = int(len(pipes)/200)
with forms.ProgressBar(title='{value}/{max_value} Rohre in aktueller Projekt',cancellable=True, step=step1) as pb1:
    n_1 = 0
    for ele in pipes_collector:
        if pb1.cancelled:
            script.exit()
        n_1 += 1
        pb1.update_progress(n_1, len(pipes))
        models = ProKurve(ele)
        ProElemSolids.append([ele,models])

# Kollisionprüfung
Datenbank = []
step2 = int(len(ProElemSolids)/1000)
with forms.ProgressBar(title='{value}/{max_value} Rohre in aktueller Projekt',cancellable=True, step=step2) as pb2:
    n_1 = 0
    for Item in ProElemSolids:
        n = 0
        pipecurve = Item[1]
        if pb2.cancelled:
            script.exit()
        n_1 += 1
        pb2.update_progress(n_1, len(ProElemSolids))
        for item in RvtLinkElemSolids:
            n1 = 0
            elelink = item[1]
            for solid in elelink[0]:
                opt = SolidCurveIntersectionOptions()
                result = solid.IntersectWithCurve(pipecurve,opt)
                if result.SegmentCount > 0:
                    n1 = n1 + 1
                    break
                # else:
                #     for face in solid.Faces:
                #         result2 = face.Intersect(pipecurve)
                #         if result2.ToString() != 'Disjoint':
                #             n1 = n1 + 1
                #             break
            if n1 > 0:
                n1 = 1
            n = n + n1
        if n> 0:
            Datenbank.append([Item[0],n])

# for Item in ProElemSolids:
#     pipecurve = Item[1]
#     n = 0
#     for item in RvtLinkElemSolids:
#         n1 = 0
#         elelink = item[1]
#         for solid in elelink[0]:
#             opt = SolidCurveIntersectionOptions()
#             result = solid.IntersectWithCurve(pipecurve,opt)
#             if result.SegmentCount > 0:
#                 n1 = n1 + 1
#                 break
#             else:
#                 for face in solid.Faces:
#                     result2 = face.Intersect(pipecurve)
#                     if result2.ToString() != 'Disjoint':
#                         n1 = n1 + 1
#                         break
#         if n1 > 0:
#             n1 = 1
#         n = n + n1
#     if n> 0:
#         print(Item[0],n)



        #                 else:
        #                     for face in solid.Faces:
        #                         array = clr.Reference[IntersectionResultArray]()
        #                         result2 = face.Intersect(curve,array)
        #                         if array != None and result2.ToString() != 'Disjoint':
        #                             n1 = n1 + 1
        #                             break
        #                 if n1 > 0:
        #                     break
        # filter = ElementIntersectsSolidFilter(solid)
        # ids = pipes_collector.WherePasses(filter).ToElementIds()
        # # filter.Dispose()
        # print(ids)
        # print(30*'--')

# for item in ProElemSolids:
#     elePro = item[1]
#     n = 0
#     for Item in RvtLinkElemSolids:
#         n1 = 0
#         elelink = Item[1]
        # for Psolid in elePro[0]:
        #     for Rsolid in elelink[0]:
        #         Pfaces = Psolid.Faces
        #         Rfaces = Rsolid.Faces
        #         for Pface in Pfaces:
        #             if Pface.GetType().ToString() != 'Autodesk.Revit.DB.CylindricalFace':
        #                 continue
        #             for Rface in Rfaces:
        #                 if Rface.FaceNormal.IsAlmostEqualTo(XYZ(0, 0, 1)) or Rface.FaceNormal.IsAlmostEqualTo(XYZ(0, 0, -1)):
        #                     continue
        #                 result = Pface.Intersect(Rface)
        #                 if result == FaceIntersectionFaceResult.Intersecting:
        #                     n1 = n1 + 1
        # if n1 > 0:
        #     n1 = 1
        #
        # n = n + n1
    #     for curve in elePro[2]:
    #         opt = SolidCurveIntersectionOptions()
    #         for solid in elelink[0]:
    #             result = solid.IntersectWithCurve(curve,opt)
    #             if result.SegmentCount > 0:
    #                 n1 = n1 + 1
    #                 break
    #             else:
    #                 for face in solid.Faces:
    #                     array = clr.Reference[IntersectionResultArray]()
    #                     result2 = face.Intersect(curve,array)
    #                     if array != None and result2.ToString() != 'Disjoint':
    #                         n1 = n1 + 1
    #                         break
    #             if n1 > 0:
    #                 break
    #     if n1 > 0:
    #         n1 = 1
    #     # if n1 > 0:
    #     #     continue
    #     for Pface in elePro[1]:
    #         if Pface.GetType().ToString() != 'Autodesk.Revit.DB.CylindricalFace':
    #             continue
    #         for Rface in elelink[1]:
    #             if Rface.FaceNormal.IsAlmostEqualTo(XYZ(0, 0, 1)) or Rface.FaceNormal.IsAlmostEqualTo(XYZ(0, 0, -1)):
    #                 continue
    #             result = Pface.Intersect(Rface)
    #             if result == FaceIntersectionFaceResult.Intersecting:
    #                 n1 = n1 + 1
    #                 break
    #     n = n + n1
    # print(n)









total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
