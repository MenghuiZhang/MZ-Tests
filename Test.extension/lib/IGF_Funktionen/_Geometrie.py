# coding: utf8
import clr
from Autodesk.Revit.DB import Line, XYZ
from Autodesk.Revit.DB import SetComparisonResult, IntersectionResultArray,ClosestPointsPairBetweenTwoCurves
from IGF_Klasse import List

import sys
import clr
import os
sys.path.Add(os.path.dirname(__file__)+'\\External')
clr.AddReference('GeometrieExternal')
import GeometrieExternal


def get_Einbauort(elem):
    mep = elem.Space[elem.Document.GetElement(elem.CreatedPhaseId)]
    if mep:
        return mep.Number
    else:
        return None

def LinkedPointTransform(revitlink,Point):
    return revitlink.GetTransform().Inverse.OfPoint(Point)

def LinkedVectorTransform(revitlink,Vector):
    return revitlink.GetTransform().Inverse.OfPoint(Vector)

def CurrentPointTransformToLink(revitlink,Point):
    return revitlink.GetTransform().OfPoint(Point)

def CurrentVectorTransformToLink(revitlink,Vector):
    return revitlink.GetTransform().OfPoint(Vector)

def get_intersection(line1, line2):
    '''Get Line Intersection'''
    results = clr.Reference[IntersectionResultArray]()
    result = line1.Intersect(line2, results)
    if result != SetComparisonResult.Overlap:
        print('No Intesection')
        return
    if results is None or results.Size != 1:
        print("Could not extract line intersection point." )
        return


    intersection = results.Item[0]
    return intersection.XYZPoint

def get_ClosestPoints(line1, line2):
    '''return ClosestPointsPairBetweenTwoCurves
    PointOnFirstCurve:get_ClosestPoints(line1, line2).XYZPointOnFirstCurve 
    PointOnSecondCurve:get_ClosestPoints(line1, line2).XYZPointOnSecondCurve 
    Distance:get_ClosestPoints(line1, line2).Distance 
    
    '''
    return GeometrieExternal.GetClosedPoints.GetClosedPoints_(line1,line2)


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
    opt.Dispose()
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
