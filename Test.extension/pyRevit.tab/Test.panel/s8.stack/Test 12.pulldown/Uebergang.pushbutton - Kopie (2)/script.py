# coding: utf8
# from IGF_log import getlog,getloglocal
# from pyrevit import forms
# from eventhandler import TransitionFitting,ExternalEvent,TransitionFitting1
# import os
# from System.Windows.Input import Key
from rpw import revit,UI,DB
import System
from System.Collections.Generic import List

__title__ = "Origintest"
__doc__ = """

[2022.08.01]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

reference = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element)
elem = doc.GetElement(reference.ElementId)

try:
    conns = elem.MEPModel.ConnectorManager.Connectors
except:
    try:
        conns = elem.ConnectorManager.Connectors
    except:
        sys.exit()

r = DB.Color(System.Byte(255),System.Byte(0),System.Byte(0))
g = DB.Color(System.Byte(0),System.Byte(255),System.Byte(0))
b = DB.Color(System.Byte(0),System.Byte(0),System.Byte(255))

t = DB.Transaction(doc,'Test')
t.Start()

def createshape(line):
    shape = DB.DirectShape.CreateElement(doc,DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
    shape.SetShape(List[DB.GeometryObject](line))
    return shape

for conn in conns:
    transform = conn.CoordinateSystem
    origin = transform.Origin
    basis_x = transform.BasisX
    basis_y = transform.BasisY
    basis_z = transform.BasisZ

    line_x = DB.Line.CreateBound(origin,origin+basis_x) 
    line_y = DB.Line.CreateBound(origin,origin+basis_y) 
    line_z = DB.Line.CreateBound(origin,origin+basis_z) 
    shape = createshape(line_x)
    graph = DB.OverrideGraphicSettings()
    graph.SetProjectLineColor(r)
    graph.SetProjectLineWeight(4)
    uidoc.ActiveView.SetElementOverrides(shape.Id,graph)

    shape = createshape(line_y)
    graph = DB.OverrideGraphicSettings()
    graph.SetProjectLineColor(g)
    graph.SetProjectLineWeight(4)
    uidoc.ActiveView.SetElementOverrides(shape.Id,graph)

    shape = createshape(line_z)
    graph = DB.OverrideGraphicSettings()
    graph.SetProjectLineColor(b)
    graph.SetProjectLineWeight(4)
    uidoc.ActiveView.SetElementOverrides(shape.Id,graph)





t.Commit()


