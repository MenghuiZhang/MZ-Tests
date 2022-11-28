# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB


class LINKS(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        xyz = DB.XYZ(self.zahl/304.8,0,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()

    def GetName(self):
        return "Links"



class SHANG(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        xyz = DB.XYZ(0,0,self.zahl/304.8)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()

    def GetName(self):
        return "SHANG"

class XIA(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        xyz = DB.XYZ(0,0,self.zahl/304.8)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()

    def GetName(self):
        return "XIA"



class RECHTS(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        xyz = DB.XYZ(self.zahl/304.8,0,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()

    def GetName(self):
        return "Rechts"


class OBEN(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        self.move = DB.ElementTransformUtils
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        xyz = DB.XYZ(0,self.zahl/304.8,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()

    def GetName(self):
        return "Oben"

class UNTEN(IExternalEventHandler):
    def __init__(self):
        self.elems = None
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        xyz = DB.XYZ(0,self.zahl/304.8,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()

    def GetName(self):
        return "Unten"