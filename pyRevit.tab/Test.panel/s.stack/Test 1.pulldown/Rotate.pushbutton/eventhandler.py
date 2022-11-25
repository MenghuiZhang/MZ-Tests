# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List


class XRotate(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        if cl.Count == 0:
            return
    
        t = DB.Transaction(doc,'rotate')
        t.Start()
        

        for el in cl:
            try:
                elem = doc.GetElement(el)
                transform = elem.GetTransform()
                origin = transform.Origin

                line = DB.Line.CreateUnbound(origin,transform.BasisX)
                DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([el]), line,self.zahl)
            except:pass

        t.Commit()

    def GetName(self):
        return "XRotate"



class YRotate(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        if cl.Count == 0:
            return
    
        t = DB.Transaction(doc,'rotate')
        t.Start()
        

        for el in cl:
            try:
                elem = doc.GetElement(el)
                transform = elem.GetTransform()
                origin = transform.Origin

                line = DB.Line.CreateUnbound(origin,transform.BasisY)
                DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([el]), line,self.zahl)
            except:pass

        t.Commit()

    def GetName(self):
        return "YRotate"

class ZRotate(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        if cl.Count == 0:
            return
    
        t = DB.Transaction(doc,'rotate')
        t.Start()
        

        for el in cl:
            try:
                elem = doc.GetElement(el)
                transform = elem.GetTransform()
                origin = transform.Origin

                line = DB.Line.CreateUnbound(origin,transform.BasisZ)
                DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([el]), line,self.zahl)
            except Exception as e:print(e)

        t.Commit()
    def GetName(self):
        return "ZRotate"
