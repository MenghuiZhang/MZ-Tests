# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB


class LINKS(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        try:
            self.zahl = 0 - float(self.GUI.zahl.Text)
        except:
            TaskDialog.Show('Fehler','Zahl passt nicht!')
            return

        xyz = DB.XYZ(self.zahl/304.8,0,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()
        t.Dispose()
        uidoc.Dispose()
        doc.Dispose()

    def GetName(self):
        return "Links"

class UNTEN(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        try:
            self.zahl = 0 - float(self.GUI.zahl.Text)
        except:
            TaskDialog.Show('Fehler','Zahl passt nicht!')
            return
        xyz = DB.XYZ(0,0,self.zahl/304.8)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()
        t.Dispose()
        uidoc.Dispose()
        doc.Dispose()

    def GetName(self):
        return "Unten"

class HINTER(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        try:
            self.zahl = float(self.GUI.zahl.Text)
        except:
            TaskDialog.Show('Fehler','Zahl passt nicht!')
            return
        xyz = DB.XYZ(0,self.zahl/304.8,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()
        t.Dispose()
        uidoc.Dispose()
        doc.Dispose()

    def GetName(self):
        return "Hinter"

class RECHTS(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        try:
            self.zahl = float(self.GUI.zahl.Text)
        except:
            TaskDialog.Show('Fehler','Zahl passt nicht!')
            return
        xyz = DB.XYZ(self.zahl/304.8,0,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()
        t.Dispose()
        uidoc.Dispose()
        doc.Dispose()

    def GetName(self):
        return "Rechts"

class OBEN(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        try:
            self.zahl = float(self.GUI.zahl.Text)
        except:
            TaskDialog.Show('Fehler','Zahl passt nicht!')
            return
        xyz = DB.XYZ(0,0,self.zahl/304.8)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()
        t.Dispose()
        uidoc.Dispose()
        doc.Dispose()

    def GetName(self):
        return "Oben"

class VORNE(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
        try:
            self.zahl = 0 - float(self.GUI.zahl.Text)
        except:
            TaskDialog.Show('Fehler','Zahl passt nicht!')
            return
        xyz = DB.XYZ(0,self.zahl/304.8,0)

        t = DB.Transaction(doc,'move')
        t.Start()

        DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        t.Commit()
        t.Dispose()
        uidoc.Dispose()
        doc.Dispose()

    def GetName(self):
        return "Vorne"