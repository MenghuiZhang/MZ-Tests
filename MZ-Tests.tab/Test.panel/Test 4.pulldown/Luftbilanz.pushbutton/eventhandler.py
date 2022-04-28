# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler
import Autodesk.Revit.DB as DB

class ResizeAnsicht(IExternalEventHandler):
    def __init__(self):
        self.box = None
        self.XYZAnpassen = DB.XYZ(1/0.3048,1/0.3048,1/0.3048)
        
    def Execute(self,app):
        doc = app.ActiveUIDocument.Document
        uidoc = app.ActiveUIDocument
        view = uidoc.ActiveView
        t = DB.Transaction(doc,'modifiziert Ansichtsgröße')
        t.Start()
        view.IsSectionBoxActive = True
        coll = DB.FilteredElementCollector(doc,view.Id).OfCategory(DB.BuiltInCategory.OST_SectionBox).WhereElementIsNotElementType().ToElements()
        for el in coll:
            if el.LookupParameter('Bearbeitungsbereich').AsValueString().find(view.Name) != -1:box0 = el.get_BoundingBox(view)
        box = view.GetSectionBox()
        Max = self.box.Max + self.XYZAnpassen + box.Max-box0.Max
        Min = self.box.Min - self.XYZAnpassen + box.Max-box0.Max
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)

        t.Commit()
        t.Dispose()
    def GetName(self):
        return "modifiziert Ansichtsgröße"

class ResetAnsicht(IExternalEventHandler):
    def __init__(self):
        self.boxmax = None
        self.boxmin = None

    def Execute(self,app):
        doc = app.ActiveUIDocument.Document
        uidoc = app.ActiveUIDocument
        view = uidoc.ActiveView
        t = DB.Transaction(doc,'setzt Ansichtsgröße zurück')
        t.Start()
        view.IsSectionBoxActive = True
        box = view.GetSectionBox()
        box.Max = self.boxmax
        box.Min = self.boxmin
        view.SetSectionBox(box)

        t.Commit()
        t.Dispose()
    def GetName(self):
        return "setzt Ansichtsgröße zurück"