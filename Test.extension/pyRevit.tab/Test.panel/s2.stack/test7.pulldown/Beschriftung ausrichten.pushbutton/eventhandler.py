# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import clr


class Filter1(Selection.ISelectionFilter):
    def AllowElement(self,element):

        if element.Category.Id.ToString() in ['-2008049','-2008055']:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class BESCHIFTUNG2(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        while(True):
            try:
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter1(),'Wählt die fixierte Beschriftung aus')
                el0 = doc.GetElement(el0_ref)
                el1_ref = uidoc.Selection.PickObjects(Selection.ObjectType.Element,Filter1(),'Wählt die zu veränderte MEP-Raumbeschriftung aus')
                
                t = DB.Transaction(doc,'anpassen')
                
                t.Start()
                for res in el1_ref:
                    try:
                        p0 = el0.Location.Point
                        el1 = doc.GetElement(res)
                        p1 = el1.Location.Point
                        if self.GUI.x.IsChecked:
                            el1.Location.Move(DB.XYZ(0,p0.Y-p1.Y,0))
                        elif self.GUI.y.IsChecked:
                            el1.Location.Move(DB.XYZ(p0.X-p1.X,0,0))
                        else:
                            el1.Location.Move(DB.XYZ(0,0,p0.Z-p1.Z))
                    except:
                        pass

                t.Commit()
                t.Dispose()
            except Exception as e:
                return
        

    def GetName(self):
        return "anpassen"