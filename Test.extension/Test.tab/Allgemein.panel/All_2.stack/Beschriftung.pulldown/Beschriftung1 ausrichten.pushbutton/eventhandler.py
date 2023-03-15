# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import clr


class Filters(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if  element.GetType().Name == 'IndependentTag':
            return True
        elif element.GetType().BaseType.Name == 'SpatialElementTag':
            return True
        # if element.Category.Id.ToString() == '-2000485':
        #     return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter1(Selection.ISelectionFilter):
    def __init__(self,elemid):
        self.elemid = elemid
    def AllowElement(self,element):

        if element.Category.Id.ToString() == self.elemid:
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
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filters(),'Wählt die fixierte Beschriftung aus')
                el0 = doc.GetElement(el0_ref)
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter1(el0.Category.Id.ToString()),'Wählt die zu veränderte MEP-Raumbeschriftung aus')
                
                t = DB.Transaction(doc,'anpassen')
                
                t.Start()
            
                el1 = doc.GetElement(el1_ref)
                try:
                    p0 = el0.Location.Point
               
                    p1 = el1.Location.Point
                    if self.GUI.x.IsChecked:
                        el1.Location.Move(DB.XYZ(0,p0.Y-p1.Y,0))
                    elif self.GUI.y.IsChecked:
                        el1.Location.Move(DB.XYZ(p0.X-p1.X,0,0))
                    else:
                        el1.Location.Move(DB.XYZ(0,0,p0.Z-p1.Z))
                except:
                    try:
                        p0 = el0.TagHeadPosition
                        p1 = el1.TagHeadPosition
                        if self.GUI.x.IsChecked:
                            el1.Location.Move(DB.XYZ(0,p0.Y-p1.Y,0))
                        elif self.GUI.y.IsChecked:
                            el1.Location.Move(DB.XYZ(p0.X-p1.X,0,0))
                        else:
                            el1.Location.Move(DB.XYZ(0,0,p0.Z-p1.Z))
                    except:pass

                t.Commit()
                t.Dispose()
            except Exception as e:
                return
        

    def GetName(self):
        return "anpassen"