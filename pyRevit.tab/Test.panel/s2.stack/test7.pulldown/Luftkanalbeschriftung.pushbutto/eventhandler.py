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

class Filter2(Selection.ISelectionFilter):
    def AllowElement(self,element):

        if element.Category.Id.ToString() == '-2008044':
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
                el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter2(),'Wählt den Fix-Rohr aus')
                el0 = doc.GetElement(el0_ref)
                el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter2(),'Wählt die Rohr aus')
                
                t = DB.Transaction(doc,'anpassen')
                
                t.Start()
                
                try:
                    el1 = doc.GetElement(el1_ref)
                    el1.getParameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM).Set(el0.getParameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId())
                    el1.getParameter(DB.BuiltInParameter.RBS_OFFSET_PARAM).Set(el0.getParameter(DB.BuiltInParameter.RBS_OFFSET_PARAM).AsDouble())

                  
                except:pass

                t.Commit()
                t.Dispose()
            except Exception as e:
                return
        

    def GetName(self):
        return "anpassen"