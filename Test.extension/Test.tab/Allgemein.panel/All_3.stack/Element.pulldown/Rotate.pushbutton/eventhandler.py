# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from math import pi

class Drehen(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if not self.GUI.winkel.Text:
            TaskDialog.Show('Info','Kein Winkel eingegeben!')
            return
        cl = uidoc.Selection.GetElementIds()
        if cl.Count == 0:
            TaskDialog.Show('Info','Kein Bauteil ausgewählt!')
            return
        
        t = DB.Transaction(doc,'Drehen')
        t.Start()
        
        if self.GUI.einzel.IsChecked:
            for el in cl:
                try:
                    elem = doc.GetElement(el)
                    transform = elem.GetTransform()
                    origin = transform.Origin 
                    if self.GUI.X_Achse.IsChecked:
                        line = DB.Line.CreateUnbound(origin,transform.BasisX)
                    elif self.GUI.Y_Achse.IsChecked:
                        line = DB.Line.CreateUnbound(origin,transform.BasisY)
                    elif self.GUI.Z_Achse.IsChecked:
                        line = DB.Line.CreateUnbound(origin,transform.BasisZ)
                    DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([el]), line,float(self.GUI.winkel.Text.replace(',','.'))/180.0*pi)
                except:pass
        else:
            max_x,max_y,max_z = -100000,-100000,-100000
            min_x,min_y,min_z = 100000,100000,100000
            for el in cl:
                try:
                    elem = doc.GetElement(el)
                    box = elem.get_BoundingBox(None)
                    max_x = max(max_x,box.Max.X)
                    max_y = max(max_y,box.Max.Y)
                    max_z = max(max_z,box.Max.Z)
                    min_x = min(min_x,box.Min.X)
                    min_y = min(min_y,box.Min.Y)
                    min_z = min(min_z,box.Min.Z)
                except:pass
                x0 = DB.XYZ((max_x+min_x)/2.0,(max_y+min_y)/2.0,(max_z+min_z)/2.0)
                if self.GUI.X_Achse.IsChecked:
                    line = DB.Line.CreateUnbound(x0,x0.BasisX)
                elif self.GUI.Y_Achse.IsChecked:
                    line = DB.Line.CreateUnbound(x0,x0.BasisY)
                elif self.GUI.Z_Achse.IsChecked:
                    line = DB.Line.CreateUnbound(x0,x0.BasisZ)
                DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([cl]), line,float(self.GUI.winkel.Text.replace(',','.'))/180.0*pi)

        t.Commit()
    
    def GetName(self):
        return "Drehen"
