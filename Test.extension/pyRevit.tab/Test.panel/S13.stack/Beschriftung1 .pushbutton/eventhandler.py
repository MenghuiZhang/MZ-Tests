# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import clr


class Filters(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if  element.Category.Name in ['Rohrzubehör','Luftkanalformteile' ]:
            return True
        # if element.Category.Id.ToString() == '-2000485':
        #     return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False
class Rohrfilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if  element.Category.Id.IntegerValue == -2008044:
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
        try:
            if element.Category.Id == self.elemid:
                return True
            else:
                return False
        except:
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
                elems = uidoc.Selection.PickElementsByRectangle(Rohrfilter())
                 
                t = DB.Transaction(doc,'anpassen')
                t.Start()
                try:
                    for elem in elems:
                        l = elem.Location.Curve
                        p0 =(l.GetEndPoint(0) + l.GetEndPoint(1))  /2
                        ref = DB.Reference(elem)
                        if self.GUI.H.IsChecked:
                            DB.IndependentTag.Create(doc, DB.ElementId(37159139), uidoc.ActiveView.Id,  ref, False, DB.TagOrientation.Horizontal, p0)
                        else:
                            DB.IndependentTag.Create(doc, DB.ElementId(37159139), uidoc.ActiveView.Id,  ref, False, DB.TagOrientation.Vertical, p0)

                except Exception as e:
                    print(e)
                    t.Commit()
                    break


                t.Commit()
                t.Dispose()
            except Exception as e:
                print(e)
                return
        

    def GetName(self):
        return "anpassen"