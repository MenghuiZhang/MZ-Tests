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
        if  element.Category.Id.IntegerValue == -2008055:
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
                        box_ = elem.get_BoundingBox(uidoc.ActiveView)
                        x0 = (box_.Max.X - box_.Min.X)/2
                        y0= (box_.Max.Y - box_.Min.Y)/2
                        box_.Dispose()
                        p = elem.Location.Point
                        
                        ref = DB.Reference(elem)
                        if self.GUI.U.IsChecked or self.GUI.D.IsChecked:
                            tag = DB.IndependentTag.Create(doc, DB.ElementId(37236383), uidoc.ActiveView.Id,  ref, False, DB.TagOrientation.Horizontal, p)
                        else:
                            tag = DB.IndependentTag.Create(doc, DB.ElementId(37236383), uidoc.ActiveView.Id,  ref, False, DB.TagOrientation.Vertical, p)
                        box = tag.get_BoundingBox(uidoc.ActiveView)
                        x = (box.Max.X - box.Min.X)/2
                        y = (box.Max.Y - box.Min.Y)/2
                        p_ = (box.Max + box.Min)/2

                        if self.GUI.L.IsChecked:
                            p0 = p + DB.XYZ(0-x0-x,0,0)
                        elif self.GUI.R.IsChecked:
                            p0 = p + DB.XYZ(x0+x,0,0)
                        elif self.GUI.U.IsChecked:
                            p0 = p + DB.XYZ(0,y+y0,0)
                        elif self.GUI.D.IsChecked:
                            p0 = p + DB.XYZ(0,0-y-y0,0)
                        tag.Location.Move(p0-p_)
                        box.Dispose()


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
