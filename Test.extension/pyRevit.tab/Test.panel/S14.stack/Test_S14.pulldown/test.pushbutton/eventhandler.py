# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import clr


class Filters(Selection.ISelectionFilter):
    def AllowElement(self,element):
        # if  element.Category.Name in ['Rohrzubehör','Luftkanalformteile' ]:
        #     return True
        if element.Category.Id.ToString() == '-2003600':
            return True
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
                elem = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filters())
                 
                t = DB.Transaction(doc,'MEPRaum')
                t.Start()
                space = doc.GetElement(elem)
                tag = doc.Create.NewSpaceTag(space,DB.UV.Zero,uidoc.ActiveView)
                doc.Regenerate()
                tag.get_Parameter(DB.BuiltInParameter.LEADER_LINE).Set(1)
                doc.Regenerate()
                tag.get_Parameter(DB.BuiltInParameter.LEADER_LINE).Set(0)


                t.Commit()
                t.Dispose()
            except Exception as e:
                return
        

    def GetName(self):
        return "anpassen"