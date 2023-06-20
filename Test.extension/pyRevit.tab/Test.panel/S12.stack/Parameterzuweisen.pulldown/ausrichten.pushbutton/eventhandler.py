# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import clr
from IGF_lib import get_value


class BESCHIFTUNG2(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        n = 9
        while(True):
            try:
                elem = doc.GetElement(uidoc.Selection.PickObject(Selection.ObjectType.Element))
                # obj0 = self.GUI.object0.Text
                obj0 = 'RVH_140.10' + str(n)
                n-=1
            
         


                t = DB.Transaction(doc,'massen')
                t.Start()
                elem.LookupParameter('ZS_MARK').Set(obj0)
                t.Commit()
                t.Dispose()
            except:
                return

    def GetName(self):
        return "anpassen"

class Aktualisieren(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        elids = uidoc.Selection.GetElementIds()

        if elids.Count > 0:
            for _id in elids:

                try:self.GUI.object0.Text = get_value(doc.GetElement(_id).LookupParameter('ZS_MARK'))
                except Exception as e:print(e)                    
      
        

    def GetName(self):
        return "anpassen"