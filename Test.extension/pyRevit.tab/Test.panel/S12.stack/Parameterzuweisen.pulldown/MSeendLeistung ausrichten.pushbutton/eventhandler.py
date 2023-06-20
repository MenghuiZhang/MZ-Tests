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
        elids = uidoc.Selection.GetElementIds()

        massen = self.GUI.massen.Text
        leistung = self.GUI.leistung.Text
        if elids.Count > 0:

            if massen:
                t = DB.Transaction(doc,'massen')
                t.Start()
                for _id in elids:
                    try:doc.GetElement(_id).LookupParameter('MC Piping Power').SetValueString(leistung)
                    except Exception as e:print(e)     
                    try:doc.GetElement(_id).LookupParameter('MC Piping Flow').SetValueString(massen)
                    except Exception as e:print(e)                  
                    
                t.Commit()
        

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
                try:self.GUI.massen.Text = str(doc.GetElement(_id).LookupParameter('MC Piping Flow').AsValueString())
                except Exception as e:print(e)
                try:self.GUI.leistung.Text = str(doc.GetElement(_id).LookupParameter('HA_POWER').AsValueString())
                except Exception as e:print(e)
            
                    
      
        

    def GetName(self):
        return "anpassen"