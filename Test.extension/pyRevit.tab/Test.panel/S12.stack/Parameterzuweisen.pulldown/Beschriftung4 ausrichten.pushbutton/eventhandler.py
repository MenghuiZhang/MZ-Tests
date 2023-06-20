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
        if elids.Count > 0:

          
            t = DB.Transaction(doc,'massen')
            t.Start()
            for _id in elids:
                try:doc.GetElement(_id).LookupParameter('ZS_MARK').Set(self.GUI.mark.Text)
                except Exception as e:print(e)
                try:doc.GetElement(_id).LookupParameter('MC Object Variable 1').Set(self.GUI.object1.Text)
                except Exception as e:print(e)
                try:doc.GetElement(_id).LookupParameter('MC Object Variable 2').Set(self.GUI.object2.Text)
                except Exception as e:print(e)
                try:doc.GetElement(_id).LookupParameter('MC Object Variable 3').Set(self.GUI.object3.Text)
                except Exception as e:print(e)
                try:doc.GetElement(_id).LookupParameter('MC Object Variable 4').Set(self.GUI.object4.Text)
                except Exception as e:print(e)
                try:doc.GetElement(_id).LookupParameter('MC Piping Flow').SetValueString(self.GUI.massen.Text) 
                except Exception as e:print(e)
                try:doc.GetElement(_id).LookupParameter('MC Piping Power').SetValueString(self.GUI.leistung0.Text)
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
                try:self.GUI.mark.Text = str(get_value(doc.GetElement(_id).LookupParameter('ZS_MARK')))
                except Exception as e:print(e)
                try:self.GUI.object1.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 1'))
                except Exception as e:print(e)
                try:self.GUI.object2.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 2'))
                except Exception as e:print(e)
                try:self.GUI.object3.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 3'))
                except Exception as e:print(e)
                try:self.GUI.object4.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 4'))
                except Exception as e:print(e)
                try:self.GUI.massen.Text = str(get_value(doc.GetElement(_id).LookupParameter('MC Piping Flow')))
                except Exception as e:print(e)
                try:self.GUI.leistung0.Text = str(get_value(doc.GetElement(_id).LookupParameter('MC Piping Power')))
                except Exception as e:print(e)
                
                
               
                    
      
        

    def GetName(self):
        return "anpassen"