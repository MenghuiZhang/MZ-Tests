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

        massen = self.GUI.massen.Text
        obj0 = self.GUI.object0.Text
        obj2 = self.GUI.object2.Text
        obj3 = self.GUI.object3.Text
        l0 = self.GUI.leistung0.Text
        l = self.GUI.leistung.Text
        tem1 = self.GUI.tem1.Text
        tem2 = self.GUI.tem2.Text
        druck = self.GUI.druck.Text
        elids = uidoc.Selection.GetElementIds()
        if elids.Count > 0:

            if massen:
                t = DB.Transaction(doc,'massen')
                t.Start()
                for _id in elids:
                    try:doc.GetElement(_id).LookupParameter('MC Piping Flow').SetValueString(str(round(float(massen)/3600.0,10)))
                    except Exception as e:print(e)
                    try:doc.GetElement(_id).LookupParameter('MC Object Variable 1').Set(obj0)
                    except Exception as e:print(e)
                    # try:doc.GetElement(_id).LookupParameter('MC Object Variable 3').Set(obj2)
                    # except Exception as e:print(e)
                    # try:doc.GetElement(_id).LookupParameter('MC Object Variable 4').Set('PN16')
                    # except Exception as e:print(e)
                    # try:doc.GetElement(_id).LookupParameter('MC Piping Power').SetValueString(str(float(l0)*1000.0))
                    # except Exception as e:
                    #     try:doc.GetElement(_id).LookupParameter('MC Piping Power').ClearValue()
                    #     except Exception as e:print(e)
                    try:doc.GetElement(_id).LookupParameter('MC Piping Power').SetValueString(str(float(l)))
                    except Exception as e:print(e)
                    # try:doc.GetElement(_id).LookupParameter('HA_MD_PRT').SetValueString(tem1)
                    # except Exception as e:print(e)
                    # try:doc.GetElement(_id).LookupParameter('HA_MD_PST').SetValueString(tem2)
                    # except Exception as e:print(e)
                    # try:doc.GetElement(_id).LookupParameter('MC Piping Pressure Drop').SetValueString(str(round(float(druck),10)))
                    # except Exception as e:print(e)
                    
                    
                t.Commit()
        

    def GetName(self):
        return "anpassen"

class Aktualisieren(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        massen = self.GUI.massen.Text
        obj0 = self.GUI.object0.Text
        obj2 = self.GUI.object2.Text
        obj3 = self.GUI.object3.Text
        l0 = self.GUI.leistung0.Text
        l = self.GUI.leistung.Text
        tem1 = self.GUI.tem1.Text
        tem2 = self.GUI.tem2.Text
        massen = self.GUI.massen.Text
        elids = uidoc.Selection.GetElementIds()
        if elids.Count > 0:
            for _id in elids:
                try:self.GUI.massen.Text = str(get_value(doc.GetElement(_id).LookupParameter('MC Piping Flow')) * 3600.0 )
                except Exception as e:print(e)
                try:self.GUI.object0.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 1'))
                except Exception as e:print(e)
                try:self.GUI.object2.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 3'))
                except Exception as e:print(e)
                try:self.GUI.object3.Text = get_value(doc.GetElement(_id).LookupParameter('MC Object Variable 4'))
                except Exception as e:print(e)
                # try:self.GUI.leistung0.Text = str(get_value(doc.GetElement(_id).LookupParameter('MC Piping Power'))/1000)
                # except Exception as e:print(e)
                try:self.GUI.leistung.Text = str(get_value(doc.GetElement(_id).LookupParameter('HA_POWER')))
                except Exception as e:print(e)
                try:self.GUI.tem1.Text = str(get_value(doc.GetElement(_id).LookupParameter('HA_MD_PRT')))
                except Exception as e:print(e)
                try:self.GUI.tem2.Text = str(get_value(doc.GetElement(_id).LookupParameter('HA_MD_PST')))
                except Exception as e:print(e)
                try:self.GUI.druck.Text = str(get_value(doc.GetElement(_id).LookupParameter('MC Piping Pressure Drop')))
                except Exception as e:print(e)
                    
      
        

    def GetName(self):
        return "anpassen"