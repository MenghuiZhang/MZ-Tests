# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class HLSFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2001140':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class RohrFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class VERBINDEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        cl = uidoc.Selection.GetElementIds()
        rohr = []
        formteil = []
        for elid in cl:
            el = doc.GetElement(elid)
            cate = el.Category.Name
            if cate == 'Rohre':rohr.append(el)
            elif cate == 'Rohrformteile':formteil.append(el)
        t = DB.Transaction(doc,'Dimension')
        t.Start()
        for el in rohr:
            el.LookupParameter('Durchmesser').SetValueString('32')
        for el in formteil:
            try:
                el.LookupParameter('Nennradius').SetValueString('16')
            except:
                try:
                    el.LookupParameter('MC_R').SetValueString('16')
                except:
                    pass
        t.Commit()
        t.Dispose()


    def GetName(self):
        return "Verbinden"