# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class AlleFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008010' or element.Category.Id.ToString() == '-2008016'\
            or element.Category.Id.ToString() == '-2008049' or element.Category.Id.ToString() == '-2008055':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class LuftFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008010' or element.Category.Id.ToString() == '-2008016':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class RohrFilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008049' or element.Category.Id.ToString() == '-2008055':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class VERSCHIEBEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,AlleFilter(),'Wählt das fixierte Luftkanalformteil/-zubehör/Rohrformteil/-zubehör aus')
        el0 = doc.GetElement(el0_ref)

        if self.GUI.Abstand.Text == None or self.GUI.Abstand.Text == '':
            self.GUI.Abstand.Text = '10'
        try:
            abstand = float(self.GUI.Abstand.Text) / 304.8
        except:
            TaskDialog.Show('Fehler','ungültigen Abstand!')
            return

        if el0.Category.Id.ToString() in ['-2008010','-2008016']:
            el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,LuftFilter(),'Wählt das zu verschiebende Luftkanalformteil/-zubehör aus')
        else:
            el1_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,RohrFilter(),'Wählt das zu verschiebende Rohrformteil/-zubehör aus')

        el1 = doc.GetElement(el1_ref)

        distance = 100000000
        co0 = None
        co1 = None
        conns0 = list(el0.MEPModel.ConnectorManager.Connectors)
        conns1 = list(el1.MEPModel.ConnectorManager.Connectors)

        for con0 in conns0:
            for con1 in conns1:
                dis = con0.Origin.DistanceTo(con1.Origin)
                if dis < distance:
                    distance = dis
                    co0 = con0
                    co1 = con1
                    
        l0 = DB.Line.CreateBound(co0.Origin,co1.Origin)
        ln = l0.Direction.Normalize()
        neu = co0.Origin + ln * abstand

        pinned = el1.Pinned
        t = DB.Transaction(doc,'Abstand')
        t.Start()
        el1.Pinned = False
        doc.Regenerate()
        try:
            el1.Location.Move(neu - co1.Origin)
        except Exception as e:
            print(e)
        el1.Pinned = pinned

        t.Commit()

    def GetName(self):
        return "verschieben"