# coding: utf8
from xml.dom.xmlbuilder import DOMEntityResolver
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection,RevitCommandId
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 
from Autodesk.Revit.UI.Selection import ISelectionFilter
import time

class MEPRaumFilter(ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2003600':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

# class ANZEIGEN(IExternalEventHandler):
#     def __init__(self):
#         self.nr = ''
            
#     def Execute(self,app):
#         uidoc = app.ActiveUIDocument
#         doc = uidoc.Document
#         re = uidoc.Selection.PickObject(Selection.ObjectType.Element,MEPRaumFilter())
#         elem = doc.GetElement(re)
#         self.nr = elem.Number + ' - ' + elem.LookupParameter('Name').AsString()
#         print(self.nr )
        
#     def GetName(self):
#         return "Show Element"


class ANZEIGEN(IExternalEventHandler):
    def __init__(self):
        self.class_GUI = None

        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        el = doc.GetElement(uidoc.Selection.GetElementIds()[0])
        phase = doc.GetElement(el.CreatedPhaseId)
        try:
            self.class_GUI.raumnummer.Text = el.Space[phase].Number + ' - ' +el.Space[phase].LookupParameter('Name').AsString()
        except:pass

    def GetName(self):
        return "Element anzeigen"

    # def DoMyWork(self,app):
    #     if self.m_nStep == 0:
    #         id = RevitCommandId.LookupCommandId( "b56a9b4b-27d7-4f41-aebf-deaa3b2ac674" )
    #         if id != None and app != None:
    #             app.PostCommand(id)
    #     self.NextStep()
    
    # def NextStep(self):
    #     self.m_dtStamp = self.time.time()
    #     self.m_nStep += 1
    
    # def OnIdle(self, sender, e):
    #     app = sender
    #     span = self.time.time() - self.m_dtStamp
    #     if 0.5 > span:return
    #     if self.m_nStep == 1:
    #         self.DoMyWork(app)
    #         return
    #     if self.m_nStep == 2:
    #         self.DoMyWork(app)
    #         return
    #     if self.m_nStep == 3:
    #         self.DoMyWork(app)
    #         app.Idling -= self.OnIdle
    #         return