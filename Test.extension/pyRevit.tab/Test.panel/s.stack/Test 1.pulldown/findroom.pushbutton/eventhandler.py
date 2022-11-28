# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 

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
            ##self.class_GUI.raumnummer.Text = el.Id.ToString()
        except:pass

    def GetName(self):
        return "Element anzeigen"