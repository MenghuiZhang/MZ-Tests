# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB

class ListeExternalEvent(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        for elid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(elid)
            if el.Category.Id.IntegerValue in [-2008000]:
                elems.append(el)
        if len(elems) == 0:
            TaskDialog.Show('Info','Kein Luftkanal ausgewählt!')
        t = DB.Transaction(doc,'Kanäle anpassen')
        t.Start()
        for el in elems:
            try:
                if self.GUI.breite.SelectedIndex != -1:
                    el.LookupParameter('Breite').SetValueString(self.GUI.breite.SelectedItem.ToString())
                if self.GUI.Hoehe.SelectedIndex != -1:
                    el.LookupParameter('Höhe').SetValueString(self.GUI.Hoehe.SelectedItem.ToString())
            except:
                try:
                    if self.GUI.Durchmesser.SelectedIndex != -1:
                        el.LookupParameter('Durchmesser').SetValueString(self.GUI.Durchmesser.SelectedItem.ToString())
                except Exception as e:
                    print(e)
        t.Commit()

    def GetName(self):
        return 'Kanalrechnenr'

    