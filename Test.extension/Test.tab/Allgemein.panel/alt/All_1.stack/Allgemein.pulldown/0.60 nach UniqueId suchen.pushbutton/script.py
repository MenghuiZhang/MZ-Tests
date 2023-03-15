# coding: utf8
from pyrevit.forms import WPFWindow
from pyIGF_logInfo import getlog

__title__ = "0.60 nach Uniqueid suchen"
__doc__ = """Element nach GUID suchen"""
__author__ = "Menghui Zhang"
getlog(__title__)
class Suche(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)

    def Anzeigen(self, sender, args):
        from pyrevit import revit, UI, DB
        uidoc = revit.uidoc
        doc = revit.doc
        GUID = self.guid.Text
        if not GUID:
            UI.TaskDialog.Show('Fehler:','Kein GUID')
        else:
            elem = doc.GetElement(GUID)
            if not elem:
                UI.TaskDialog.Show('Fehler:','falsche GUID')
            else:
                sel = uidoc.Selection.GetElementIds()
                sel.Clear()
                sel.Add(elem.Id)
                uidoc.Selection.SetElementIds(sel)
                uidoc.ShowElements(elem)

    def OK(self, sender, args):
        from pyrevit import revit, UI, DB
        uidoc = revit.uidoc
        doc = revit.doc
        GUID = self.guid.Text
        if not GUID:
            UI.TaskDialog.Show('Fehler:','Kein GUID')
        else:
            elem = doc.GetElement(GUID)
            if not elem:
                UI.TaskDialog.Show('Fehler:','Falsche GUID')
            else:
                sel = uidoc.Selection.GetElementIds()
                sel.Clear()
                sel.Add(elem.Id)
                uidoc.Selection.SetElementIds(sel)

    def Abbrechen(self, sender, args):
        self.Close()

Suche = Suche('window.xaml')
Suche.Show()
