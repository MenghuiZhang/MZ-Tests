# coding: utf8
from pyrevit.forms import WPFWindow
#from pyrevit import revit
#doc = revit.doc
from pyIGF_logInfo import getlog

__title__ = "0.61 nach ElementId suchen"
__doc__ = """Element nach ElementId suchen"""
__author__ = "Menghui Zhang"

getlog(__title__)#,doc)

class Suche(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)

    def Anzeigen(self, sender, args):
        from pyrevit import revit, UI, DB
        uidoc = revit.uidoc
        doc = revit.doc
        GUID = self.guid.Text
        if not GUID:
            UI.TaskDialog.Show('Fehler:','Kein ElementId')
        else:
            elem = doc.GetElement(DB.ElementId(int(GUID)))
            if not elem:
                UI.TaskDialog.Show('Fehler:','falsche ElementId')
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
            UI.TaskDialog.Show('Fehler:','Kein ElementId')
        else:
            elem = doc.GetElement(DB.ElementId(int(GUID)))
            if not elem:
                UI.TaskDialog.Show('Fehler:','Falsche ElementId')
            else:
                sel = uidoc.Selection.GetElementIds()
                sel.Clear()
                sel.Add(elem.Id)
                uidoc.Selection.SetElementIds(sel)

    def Abbrechen(self, sender, args):
        self.Close()

Suche = Suche('window.xaml')
Suche.Show()
