# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit.forms import WPFWindow



__title__ = "nach ElementId suchen"
__doc__ = """Element nach ElementId suchen

[2021.10.19]
Version: 2.0
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

class Suche(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)

    def Anzeigen(self, sender, args):
        try:
            from rpw import revit,DB
            from System.Windows import Visibility 
            uidoc = revit.uidoc
            doc = revit.doc
            elementId = ''
            self.fehler.Visibility = Visibility.Hidden
            try:
                elementId = self.elemetid.Text
            except:
                self.fehler.Text = 'ungültige ElementId'
                self.fehler.Visibility = Visibility.Visible
                return
            if not elementId:
                self.fehler.Text = 'keine ElementId'
                self.fehler.Visibility = Visibility.Visible
                return
            else:
                try:
                    elem = doc.GetElement(DB.ElementId(int(elementId)))
                except:
                    self.fehler.Text = 'ungültige ElementId'
                    self.fehler.Visibility = Visibility.Visible
                    return
                if not elem:
                    self.fehler.Text = 'falsche ElementId'
                    self.fehler.Visibility = Visibility.Visible
                else:
                    sel = uidoc.Selection.GetElementIds()
                    sel.Clear()
                    sel.Add(elem.Id)
                    uidoc.Selection.SetElementIds(sel)
                    uidoc.ShowElements(elem)
        except:
            self.Close()

    def OK(self, sender, args):
        try:
            from rpw import revit,DB
            from System.Windows import Visibility 
            uidoc = revit.uidoc
            doc = revit.doc
            elementId = ''
            self.fehler.Visibility = Visibility.Hidden
            try:
                elementId = self.elemetid.Text
            except:
                self.fehler.Text = 'ungültige ElementId'
                self.fehler.Visibility = Visibility.Visible
                return
            if not elementId:
                self.fehler.Text = 'keine ElementId'
                self.fehler.Visibility = Visibility.Visible
                return
            else:
                try:
                    elem = doc.GetElement(DB.ElementId(int(elementId)))
                except:
                    self.fehler.Text = 'ungültige ElementId'
                    self.fehler.Visibility = Visibility.Visible
                    return
                if not elem:
                    self.fehler.Text = 'falsche ElementId'
                    self.fehler.Visibility = Visibility.Visible
                else:
                    sel = uidoc.Selection.GetElementIds()
                    sel.Clear()
                    sel.Add(elem.Id)
                    uidoc.Selection.SetElementIds(sel)
        except:
            self.Close()
                    

    def Abbrechen(self, sender, args):
        self.Close()

Suche = Suche('window.xaml')
Suche.Show()