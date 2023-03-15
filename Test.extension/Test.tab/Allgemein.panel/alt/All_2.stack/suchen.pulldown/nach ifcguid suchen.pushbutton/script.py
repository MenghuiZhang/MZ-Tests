# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit.forms import WPFWindow



__title__ = "nach IFC GUID suchen"
__doc__ = """Element nach ifcguid suchen

[2021.10.19]
Version: 1.0
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
            ifcguid = ''
            self.fehler.Visibility = Visibility.Hidden
            
            try:
                ifcguid = self.ifcguid.Text
            except:
                self.fehler.Text = 'ungültige IFC-GUID'
                self.fehler.Visibility = Visibility.Visible
                return
            if not ifcguid:
                self.fehler.Text = 'keine IFC-GUID'
                self.fehler.Visibility = Visibility.Visible
                return
            else:
                try:
                    param_equality=DB.FilterStringEquals()
                    param_id = DB.ElementId(DB.BuiltInParameter.IFC_GUID)
                    param_prov=DB.ParameterValueProvider(param_id)
                    param_value_rule=DB.FilterStringRule(param_prov,param_equality,ifcguid,True)
                    param_filter = DB.ElementParameterFilter(param_value_rule)
                    elemid = DB.FilteredElementCollector(doc).WherePasses(param_filter).WhereElementIsNotElementType().ToElementIds()[0]
                    elem = doc.GetElement(elemid)
                except:
                    self.fehler.Text = 'ungültige IFC-GUID'
                    self.fehler.Visibility = Visibility.Visible
                    return
                if not elem:
                    self.fehler.Text = 'falsche IFC-GUID'
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
            ifcguid = ''
            self.fehler.Visibility = Visibility.Hidden
            
            try:
                ifcguid = self.ifcguid.Text
            except:
                self.fehler.Text = 'ungültige IFC-GUID'
                self.fehler.Visibility = Visibility.Visible
                return
            if not ifcguid:
                self.fehler.Text = 'keine IFC-GUID'
                self.fehler.Visibility = Visibility.Visible
                return
            else:
                try:
                    param_equality=DB.FilterStringEquals()
                    param_id = DB.ElementId(DB.BuiltInParameter.IFC_GUID)
                    param_prov=DB.ParameterValueProvider(param_id)
                    param_value_rule=DB.FilterStringRule(param_prov,param_equality,ifcguid,True)
                    param_filter = DB.ElementParameterFilter(param_value_rule)
                    elemid = DB.FilteredElementCollector(doc).WherePasses(param_filter).WhereElementIsNotElementType().ToElementIds()[0]
                    elem = doc.GetElement(elemid)
                except:
                    self.fehler.Text = 'ungültige IFC-GUID'
                    self.fehler.Visibility = Visibility.Visible
                    return
                if not elem:
                    self.fehler.Text = 'falsche IFC-GUID'
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