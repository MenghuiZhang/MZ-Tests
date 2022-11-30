# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles, Visibility
from System.Windows.Media import Brushes, BrushConverter
from System.Windows.Forms import OpenFileDialog,DialogResult
from System.Collections.Generic import List
from rpw import revit, DB, UI
from pyrevit import script, forms

doc = revit.doc
uidoc = revit.uidoc
uiapp = revit.uiapp
app = revit.app
BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter
Visible = Visibility.Visible
Hidden = Visibility.Hidden
Collector = DB.FilteredElementColector
Transaction = DB.Transaction