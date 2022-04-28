# coding: utf8
from rpw import revit, DB, UI

doc = revit.doc
uidoc = revit.uidoc

a = uidoc.Selection.PickObject(UI.Selection.Edge)
print(a)