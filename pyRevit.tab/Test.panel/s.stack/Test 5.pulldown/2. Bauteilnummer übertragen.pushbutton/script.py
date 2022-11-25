# coding: utf8
from IGF_log import getlog,getloglocal


from rpw import revit,DB,UI


__title__ = "Element-ID schreiben "
__doc__ = """ Schreibt die Element-ID an alle Bauteile, die sich in der Ansicht (Grundriss/Bauteilliste) befinden. 

"""

__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass




uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp
active_view = uidoc.ActiveView



Bauteile = DB.FilteredElementCollector(doc,active_view.Id).ToElements()

t = DB.Transaction(doc, 'ElementId schreiben')
t.Start()
for el in Bauteile:
    el.LookupParameter('IGF_X_ElementId').Set(el.Id.ToString())
t.Commit()
