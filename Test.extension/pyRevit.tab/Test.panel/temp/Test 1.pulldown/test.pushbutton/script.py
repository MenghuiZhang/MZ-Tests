# coding: utf8
from pyrevit import revit, DB
doc = revit.doc
t = DB.Transaction(doc,'test')
t.Start()
DB.Mechanical.MechanicalSystemType.Create(doc,DB.MEPSystemClassification.OtherAir,'31_Überstromluft_')
t.Commit()