# coding: utf8
from rpw import revit, DB, UI
import Autodesk.Revit.Creation as create

doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveView
commandid = UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.Move)
de = doc.GetElement(uidoc.Selection.GetElementIds()[0])
# a = uidoc.Selection.PickObject(UI.Selection.ObjectType.Edge)
b = uidoc.Selection.PickObject(UI.Selection.ObjectType.Edge)
# print(a)

# referenceArray = DB.ReferenceArray()
# referenceArray.Append(a)
# referenceArray.Append(b)

# c = uidoc.Selection.PickPoint()
# d = DB.XYZ(c.X,c.Y+1,c.Z)
# # d = uidoc.Selection.PickPoint()
# line = DB.Line.CreateBound(c,d)

t = DB.Transaction(doc,'Test')
t.Start()
# de = doc.Create.NewDimension(view,line,referenceArray)
sel = uidoc.Selection.GetElementIds()
sel.Clear()
sel.Add(b.ElementId)
uidoc.Selection.SetElementIds(sel)
print(de)
de.ValueOverride = '150'
#create.Document.NewDimension(None,view,line,referenceArray)
t.Commit()
# [(0.969720227, -0.184511651, 0.159994162)]>
# [(0.963308056, 0.163811282, 0.212611039)]>
# uidoc.ActiveView.ViewDirection
