# coding: utf8
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
from IGF_Klasse._mepraum import ReduzierFaktor

def get_AnlagenInSchacht_Luft(doc,schacht):
    outlin = DB.Outline(schacht.get_BoundingBox(None).Min,schacht.get_BoundingBox(None).Max)
    fil = DB.BoundingBoxIntersectsFilter(outlin,2)
    coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().WherePasses(fil).ContainedInDesignOption(DB.ElementId(-1)).ToElementIds()
    Liste = []
    for elid in coll:
        system = doc.GetElement(elid).LookupParameter('Systemtyp').AsElementId()
        if system.IntegerValue != -1:
            param = doc.GetElement(system).LookupParameter('IGF_X_AnlagenNr')
            if param:
                anlnr = param.AsValueString()
                if anlnr not in Liste:Liste.append(anlnr)
    Items = ObservableCollection[ReduzierFaktor]()
    for anlnr in sorted(Liste):
        Items.Add(ReduzierFaktor(anlnr))
    outlin.Dispose()
    fil.Dispose()
    del(Liste)
    del(coll)
    return Items
