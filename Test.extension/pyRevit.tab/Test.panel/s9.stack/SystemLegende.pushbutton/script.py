# coding: utf8
from rpw import revit,DB

doc = revit.doc
activeview = revit.active_view

coll = DB.FilteredElementCollector(doc,activeview.Id)

berechnung_nach = {
    "Fläche": '1',
    "Luftwechsel": '2',
    "Person": '3',
    "manuell": '4',
    "nurZUMa": '5',
    "nurABMa": '6',
    "nurZU_Fläche": '5.1',
    "nurZU_Luftwechsel": '5.2',
    "nurZU_Person": '5.3',
    "nurAB_Fläche": '6.1',
    "nurAB_Luftwechsel": '6.2',
    "nurAB_Person": '6.3',
    "keine": '9'
}

einheit = {
    '1': 'm³/h pro m²',
    '2': '-1/h',
    '3': 'm³/h pro P',
    '4': 'm³/h ',
    '5': 'm³/h ',
    '6': 'm³/h' ,
    '5.1': "m³/h pro m²",
    '5.2': '-1/h',
    '5.3': 'm³/h pro P',
    '6.1': "m³/h pro m²",
    '6.2': '-1/h',
    '6.3': 'm³/h pro P',
    "keine": 'keine'
}

t = DB.Transaction(doc,'nummer')
t.Start()
for el in coll:
    name = el.LookupParameter('TGA_RLT_VolumenstromProName').AsString()

    if name in berechnung_nach.keys():
        el.LookupParameter('TGA_RLT_VolumenstromProNummer').SetValueString(berechnung_nach[name])
    el.LookupParameter('IGF_RLT_Nachtbetrieb').Set(1)
t.Commit()




