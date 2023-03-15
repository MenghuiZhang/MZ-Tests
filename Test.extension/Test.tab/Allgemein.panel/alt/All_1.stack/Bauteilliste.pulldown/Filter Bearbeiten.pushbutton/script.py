# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *

__title__ = "Bauteilliste duplizieren für jeden Planausschnitt/Ebene(MFC)"
__doc__ = """Fliter Bearbeiten
Filter: 
IGF_Trassenzugehörihkeit: 1.UG, EG ,1.OG ,2.OG ,3.OG
IGF_X_Bildausschnitt: A1, A2, B1, B2, C1, C2, V1"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
exportfolder = script.get_config()

uidoc = revit.uidoc
doc = revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)

bauteilliste_coll = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Schedules) \
    .WhereElementIsNotElementType()
bauteillisteids = bauteilliste_coll.ToElementIds()
bauteilliste_coll.Dispose()

dict_bauteilliste = {}
liste_bauteilliste = []

for elid in bauteillisteids:
    elem = doc.GetElement(elid)
    name = elem.Name
    bb = ''
    try:
        bb = elem.LookupParameter('Bearbeitungsbereich').AsValueString()
    except:
        continue
    if bb.find(name) != -1:
        liste_bauteilliste.append(name)
        dict_bauteilliste[name] = elem

liste_bauteilliste.sort()
Dict_Ebene = {'090':'1.UG','100':'EG','110':'1.OG','120':'2.OG','130':'3.OG'}
auswahl = forms.SelectFromList.show(liste_bauteilliste, multiselect=True, button_name='Select Category')
if any(auswahl):
    t = DB.Transaction(doc)
    t.Start('Bauteilliste')
    for el in auswahl:
        elem = dict_bauteilliste[el]
        defi = elem.Definition
        fs = defi.GetFilters()

        for f in fs:
            fieldName = defi.GetField(f.FieldId).GetName()
            if fieldName == 'CAx Typ':
                para1 = f.GetStringValue()
                break
        
        for Ebene in ['090','100','110','120','130']:
            for Bereich in ['A1','A2','B1','B2','C1','C2','V1']:
                name = 'LP5_Liste ' + para1 + ' - ' + Ebene + '-' + Bereich
                if not name in liste_bauteilliste:
                    liste_bauteilliste.append(name)
                    elem_neu_id = elem.Duplicate(DB.ViewDuplicateOption.Duplicate)
                    elem_neu = doc.GetElement(elem_neu_id)
                    elem_neu.Name = name
                    defi2 = elem_neu.Definition

                    fs2 = defi2.GetFilters()
                    for f2 in fs2:
                        fieldName2 = defi2.GetField(f2.FieldId).GetName()
                        if fieldName2 == 'IGF_Trassenzugehörigkeit':
                            f2.SetValue(Dict_Ebene[Ebene])
                            defi2.SetFilter(1, f2)
                        elif fieldName2 == 'IGF_X_Bildausschnitt':
                            f2.SetValue(Bereich)
                            defi2.SetFilter(2, f2)
    t.Commit()