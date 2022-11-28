# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from IGF_forms import abfrage

__title__ = "1.01 Raumtrennlinie löschen"
__doc__ = """Raumtrennlinie löschen

[2021.11.22]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

doc = revit.doc
active_view = doc.ActiveView

frage_schicht = abfrage(title= '1.70 Raumtrennlinie löschen',
 info = 'In aktueller Ansicht oder im ganzen Projekt?' ,
 anmerkung = 'Achtung! durch das Löschen der MEP-Raumtrennlinien können Flächen-, Luftmengenberechnung oder weitere Berechnung fehlschlagen. Erst wenn das Berechnungsmodell eine gewisse Qualität hat, können diese gelöscht werden.', 
 ja = True,ja_text= 'Ansicht',nein_text='Projekt')

frage_schicht.ShowDialog()


coll = None
if frage_schicht.antwort == 'Ansicht':
    coll = DB.FilteredElementCollector(doc,active_view.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaceSeparationLines).WhereElementIsNotElementType()
elif frage_schicht.antwort == 'Projekt':
    coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaceSeparationLines).WhereElementIsNotElementType()
if not coll:
    script.exit()


frage_delete = abfrage(title= 'Achtung!', info = 'MEP-Raumtrennlinie wirklich löschen?', ja = True,ja_text= 'Ja',nein_text='Nein')
frage_delete.ShowDialog()

if coll:
    elementids = coll.ToElementIds()
    coll.Dispose()

if frage_delete.antwort == 'Ja':
    t_del = DB.Transaction(doc,'Raumtrennlinie löschen')
    t_del.Start()
    with forms.ProgressBar(title="{value}/{max_value} MEP-Raumtrennlinie", cancellable=True, step=1) as pb:
        for n, elementid in enumerate(elementids):
            if pb.cancelled:
                t_del.RollBack()
                script.exit()
            pb.update_progress(n + 1, len(elementids))

            try:
                doc.Delete(elementid)
            except:
                pass

    t_del.Commit()

elif frage_delete.antwort == 'Nein':
    pass