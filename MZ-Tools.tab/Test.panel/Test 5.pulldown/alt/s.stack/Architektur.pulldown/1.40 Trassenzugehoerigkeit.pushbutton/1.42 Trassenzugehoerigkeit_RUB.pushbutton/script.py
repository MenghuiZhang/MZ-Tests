# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import clr
import time

start = time.time()


__title__ = "1.42 Trassenzugehörigkeit_RUB"
__doc__ = """IGF_Trassenzugehörigkeit für Projekt RUB"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()


uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

HLS_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
    .WhereElementIsNotElementType()
HLS = HLS_collector.ToElementIds()

Rohre_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_PipeCurves)\
    .WhereElementIsNotElementType()
Rohre = Rohre_collector.ToElementIds()

Luft_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctCurves)\
    .WhereElementIsNotElementType()
Luft = Luft_collector.ToElementIds()

Kabel_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_CableTray)\
    .WhereElementIsNotElementType()
Kabel = Kabel_collector.ToElementIds()

Leer_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Conduit)\
    .WhereElementIsNotElementType()
Leer = Leer_collector.ToElementIds()

def DatenSchreiben(collector,Kennzeichen,ids):
    title = '{value}/{max_value} Trassenzugehörigkeit-' + Kennzeichen
    step = int(len(ids)/200)
    if step == 0:
        step = 1
    with forms.ProgressBar(title=title, cancellable=True, step=step) as pb:
        n = 0

        for el in collector:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(ids))
            Ebene = el.LookupParameter('CAx_Trassenbezugsebene').AsString()
            if Ebene:
                if len(Ebene) > 8:
                    new_ebene = Ebene[:9]
                    if new_ebene[8] == ' ':
                        new_ebene = Ebene[:8]
                    if new_ebene == 'Bodenplat':
                        new_ebene = 'NN'
                    elif new_ebene == 'Dachaufsi':
                        new_ebene = 'Dach'
                else:
                    new_ebene = Ebene
                el.LookupParameter('IGF_Trassenzugehörigkeit').Set(new_ebene)

t = Transaction(doc,'Trassenzugehörigkeit')
t.Start()
DatenSchreiben(HLS_collector,'HLS_Bauteile',HLS)
DatenSchreiben(Rohre_collector,'Rohre',Rohre)
DatenSchreiben(Luft_collector,'Luftkanäle',Luft)
DatenSchreiben(Kabel_collector,'Kabeltrassen',Kabel)
DatenSchreiben(Leer_collector,'Leerrohr',Leer)
t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
