# coding: utf8
import time
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms


start = time.time()
datum = time.strftime("%d.%m.%Y", time.localtime())

__title__ = "1.30 Architektur-Raumzustandsänderung"
__doc__ = """schreibt Raumzustandsdaten von Architektur Modell in Zustandsparameter des Berechnungsmodells.
Bezug auf Volumen, Umfang, Licht Höhe, Fläche.

Kategorie: Räume

Parameter:
IGF_A_Datum_Architektur, 
IGF_A_Fläche_Architektur, 
IGF_A_Umfang_Architektur, 
IGF_A_Höhe_Architektur, 
IGF_A_Volumen_Architektur,
IGF_A_Nummer_Architektur

[2021.11.25]
Version: 1.1
"""

__author__ = "Menghui Zhang"

logger = script.get_logger()
uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass


spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

if not spaces:
    logger.error("Keine Räume in aktueller Projekt gefunden")
    script.exit()

revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el

revitLinks_collector.Dispose()


rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()
rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()

RaumRVTLink = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
RaumRVTLinkIDs = RaumRVTLink.ToElementIds()
RaumRVTLink.Dispose()

class Raum:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem_doc = doc.GetElement(self.elemid)
        self.elem_rvtl = rvtdoc.GetElement(self.elemid)
        self.Name = get_value(self.elem_doc.LookupParameter('Name'))
        self.nummer = self.elem_doc.Number

        if not self.elem_rvtl:
            logger.error('Raum {}: {} nur in Berechnungsmodell.'.format(self.nummer,self.Name))
        else:
            self.Nummer_rvtl = get_value(self.elem_rvtl.LookupParameter('Nummer'))
            self.Flaeche_rvtl = get_value(self.elem_rvtl.LookupParameter('Fläche'))
            self.Umfang_rvtl = get_value(self.elem_rvtl.LookupParameter('Umfang'))
            self.Hoehe_rvtl = get_value(self.elem_rvtl.LookupParameter('Lichte Höhe'))
            self.Volumen_rvtl = get_value(self.elem_rvtl.LookupParameter('Volumen'))

    def wert_schreiben(self,para,wert):
        param = self.elem_doc.LookupParameter(para)
        if param:
            if param.StorageType.ToString() == 'Double':
                param.SetValueString(str(wert))
            else:
                param.Set(wert)
    
    def werte_schreiben(self):
        self.wert_schreiben('IGF_A_Nummer_Architektur',self.Nummer_rvtl)
        self.wert_schreiben('IGF_A_Fläche_Architektur',self.Flaeche_rvtl)
        self.wert_schreiben('IGF_A_Umfang_Architektur',self.Umfang_rvtl)
        self.wert_schreiben('IGF_A_Höhe_Architektur',self.Hoehe_rvtl)
        self.wert_schreiben('IGF_A_Volumen_Architektur',self.Volumen_rvtl)
        self.wert_schreiben('IGF_A_Datum_Architektur',datum)


raum_liste = []
with forms.ProgressBar(title='{value}/{max_value} Räume in Berechnungsmodell', cancellable=True, step=10) as pb:
    for n, spaceid in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        raum = Raum(spaceid)
        if raum.elem_rvtl:
            raum_liste.append(raum)

if forms.alert("Raumzustand in Berechnungsmodell schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Räume",cancellable=True, step=10) as pb1:
        t = DB.Transaction(doc)
        t.Start('Raumzustand')
        for n,raum in enumerate(raum_liste):
            if pb1.cancelled:
                t.RollBack()
                script.exit()
            pb1.update_progress(n+1, len(raum_liste))
            raum.werte_schreiben()

        t.Commit()
        t.Dispose()


rvtl_ids = [i.ToString() for i in RaumRVTLinkIDs]
docraum_ids = [i.ToString() for i in spaces]
for item in rvtl_ids:
    if item not in docraum_ids:
        elem = rvtdoc.GetElement(DB.ElementId(int(item)))
        nummer = get_value(elem.LookupParameter('Nummer'))
        name = get_value(elem.LookupParameter('Name'))
        logger.error('Raum {}: {} nur in Revitverknüpfungsmodell.'.format(nummer,name))