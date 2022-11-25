# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit,DB
from pyrevit import forms,script

__title__ = "4.30 Dachfläche übernehmen (Regenwasser)"
__doc__ = """
werte von Parameter 'Fläche' in 'IGF_S_Dachfläche' übernehmen, nur wenn IGF_S_RW_istDachfläche angehakt ist.
Kategorien: Detailelemente

input Parameter:
IGF_S_RW_istDachfläche: Ja/Nein

output Parameter:
IGF_S_Dachfläche: Dachfläche 

[2021.11.30]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

doc = revit.doc
logger = script.get_logger()
coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType()
elementids = coll.ToElementIds()
coll.Dispose()
        

class Flaeche:
    def __init__(self,elementid):
        self.elem_id = elementid
        self.elem = doc.GetElement(self.elem_id)
        self.Area = self.elem.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble()*0.3048*0.3048
        try:
            self.isDachFlaeche = self.elem.LookupParameter('IGF_S_RW_istDachfläche').AsInteger()
        except:
            logger.error('Parameter IGF_S_RW_istDachfläche nicht gefunden')
            self.isDachFlaeche = False


    def werte_schreiben(self):
        param = self.elem.LookupParameter('IGF_S_Dachfläche')
        if param:
            try:
                param.Set(round(float(self.Area),2))
            except:
                pass
        else:
            logger.error('Parameter IGF_S_Dachfläche nicht gefunden')



Liste = []
with forms.ProgressBar(title="{value}/{max_value} Detailelemente in Projekt", cancellable=True, step=10) as pb:
    for n, elemid in enumerate(elementids):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(elementids))
        flaeche = Flaeche(elemid)
        if flaeche.isDachFlaeche and flaeche.isDachFlaeche != 0: 
            Liste.append(flaeche)
if len(Liste) == 0:
    logger.error('Keine Dachfläche in Projekt gefunden')
    script.exit()

# Werte zuückschreiben + Abfrage
if forms.alert('Dachfläche in Parameter "IGF_S_Dachfläche" schreiben?', ok=False, yes=True, no=True):
    t = DB.Transaction(doc,'Fläche schreiben')
    t.Start()
    with forms.ProgressBar(title="{value}/{max_value} Dachfläche in Projekt", cancellable=True, step=10) as pb1:
        for n1, flaeche in enumerate(Liste):
            if pb1.cancelled:
                t.Rollback()
                script.exit()
            pb.update_progress(n1 + 1, len(Liste))
            flaeche.werte_schreiben()

    t.Commit()