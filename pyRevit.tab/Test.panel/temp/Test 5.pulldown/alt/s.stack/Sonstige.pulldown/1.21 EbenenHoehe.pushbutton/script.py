# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms

__title__ = "1.41 Geschosshöhe"
__doc__ = """
DeckenHöhen berechnen und entsprechende Globale Parameter erstellen. 

z.B. '4.OG UKRD-4.OG OKRB' 

Achtung: Ebenename muss 'UKRD' oder 'OKRB' enthalten.

[2021.11.25]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
uidoc = revit.uidoc
doc = revit.doc

# Ebene aus aktueller Projekt
Levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
levels = Levels_collector.ToElementIds()

if not levels:
    logger.error("Keine Ebene in aktueller Projekt gefunden")
    script.exit()

dict_UKRD = {}
dict_OKRB = {}
for level in Levels_collector:
    name = level.Name
    if name.upper().find('UKRD') != -1:
        name = name.replace('UKRD','')
        dict_UKRD[name] = level
    elif name.upper().find('OKRB') != -1:
        name = name.replace('OKRB','')
        dict_OKRB[name] = level

liste_final = []
for ukrd in dict_UKRD.keys():
    for okrb in dict_OKRB.keys():
        if ukrd.find(okrb) != -1 or okrb.find(ukrd) != -1:
            liste_final.append([dict_UKRD[ukrd],dict_OKRB[okrb]])

class Ebene:
    def __init__(self,ukrd,okrb):
        self.ukrd = ukrd
        self.okrb = okrb
        self.ukrdname = self.ukrd.Name
        self.okrbname = self.okrb.Name
        self.ukrdhoehe = get_value(self.ukrd.LookupParameter('Ansicht'))
        self.okrbhoehe = get_value(self.okrb.LookupParameter('Ansicht'))
        self.deckenhoehe = self.ukrdhoehe - self.okrbhoehe
        self.gp_name = self.ukrdname + '-' + self.okrbname
    def create_uzuo(self):
        gp = doc.GetElement(DB.GlobalParametersManager.FindByName(doc,self.gp_name))
        if not gp:
            gp = DB.GlobalParameter.Create(doc,self.gp_name,DB.ParameterType.Number)
        gp.SetFormula(str(int(round(self.deckenhoehe))))
        gp.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS

if len(liste_final) > 1:   
    # Parameters erstellen, Werte zurückschreiben + Abfrage
    if forms.alert('Globale Parameter für Deckenhöhe erstellen?', ok=False, yes=True, no=True):
        t = DB.Transaction(doc, "Deckenhöhe")
        t.Start()
        with forms.ProgressBar(title='{value}/{max_value} globale Parameter',cancellable=True, step=1) as pb:
            for n, Ebe in enumerate(liste_final):
                if pb.cancelled:
                    t.RollBack()
                    script.exit()
                pb.update_progress(n + 1, len(liste_final))
                ebene = Ebene(Ebe[0], Ebe[1])
                if ebene.deckenhoehe < 0:
                    if forms.alert('Globale Parameter {} erstellen? Achtung: Deckenhöhe ist kleiner als 0!!!', ok=False, yes=True, no=True):
                        ebene.create_uzuo()
                else:
                    ebene.create_uzuo()
        t.Commit()
        t.Dispose()