# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import time
import clr
import operator
from pyIGF_logInfo import getlog
start = time.time()


__title__ = "1.21 erstellt Globale Parameter für Geschosshöhe"
__doc__ = """DeckenHöhen berechnen und entsprechende Globale Parameter erstellen. z.B. '4.OG UKRD-4.OG OKRB' """
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
getlog(__title__)
uidoc = revit.uidoc
doc = revit.doc

# Ebene aus aktueller Projekt
Levels_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Levels)\
    .WhereElementIsNotElementType()
levels = Levels_collector.ToElementIds()

logger.info("{} Ebenen ausgewählt".format(len(levels)))

if not levels:
    logger.error("Keine Ebene in aktueller Projekt gefunden")
    script.exit()

Liste = []
for i in Levels_collector:
    name = i.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()
    height = i.LookupParameter('Ansicht').AsValueString()
    if 'F' in [i for i in name]:
        continue
    if 't' in [i for i in name]:
        continue
    Liste.append([name,height])

Liste2 = []
for i in range(len(Liste)):
    liste_temp = []
    for j in range(i+1,len(Liste)):
        if Liste[i][0][:3] == Liste[j][0][:3]:
            liste_temp = [Liste[i],Liste[j]]
            Liste2.append(liste_temp)



Liste3 = []
for i in Liste2:
    name_list = [i[0][0],i[1][0]]
    name_list.sort()
    Name = name_list[1] + '-' + name_list[0]
    height = abs(int(i[1][1])-int(i[0][1]))
    if height > 1500:
        Liste3.append([Name,height])
        print(Name)

# Globale Parameters aus aktureller Projekt
GlobaleParameter_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.GlobalParameter))\
    .WhereElementIsNotElementType()
GlParameter = GlobaleParameter_collector.ToElementIds()
logger.info("{} globale Parameters in Projekt".format(len(GlParameter)))

AllGPName = []
for g in GlobaleParameter_collector:
    AllGPName.append(g.Name)

# Parameters erstellen, Werte zuückschreiben + Abfrage
if forms.alert('Globale Parameter erstellen?', ok=False, yes=True, no=True):
    t = DB.Transaction(doc, "Erstellen")
    t.Start()

    i = 0
    j = 0

    with forms.ProgressBar(title='globale Parameter erstellen',
                           cancellable=True, step=1) as pb2:


        n_1 = 0
        for Ebe in Liste3:

            if pb2.cancelled:
                script.exit()
            n_1 += 1
            pb2.update_progress(n_1, len(Liste3))

            if not Ebe[0] in AllGPName:
                new_Para = DB.GlobalParameter.Create(doc,Ebe[0],DB.ParameterType.Number)
                new_Para.SetFormula(str(Ebe[1]))
                new_Para.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS
                logger.info("{},  {}".format(Ebe[0], Ebe[1]))

    t.Commit()



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
