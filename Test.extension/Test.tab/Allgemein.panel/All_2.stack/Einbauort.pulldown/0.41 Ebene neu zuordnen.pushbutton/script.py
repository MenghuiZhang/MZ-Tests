# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit import script, forms
from rpw import revit,DB


__title__ = "0.41 Ebene neu zuordnen(nicht fertig)"
__doc__ = """Einbauort von HLS-Bauteile, Sanitärinstallationen, Luftauslässe, Luftkanalzubehör und Rohrzubehör"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

try:
    getlog(__title__)
except:
    pass


uidoc = revit.uidoc
doc = revit.doc
active = doc.ActiveView
# HLS-Bauteile
bauteile_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
    .WhereElementIsNotElementType()
bauteile = bauteile_collector.ToElementIds()

logger.info("{} HLS Bauteile ausgewählt".format(len(bauteile)))
#Sanitärinstallationen
sanitaer_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_PlumbingFixtures)\
    .WhereElementIsNotElementType()
sanitaer = sanitaer_collector.ToElementIds()

#GA-Installationen
GA_collector = DB.FilteredElementCollector(doc,active.Id) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)\
    .WhereElementIsNotElementType()
GA = GA_collector.ToElementIds()

logger.info("{} GA-Installationen ausgewählt".format(len(GA)))
# Luftauslässe
Luftauslaesse_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctTerminal)\
    .WhereElementIsNotElementType()
Luftauslaesse = Luftauslaesse_collector.ToElementIds()

logger.info("{} Luftauslässe ausgewählt".format(len(Luftauslaesse)))
# Luftkanalzubehör
luftkanalzubehoer_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctAccessory)\
    .WhereElementIsNotElementType()
luftkanalzubehoer = luftkanalzubehoer_collector.ToElementIds()

logger.info("{} Luftkanalzubehör ausgewählt".format(len(luftkanalzubehoer)))
# Rohrzubehör
rohrzubehoer_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_PipeAccessory)\
    .WhereElementIsNotElementType()
rohrzubehoer = rohrzubehoer_collector.ToElementIds()

logger.info("{} Rohrzubehör ausgewählt".format(len(rohrzubehoer)))

phase = doc.Phases[0]
# Werte schreiben
def Daten_Schreiben(Familie,Familie_name,Familie_Id):
    Title = '{value}/{max_value} Einbauort--' +Familie_name +' schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=10) as pb:
        n = 0

        t = DB.Transaction(doc, "Einbauort-"+Familie_name)
        t.Start()

        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))

            Raumnummer = None
            if Item.Space[phase]:
                Familie_Name = Item.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                Familie_Typ = Item.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
                if not Familie_Name in ['_K_IGF_435_Deckensegel','_K_IGF_435_Deckensegelverbinder','_K_IGF_435_Deckensegelverbinder_Berechnung']:
                    continue
                Raumnummer = Item.Space[phase].Number


                logger.info('Familie: {}, Typ: {}, Raumnummer {}'.format(Familie_Name,Familie_Typ,Raumnummer))
                logger.info(30*'-')
                Item.LookupParameter("IGF_X_Einbauort").Set(str(Raumnummer))

        t.Commit()


if forms.alert('Einbauort schreiben?', ok=False, yes=True, no=True):
    Daten_Schreiben(bauteile_collector,'HLS-Bauteile',bauteile)
    # Daten_Schreiben(GA_collector, 'GA-Installationen', GA)
    # Daten_Schreiben(sanitaer_collector,'Sanitärinstallationen',sanitaer)
    # Daten_Schreiben(Luftauslaesse_collector,'Luftauslässe',Luftauslaesse)
    # Daten_Schreiben(luftkanalzubehoer_collector,'Luftkanalzubehör',luftkanalzubehoer)
    # Daten_Schreiben(rohrzubehoer_collector,'Rohrzubehör',rohrzubehoer)