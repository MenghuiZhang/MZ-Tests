# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr
import csv

start = time.time()


__title__ = "Test"
__doc__ = """Temperatur und Leistung von Umluftkühler in MEP Raum schreiben"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

bauteile_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
bauteile = bauteile_collector.ToElementIds()


logger.info("{} HLS Bauteile ausgewählt".format(len(bauteile)))

if not bauteile:
    logger.error("Keine HLS Bauteile in aktueller Projekt gefunden")
    script.exit()


def get_value(param):
    """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""

    value = revit.query.get_param_value(param)

    try:
        unit = param.DisplayUnitType

        value = DB.UnitUtils.ConvertFromInternalUnits(
            value,
            unit)

    except Exception as e:
        pass

    return value

with forms.ProgressBar(title='{value}/{max_value} Temperatur übernehmen',
                           cancellable=True, step=10) as pb1:
        n_1 = 0

        t = Transaction(doc, "Temperatur übernehmen")
        t.Start()

        for HLS_BT in bauteile_collector:

            if pb1.cancelled:
                script.exit()

            n_1 += 1

            pb1.update_progress(n_1, len(bauteile))

            KG = HLS_BT.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()

            print(HLS_BT.Id,KG)



        t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))

# class HLS_Bauteile:
#
#     def __init__(self, element_id):
#         """
#         Definiert Umluftkühler Klasse mit allen object properties für die
#         Luftmengen Berechnung.
#         """
#
#         self.element_id = element_id
#         self.element = doc.GetElement(self.element_id)
#
#         self.KG = None
#
#
#         logger.info(50 * "=")
#         logger.info("{}\t{}".format(self.element, self.element_id))
#
#
#     def __get_element_attr_BT(self, param_name):
#         param = self.element.LookupParameter(param_name)
#
#         if not param:
#             logger.error(
#                 "Parameter {} konnte nicht gefunden werden".format(param_name))
#             return
#
#         return self.__get_value_in_project_units(param)
#
#     def __get_element_attr_Raum(self, param_name):
#         if self.element.Space[phase]:
#             param = self.element.Space[phase].LookupParameter(param_name)
#
#         else:
#             param = None
#
#         if not param:
#             logger.warning(
#                 "Parameter {} konnte nicht gefunden werden".format(param_name))
#             return
#
#         return self.__get_value_in_project_units(param)
#
#
#     def __get_value_in_project_units(self, param):
#         """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""
#
#         value = revit.query.get_param_value(param)
#
#         try:
#             unit = param.DisplayUnitType
#
#             value = DB.UnitUtils.ConvertFromInternalUnits(
#                 value,
#                 unit)
#
#         except Exception as e:
#             pass
#
#         return value
#
#
#     def __repr__(self):
#         return "HLS_Bauteile({})".format(self.element_id)


#
# HLS_Bauteile_Liste = []
#
# with forms.ProgressBar(title='{value}/{max_value} Temperaturermittlung',
#                        cancellable=True, step=10) as pb:
#
#     for n, bauteile_id in enumerate(bauteile):
#         if pb.cancelled:
#             script.exit()
#
#         pb.update_progress(n + 1, len(bauteile))
#         BT = HLS_Bauteile(bauteile_id)
#
#
#         HLS_Bauteile_Liste.append(BT)



# # Werte zuückschreiben + Abfrage
# if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
#     with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
#                            cancellable=True, step=10) as pb2:
#
#         n_1 = 0
#
#         with rpw.db.Transaction("Luftwechsel berechnen"):
#             for BLS_BT in HLS_Bauteile_Liste:
#                 if pb2.cancelled:
#                     script.exit()
#                 n_1 += 1
#                 pb2.update_progress(n_1, len(HLS_Bauteile_Liste))
#
#                 BLS_BT.werte_schreiben()

# t = Transaction(doc, "Temperatur übernehmen")
# t.Start()

# for Spaces in spaces_collector:
#
#
#     for Date in table_Data:
#         if Nummer == Date[0]:
#             Spaces.LookupParameter("IGF_K_ULK-VL_Temp").SetValueString(str(Date[5]))
#             Spaces.LookupParameter("IGF_K_ULK-RL_Temp").SetValueString(str(Date[6]))
#             Spaces.LookupParameter('IGF_K_ULK_Wassermenge').SetValueString(str(Date[3]))
#             Spaces.LookupParameter('IGF_K_ULK_Leistung').SetValueString(str(Date[4]))
#
# t.Commit()
# output.print_table(
#     table_data=HLS_Bauteile.Liste,
#     title="Luftmengen",
#     columns=['ID', 'Name', 'Wert', ]#  'Heizleistung_HK', 'Heizleistung_Gesamt',
#
# )

#
# table_Data = [[table_data[0][0],table_data[0][1],table_data[0][2],0,0,0,0]] # Vol, Leistung, V_H, R_H
# for i in range(1,len(table_data)):
#     j = i - 1
#     if str(table_data[i][0]) != str(table_data[j][0]):
#         table_Data.append([table_data[i][0],table_data[i][1],table_data[i][2],0,0,0,0,])
#
# # Volemen con MEP Raum rechnen
# for data in table_data:
#     for room in table_Data:
#         if data[0] == room[0]:
# #            room[3] = room[3] + data[2]
#             room[4] = room[4] + data[3]
#
#

#
# #Rohr System Type aus aktueller Projekt
# systemtype_collector = DB.FilteredElementCollector(doc) \
#     .OfClass(clr.GetClrType(DB.Plumbing.PipingSystemType))
#
# systemtype = systemtype_collector.ToElementIds()
#
# logger.info("{} Rohr System Type ausgewählt".format(len(systemtype)))
# if not systemtype:
#     logger.error("Keine Rohr System Type in aktueller Projekt gefunden")
#     script.exit()
#
# # Rohr System aus aktueller Projekt
# systems_collector = DB.FilteredElementCollector(doc) \
#     .OfClass(clr.GetClrType(DB.Plumbing.PipingSystem))
#
# systems = systems_collector.ToElementIds()
#
# logger.info("{} Rohr Systems ausgewählt".format(len(systems)))
#
# if not systems:
#     logger.error("Keine Rohr Systems in aktueller Projekt gefunden")
#     script.exit()

# Sys_Date = []
# for Sys in systems_collector:
#     SysName = get_value(Sys.LookupParameter('Systemname'))
#     Typ = get_value(Sys.LookupParameter('Typ'))
#     Sys_Date.append([SysName,Typ])
#
# SysT_Date = []
# for SysT in systemtype_collector:
#     ID = SysT.Id
#     Temp = get_value(SysT.LookupParameter('Temperatur von Medium'))
#     Abk = get_value(SysT.LookupParameter('Abkürzung'))
#     if Temp:
#         Temp = round(Temp,0)
#     SysT_Date.append([ID,Temp,Abk])
#
# Sys_Temp = []
# for Date in Sys_Date:
#     ID = Date[1]
#     for date in SysT_Date:
#         if ID == date[0]:
#             Sys_Temp.append([Date[0],date[1],date[2]])
#
# Sys_Temp.sort()
#
#
#
# for Daten in table_Data:
#     for DATEN in Sys_Temp:
#         for Name in Daten[2]:
#             if Name == DATEN[0] and DATEN[2] == 'K VL':
#                 Daten[5] = DATEN[1]
#             if Name == DATEN[0] and DATEN[2] == 'K RL':
#                 Daten[6] = DATEN[1]






# # MEP Räume
#
# spaces_collector = DB.FilteredElementCollector(doc) \
#     .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
# spaces = spaces_collector.ToElementIds()
#
# logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
#
# if not spaces:
#     logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
#     script.exit()

# # Schächte von Lüftung übernehmen + Abfrage
# if forms.alert('Temperatur übernehmen?', ok=False, yes=True, no=True):
#
#
