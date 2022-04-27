# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr

start = time.time()


__title__ = "Deckensegel"
__doc__ = """Temperatur und Volumenstrom der Deckensegel in MEP Raum schreiben"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

bauteile_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
bauteile = bauteile_collector.ToElementIds()

phase = list(doc.Phases)[-1]

logger.info("{} HLS Bauteile ausgewählt".format(len(bauteile)))

if not bauteile:
    logger.error("Keine HLS Bauteile in aktueller Projekt gefunden")
    script.exit()


class HLS_Bauteile:
    Raum_Liste = []

    def __init__(self, element_id):
        """
        Definiert Deckensegel Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr_Bauteile = [
            ['Volumen', 'Volumenstrom'],
            ['KZ','Systemklassifizierung'],
            ['V_H','Vorlauftemperatur_Heizen'],
            ['R_H','Rücklauftemperatur_Heizen'],
            ['V_K','Vorlauftemperatur_Kühlen'],
            ['R_K','Rücklauftemperatur_Kühlen'],
        ]

        attr_Raum = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
        ]

        for a in attr_Bauteile:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr_BT(revit_name))

        if self.KZ == 'Vorlauf,Rücklauf' and self.V_H:
            for b in attr_Raum:
                python_name, revit_name = b
                setattr(self, python_name, self.__get_element_attr_Raum(revit_name))

            if self.nummer and self.name:
                self.Raum_Liste.append([self.nummer,self.name,self.Volumen,self.V_H,self.R_H, self.V_K, self.R_K])



        logger.info(50 * "=")
        logger.info("{}\t{}".format(self.element, self.element_id))


    def __get_element_attr_BT(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return self.__get_value_in_project_units(param)

    def __get_element_attr_Raum(self, param_name):
        if self.element.Space[phase]:
            param = self.element.Space[phase].LookupParameter(param_name)

        else:
            param = None

        if not param:
            logger.warning(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return self.__get_value_in_project_units(param)


    def __get_value_in_project_units(self, param):
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


    def __repr__(self):
        return "HLS_Bauteile({})".format(self.element_id)



HLS_Bauteile_Liste = []

with forms.ProgressBar(title='{value}/{max_value} Temperaturermittlung',
                       cancellable=True, step=10) as pb:

    for n, bauteile_id in enumerate(bauteile):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(bauteile))
        BT = HLS_Bauteile(bauteile_id)

        HLS_Bauteile_Liste.append(BT)

table_data = HLS_Bauteile.Raum_Liste
table_data.sort()

table_Data = [table_data[0]]
for i in range(1,len(table_data)):
    j = i - 1
    if str(table_data[i][0]) != str(table_data[j][0]):
        table_Data.append(table_data[i])


# Volemen con MEP Raum rechnen
for data in table_data:
    for room in table_Data:
        if data[0] == room[0]:
            room[2] = room[2] + data[2]


for item in table_Data:
    item[2] = item[2] - 0.01
output.print_table(
    table_data=table_Data,
    title="Raum",
    columns=['Nummer','Name','Volumen','Vor_Heizen','Rück_Heizen','Vor_Kühlen','Rück_Kühlen']
)


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

# MEP Räume

spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

# Schächte von Lüftung übernehmen + Abfrage
if forms.alert('Temperatur und Volemenstrom übernehmen?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Temperatur und Volumenstrom schreiben',
                           cancellable=True, step=10) as pb1:
        n_1 = 0

        t = Transaction(doc, "Temperatur übernehmen")
        t.Start()

        for Spaces in spaces_collector:

            if pb1.cancelled:
                script.exit()

            n_1 += 1

            pb1.update_progress(n_1, len(spaces))

            Nummer = get_value(Spaces.LookupParameter("Nummer"))
            Name = get_value(Spaces.LookupParameter('Name'))
            for Date in table_Data:
                if Nummer == Date[0]:
                    Spaces.LookupParameter("IGF_H_DeS-VL_Win_Temp").SetValueString(str(Date[3]))
                    Spaces.LookupParameter("IGF_H_DeS-RL_Win_Temp").SetValueString(str(Date[4]))
                    Spaces.LookupParameter("IGF_K_DeS-VL_Som_Temp").SetValueString(str(Date[5]))
                    Spaces.LookupParameter("IGF_K_DeS-RL_Som_Temp").SetValueString(str(Date[6]))
                    Spaces.LookupParameter('IGF_H_DeS_Winter').SetValueString(str(Date[2]))
                    Spaces.LookupParameter('IGF_K_DeS_Sommer').SetValueString(str(Date[2]))

        t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
