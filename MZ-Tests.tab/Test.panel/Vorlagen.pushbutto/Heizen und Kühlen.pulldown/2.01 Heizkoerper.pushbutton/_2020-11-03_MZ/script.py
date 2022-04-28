# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr

start = time.time()


__title__ = "Heizkörper"
__doc__ = """Temperatur und Volumenstrom bzw. Heizleistung von Heizkörper in MEP Raum schreiben"""
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
        Definiert Heizkörper Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr_Bauteile = [
            ['Volumen', 'Massenstrom (R)'],
            ['Name', 'Systemname'],
            ['KZ','Systemklassifizierung'],
            ['Temp','Vorlauftemperatur_Heizen'],
            ['HL','Heizleistung'],
        ]

        attr_Raum = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
        ]

        for a in attr_Bauteile:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr_BT(revit_name))

        if self.KZ == 'Vorlauf,Rücklauf' or self.KZ == 'Rücklauf,Vorlauf':
            if self.Temp == None and self.element.Space[phase]:
                for b in attr_Raum:
                    python_name, revit_name = b
                    setattr(self, python_name, self.__get_element_attr_Raum(revit_name))
                self.BT_Name = self.Name.split(',')
                self.Raum_Liste.append([self.nummer,self.name,self.BT_Name,self.Volumen,self.HL])


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

table_Data = [[table_data[0][0],table_data[0][1],table_data[0][2],0,0,0,0]] # Vol_Som,HL, V_H, R_H
for i in range(1,len(table_data)):
    j = i - 1
    if str(table_data[i][0]) != str(table_data[j][0]):
        table_Data.append([table_data[i][0],table_data[i][1],table_data[i][2],0,0,0,0,])

output.print_table(
    table_data=table_data,
    title="Raum",
    columns=['Nummer','Name','Sys_Namen','Vol', 'Heizleistung',]
)


# Volemen con MEP Raum rechnen
for data in table_data:
    for room in table_Data:
        if data[0] == room[0]:
            if data[3]:
                room[3] = room[3] + data[3]
            if data[4]:
                room[4] = room[4] + data[4]


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

#Rohr System Type aus aktueller Projekt
systemtype_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.Plumbing.PipingSystemType))

systemtype = systemtype_collector.ToElementIds()

logger.info("{} Rohr Systems ausgewählt".format(len(systemtype)))
if not systemtype:
    logger.error("Keine Rohr System Type in aktueller Projekt gefunden")
    script.exit()

# Rohr System aus aktueller Projekt
systems_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.Plumbing.PipingSystem))

systems = systems_collector.ToElementIds()

logger.info("{} Rohr Systems ausgewählt".format(len(systems)))

if not systems:
    logger.error("Keine Rohr Systems in aktueller Projekt gefunden")
    script.exit()

Sys_Date = []
for Sys in systems_collector:
    SysName = get_value(Sys.LookupParameter('Systemname'))
    Typ = get_value(Sys.LookupParameter('Typ'))
    Sys_Date.append([SysName,Typ])

SysT_Date = []
for SysT in systemtype_collector:
    ID = SysT.Id
    Temp = get_value(SysT.LookupParameter('Temperatur von Medium'))
    Abk = get_value(SysT.LookupParameter('Abkürzung'))
    if Temp:
        Temp = round(Temp,0)
    SysT_Date.append([ID,Temp,Abk])

Sys_Temp = []
for Date in Sys_Date:
    ID = Date[1]
    for date in SysT_Date:
        if ID == date[0]:
            Sys_Temp.append([Date[0],date[1],date[2]])

Sys_Temp.sort()



for Daten in table_Data:
    for DATEN in Sys_Temp:
        for Name in Daten[2]:
            if Name == DATEN[0] and DATEN[2] == 'H VL':
                Daten[5] = DATEN[1]
            if Name == DATEN[0] and DATEN[2] == 'H RL':
                Daten[6] = DATEN[1]


output.print_table(
    table_data=table_Data,
    title="Raum",
    columns=['Nummer','Name','Sys_Namen','Vol', 'Heizleistung', 'Vor_Temp','Rück_Temp',]
)

# MEP Räume

spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

# Schächte von Lüftung übernehmen + Abfrage
if forms.alert('Temperatur übernehmen?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Temperatur übernehmen',
                           cancellable=True, step=10) as pb1:
        n_1 = 0

        t = Transaction(doc, "Temperatur übernehmen")
        t.Start()

        for Spaces in spaces_collector:

            if pb1.cancelled:
                script.exit()

            n_1 += 1

            pb1.update_progress(n_1, len(spaces))
#
            Nummer = get_value(Spaces.LookupParameter("Nummer"))
            Name = get_value(Spaces.LookupParameter('Name'))
            for Date in table_Data:
                if Nummer == Date[0]:
                    Spaces.LookupParameter("IGF_H_HK-VL_Temp").SetValueString(str(Date[5]))
                    Spaces.LookupParameter("IGF_H_HK-RL_Temp").SetValueString(str(Date[6]))
                    Spaces.LookupParameter('IGF_H_HK_Wassermenge').SetValueString(str(Date[3]))
                    Spaces.LookupParameter('IGF_H_HK_Leistung').SetValueString(str(Date[4]))

        t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
