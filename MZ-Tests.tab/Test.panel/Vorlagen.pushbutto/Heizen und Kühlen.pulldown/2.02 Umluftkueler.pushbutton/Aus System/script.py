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
            ['Name', 'Systemname'],
            ['KZ','Systemklassifizierung'],
            ['Temp','Vorlauftemperatur_Heizen'],
        ]

        attr_Raum = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
        ]

        for a in attr_Bauteile:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr_BT(revit_name))

        if self.Temp:
            for b in attr_Raum:
                python_name, revit_name = b
                setattr(self, python_name, self.__get_element_attr_Raum(revit_name))


        logger.info(50 * "=")
        logger.info("{}\t{}".format(self.element, self.element_id))

        if self.Name and self.KZ == 'Vorlauf,Rücklauf' and self.Temp and self.name:
            self.BT_Name = self.Name.split(',')
            self.Raum_Liste.append([self.nummer,self.name,self.Volumen,self.BT_Name,])

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

table_Data = [[table_data[0][0],table_data[0][1],table_data[0][3],0,0,0,0,0,0]] # Vol_Som, Vol_Win, Vor_K, R_K, V_H, R_H
for i in range(1,len(table_data)):
    j = i - 1
    if str(table_data[i][0]) != str(table_data[j][0]):
        table_Data.append([table_data[i][0],table_data[i][1],table_data[i][3],0,0,0,0,0,0])

# Volemen con MEP Raum rechnen
for data in table_data:
    for room in table_Data:
        if data[0] == room[0]:
            room[3] = room[3] + data[2]
            room[4] = room[4] + data[2]


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
            if Name == DATEN[0] and DATEN[2] == 'K VL DS':
                Daten[5] = DATEN[1]
            if Name == DATEN[0] and DATEN[2] == 'K RL DS':
                Daten[6] = DATEN[1]


output.print_table(
    table_data=table_Data,
    title="Raum",
    columns=['Nummer','Name','Sys_Namen','Vol_Win','Vol_Som',
             'Vor_Kühlen','Rück_Kühlen','Vor_Heizen','Rück_Heizen']
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

            Nummer = get_value(Spaces.LookupParameter("Nummer"))
            Name = get_value(Spaces.LookupParameter('Name'))
            for Date in table_Data:
                if Nummer == Date[0]:
#                    Spaces.LookupParameter("IGF_H_DeS-VL_Win_Temp").SetValueString(str(Date[2]))
#                    Spaces.LookupParameter("IGF_H_DeS-RL_Win_Temp").SetValueString(str(Date[3]))
                    Spaces.LookupParameter("IGF_K_DeS-VL_Som_Temp").SetValueString(str(Date[5]))
                    Spaces.LookupParameter("IGF_K_DeS-RL_Som_Temp").SetValueString(str(Date[6]))
                    Spaces.LookupParameter('IGF_H_DeS_Winter').SetValueString(str(Date[3]))
                    Spaces.LookupParameter('IGF_K_DeS_Sommer').SetValueString(str(Date[4]))

        t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
#output.print_table(
#    table_data=Sys_Date,
#    title="System",
#    columns=['SysName', 'Typ', ]
#)

#output.print_table(
#    table_data=SysT_Date,
#    title="SystemType",
#    columns=['SysTName', 'Temp', ]
#)

#output.print_table(
#    table_data=Sys_Temp,
#    title="System",
#    columns=['SysName', 'Temp', ]
#)
# HLS Buateile aus aktueller Projekt




# Schachtsname von Zu- und Abluft jedes MEP Raum lesen
#tabel_Date_MEP = []
#schacht = []
#with forms.ProgressBar(title='{value}/{max_value} Raumliste',
#                       cancellable=True, step=10) as pb:
#
#    n = 0
#    for Space in spaces_collector:
#
#        n += 1
#
#        if pb.cancelled:
#            script.exit()

#        pb.update_progress(n, len(spaces))
#
#        nummer = get_value(Space.LookupParameter("Nummer"))
#        name = get_value(Space.LookupParameter("Name"))
#        ins_schacht = get_value(Space.LookupParameter("TGA_RLT_InstallationsSchacht"))
#        install_schacht_name = get_value(Space.LookupParameter("TGA_RLT_InstallationsSchachtName"))
##        luft_zu = get_value(Space.LookupParameter("TGA_RLT_SchachtZuluft"))
#        luft_ab = get_value(Space.LookupParameter("TGA_RLT_SchachtAbluft"))
#        vewasser_vl = get_value(Space.LookupParameter("IGF_S_Schacht_VEWasserVL"))
#        vewasser_rl = get_value(Space.LookupParameter("IGF_S_Schacht_VEWasserRL"))
#        tabel_Date_MEP.append([nummer,name,luft_zu,luft_ab,vewasser_vl,vewasser_rl])
#        if ins_schacht:
#            schacht.append([nummer,name,install_schacht_name,luft_zu,luft_ab,vewasser_vl,vewasser_rl])
#
#
#        logger.info(30 * "=")
#        logger.info("{} {}".format(nummer, name))
#
## Schachtliste, Schachtliste von Zu- und Abluft
#
#S_0 = []
#S_1 = []
#S_2 = []
#for Sch in schacht:
#    S_0.append(Sch[2])
#for MEP in tabel_Date_MEP:
#    S_1.append(MEP[2])
#    S_2.append(MEP[3])
#s_0 = set(S_0)
#S0 = list(s_0)
#s_1 = set(S_1)
#S1 = list(s_1)
#s_2 = set(S_2)
#S2 = list(s_2)
#S0.sort()
#S1.sort()
#S2.sort()
#logger.info("Schachtlist: {}".format(S0))
#logger.info("Schachtlist von Zuluft: {}".format(S1))
#logger.info("Schachtlist von Abluft: {}".format(S2))



#output.print_table(
#    table_data=tabel_Date_MEP,
#    title="Mep Räume",
#    columns=['nummer', 'name',  'Zuluft_Schacht','Abluft_Schacht',
#             'Schacht_Vorlauf','Schacht_Rücklauf']
#)

#output.print_table(
#    table_data=schacht,
#    title="Schächte",
#    columns=['nummer', 'name', 'Install_Name', 'Zuluft_Schacht','Abluft_Schacht',
#            'Schacht_Vorlauf','Schacht_Rücklauf']
#)

# Schächte von Lüftung übernehmen + Abfrage
#if forms.alert('Schächte von Lüftung übernehmen?', ok=False, yes=True, no=True):

#    with forms.ProgressBar(title='{value}/{max_value} Schächte übernehmen',
#                           cancellable=True, step=10) as pb1:
#        n_1 = 0
#
#        t = Transaction(doc, "Schächte übernehmen")
#        t.Start()
#
#        for Spaces in spaces_collector:
#
#            if pb1.cancelled:
#                script.exit()
#
#            n_1 += 1
#
#            pb1.update_progress(n_1, len(spaces))
#
#            Nummer = get_value(Spaces.LookupParameter("Nummer"))
#            Name = get_value(Spaces.LookupParameter("Name"))
#            Zuluft_Schacht = get_value(Spaces.LookupParameter("TGA_RLT_SchachtZuluft"))
#            Abluft_Schacht = get_value(Spaces.LookupParameter("TGA_RLT_SchachtAbluft"))
#            VEW_VL_Schacht = Spaces.LookupParameter("IGF_S_Schacht_VEWasserVL")
#            VEW_RL_Schacht = Spaces.LookupParameter("IGF_S_Schacht_VEWasserRL")
#            if Zuluft_Schacht and Zuluft_Schacht != 'None':
#                VEW_VL_Schacht.Set(str(Zuluft_Schacht))
#            if Abluft_Schacht and Abluft_Schacht != 'None':
#                VEW_RL_Schacht.Set(str(Abluft_Schacht))
#
#
#        t.Commit()
#


#system_liste = []
#table_data = []
#with forms.ProgressBar(title='{value}/{max_value} Raumliste',
#                       cancellable=True, step=10) as pb2:

#    for n_2, system_id in enumerate(systems):
#        if pb2.cancelled:
#            script.exit()

#        pb2.update_progress(n_2 + 1, len(systems))
#        RohrSys = System(system_id)

#        system_liste.append(RohrSys)
#        table_data.append(RohrSys.table_row())

#Sortieren nach Raumnummer
#table_data.sort()

# VEWasserSumme Liste erstellen
#VEW_VL_Liste = []
#VEW_RL_Liste = []

# Schacht Liste
#Schacht = S0


#for i in range(len(Schacht)):
#    VEW_VL_Liste.append(0)
#    VEW_RL_Liste.append(0)


# Wassermenge berechnen
#with forms.ProgressBar(title='{value}/{max_value} Wärmeabgabe Berechnen',
#                       cancellable=True, step=10) as pb3:
#    n_3 = 0
#    for space in spaces_collector:

#        if pb3.cancelled:
#            script.exit()

#        n_3 += 1
#        pb3.update_progress(n_3, len(spaces))

#        Schacht_VL = get_value(space.LookupParameter('IGF_S_Schacht_VEWasserVL'))
#        Schacht_RL = get_value(space.LookupParameter('IGF_S_Schacht_VEWasserRL'))
#        VEWasser = get_value(space.LookupParameter('IGF_S_VEWasserMenge'))
#        for ii in range(len(Schacht)):
#            if Schacht_VL == Schacht[ii]:
#                VEW_VL_Liste[ii] = VEWasser + VEW_VL_Liste[ii]
#            if Schacht_RL == Schacht[ii]:
#                VEW_RL_Liste[ii] = VEWasser + VEW_RL_Liste[ii]


#table_data_Schacht = []
#for Nr in range(len(Schacht)):
#    table_data_Schacht.append([Schacht[Nr],VEW_VL_Liste[Nr],VEW_RL_Liste[Nr]])




#output.print_table(
#    table_data=table_data_Schacht,
#    title="Schächte",
#    columns=[ 'Schacht','VEWasser_VL', 'VEWasser_RL' ]
#)


# Werte zuückschreiben + Abfrage
#if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):

#    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
#                           cancellable=True, step=10) as pb4:
#        n_4 = 0

#        t1 = Transaction(doc, "Werteschreiben")
#        t1.Start()

#        for SPACE in spaces_collector:
#
#            if pb4.cancelled:
#                script.exit()

#            n_4 += 1

#            pb4.update_progress(n_4, len(spaces))

#            Ins_Schacht = get_value(SPACE.LookupParameter('TGA_RLT_InstallationsSchacht'))
#            Ins_Schacht_Name = get_value(SPACE.LookupParameter('TGA_RLT_InstallationsSchachtName'))
#            if Ins_Schacht:
#                for num in range(len(Schacht)):
#                    if Ins_Schacht_Name == Schacht[num]:
#                        schacht_vl = SPACE.LookupParameter('IGF_S_Schacht_VEWasserMenge_VL')
#                        schacht_vl.SetValueString(str(VEW_VL_Liste[num]))
#                        schacht_rl = SPACE.LookupParameter('IGF_S_Schacht_VEWasserMenge_RL')
#                        schacht_rl.SetValueString(str(VEW_RL_Liste[num]))

#        t1.Commit()


#logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
#logger.info("{} Schacht in MEP_Raumliste".format(len(Schacht)))
