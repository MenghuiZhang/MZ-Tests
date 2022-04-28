# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time


start = time.time()


__title__ = "P_KüW"
__doc__ = """Schacht von Lüftung Übernehmen, PKW-Leistung von MEP-Raum und Leistung bzw. Wassermenge von Schacht berechnen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()


logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
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


# Schachtsname von Zu- und Abluft jedes MEP Raum lesen
tabel_Date_MEP = []
schacht = []
with forms.ProgressBar(title='{value}/{max_value} Raumliste',
                       cancellable=True, step=10) as pb:

    n = 0
    for Space in spaces_collector:

        n += 1

        if pb.cancelled:
            script.exit()

        pb.update_progress(n, len(spaces))

        nummer = get_value(Space.LookupParameter("Nummer"))
        name = get_value(Space.LookupParameter("Name"))
        ins_schacht = get_value(Space.LookupParameter("TGA_RLT_InstallationsSchacht"))
        install_schacht_name = get_value(Space.LookupParameter("TGA_RLT_InstallationsSchachtName"))
        luft_zu = get_value(Space.LookupParameter("TGA_RLT_SchachtZuluft"))
        luft_ab = get_value(Space.LookupParameter("TGA_RLT_SchachtAbluft"))
        pkw_vl = get_value(Space.LookupParameter("IGF_S_Schacht_PKW-VL"))
        pkw_rl = get_value(Space.LookupParameter("IGF_S_Schacht_PKW-RL"))
        tabel_Date_MEP.append([nummer,name,luft_zu,luft_ab,pkw_vl,pkw_rl])
        if ins_schacht:
            schacht.append([nummer,name,install_schacht_name,luft_zu,luft_ab,pkw_vl,pkw_rl])


        logger.info(30 * "=")
        logger.info("{} {}".format(nummer, name))

# Schachtliste, Schachtliste von Zu- und Abluft

S_0 = []
S_1 = []
S_2 = []
for Sch in schacht:
    S_0.append(Sch[2])
for MEP in tabel_Date_MEP:
    S_1.append(MEP[2])
    S_2.append(MEP[3])
s_0 = set(S_0)
S0 = list(s_0)
s_1 = set(S_1)
S1 = list(s_1)
s_2 = set(S_2)
S2 = list(s_2)
S0.sort()
S1.sort()
S2.sort()
logger.info("Schachtlist: {}".format(S0))
logger.info("Schachtlist von Zuluft: {}".format(S1))
logger.info("Schachtlist von Abluft: {}".format(S2))



output.print_table(
    table_data=tabel_Date_MEP,
    title="Mep Räume",
    columns=['nummer', 'name',  'Zuluft_Schacht','Abluft_Schacht',
             'PKW-VL_Schacht', 'PKW-RL_Schacht']
)

output.print_table(
    table_data=schacht,
    title="Schächte",
    columns=['nummer', 'name', 'Install_Name', 'Zuluft_Schacht','Abluft_Schacht',
             'PKW-VL_Schacht','PKW-RL_Schacht']
)

# Schächte von Lüftung übernehmen + Abfrage
if forms.alert('Schächte von Lüftung übernehmen?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Schächte übernehmen',
                           cancellable=True, step=10) as pb1:
        n_1 = 0

        t = Transaction(doc, "Schächte übernehmen")
        t.Start()

        for Spaces in spaces_collector:

            if pb1.cancelled:
                script.exit()

            n_1 += 1

            pb1.update_progress(n_1, len(spaces))

            Nummer = get_value(Spaces.LookupParameter("Nummer"))
            Name = get_value(Spaces.LookupParameter("Name"))
            Zuluft_Schacht = get_value(Spaces.LookupParameter("TGA_RLT_SchachtZuluft"))
            Abluft_Schacht = get_value(Spaces.LookupParameter("TGA_RLT_SchachtAbluft"))
            PKW_VL_Schacht = Spaces.LookupParameter("IGF_S_Schacht_PKW-VL")
            PKW_RL_Schacht = Spaces.LookupParameter("IGF_S_Schacht_PKW-RL")
            if Zuluft_Schacht and Zuluft_Schacht != 'None':
                PKW_VL_Schacht.Set(str(Zuluft_Schacht))
            if Abluft_Schacht and Abluft_Schacht != 'None':
                PKW_RL_Schacht.Set(str(Abluft_Schacht))

        t.Commit()


class MEPRaum:

    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],
            ['Schacht_VL', 'IGF_S_Schacht_PKW-VL'],
            ['Schacht_RL', 'IGF_S_Schacht_PKW-RL'],
            ['Wassermenge', 'IGF_S_PKW_Menge'],
            ['Temp_VL', 'IGF_S_PKW-VL_Temp'],
            ['Temp_RL', 'IGF_S_PKW-RL_Temp'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        self.Leistung = self.LeistungBerechnen()



    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
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

    def LeistungBerechnen(self):
        leistung = 0

        if self.Wassermenge and self.Temp_VL and self.Temp_RL:
            leistung = 4200 * self.Wassermenge * (self.Temp_VL- self.Temp_RL)

        return leistung


    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.Schacht_VL,
            self.Schacht_RL,
            self.Wassermenge,
            self.Temp_VL,
            self.Temp_RL,
            self.Leistung,
            ]


    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


mepraum_liste = []
table_data = []
with forms.ProgressBar(title='{value}/{max_value} Raumliste',
                       cancellable=True, step=10) as pb2:

    for n_2, space_id in enumerate(spaces):
        if pb2.cancelled:
            script.exit()

        pb2.update_progress(n_2 + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())

#Sortieren nach Raumnummer
table_data.sort()

# PKW_Menge_Liste erstellen
PKW_M_VL_Liste = []
PKW_M_RL_Liste = []

# PKW_Leistung_Liste erstellen
PKW_L_VL_Liste = []
PKW_L_RL_Liste = []

# Schacht Liste
Schacht = S0

for i in range(len(Schacht)):
    PKW_M_VL_Liste.append(0)
    PKW_M_RL_Liste.append(0)
    PKW_L_VL_Liste.append(0)
    PKW_L_RL_Liste.append(0)

output.print_table(
    table_data=table_data,
    title="Mep Räume",
    columns=['Nummer', 'Name',  'Schacht_VL', 'Schacht_RL', 'Wassermenge',
             'Temp_VL', 'Temp_RL', 'Heizleistung', ]
)

# Werte zuückschreiben + Abfrage
if forms.alert('Leistung[W] von MEPRaum in Modell schreiben?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb4:
        n_4 = 0
        t1 = Transaction(doc, "Leistung Schreiben")
        t1.Start()

        for SPACE in spaces_collector:

            if pb4.cancelled:
                script.exit()

            n_4 += 1

            pb4.update_progress(n_4, len(spaces))

            Raum_Nr = get_value(SPACE.LookupParameter('Nummer'))
            for Raum_Date in table_data:
                if Raum_Date[0] == Raum_Nr:
                    SPACE.LookupParameter('IGF_S_PKW_Leistung').SetValueString(str(Raum_Date[7]))
        t1.Commit()


# PKW_Leistung berechnen
with forms.ProgressBar(title='{value}/{max_value} Heizleistung von Schacht Berechnen',
                       cancellable=True, step=10) as pb3:
    n_3 = 0
    for space in spaces_collector:

        if pb3.cancelled:
            script.exit()

        n_3 += 1
        pb3.update_progress(n_3, len(spaces))
        PKW_V_S = get_value(space.LookupParameter('IGF_S_Schacht_PKW-VL'))
        PKW_R_S = get_value(space.LookupParameter('IGF_S_Schacht_PKW-RL'))
        Wassermenge = get_value(space.LookupParameter('IGF_S_PKW_Menge'))
        Leistung = get_value(space.LookupParameter('IGF_S_PKW_Leistung'))

        for ii in range(len(Schacht)):
            if PKW_V_S == Schacht[ii]:
                PKW_M_VL_Liste[ii] = Wassermenge + PKW_M_VL_Liste[ii]
                PKW_L_VL_Liste[ii] = Leistung + PKW_L_VL_Liste[ii]
            if PKW_R_S == Schacht[ii]:
                PKW_M_RL_Liste[ii] = Wassermenge + PKW_M_RL_Liste[ii]
                PKW_L_RL_Liste[ii] = Leistung + PKW_L_RL_Liste[ii]

table_data_Schacht = []
for Nr in range(len(Schacht)):
    table_data_Schacht.append([Schacht[Nr], PKW_M_VL_Liste[Nr], PKW_M_RL_Liste[Nr],PKW_L_VL_Liste[Nr], PKW_L_RL_Liste[Nr],])



output.print_table(
    table_data=table_data_Schacht,
    title="Schächte",
    columns=[ 'Schacht','PKW-VL_Menge','PKW-RL_Menge','PKW-VL_Leistung','PKW-RL_Leistung',  ]
)

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb4:
        n_4 = 0
        t1 = Transaction(doc, "Werteschreiben")
        t1.Start()

        for SPACE in spaces_collector:

            if pb4.cancelled:
                script.exit()

            n_4 += 1

            pb4.update_progress(n_4, len(spaces))

            Ins_Schacht = get_value(SPACE.LookupParameter('TGA_RLT_InstallationsSchacht'))
            Ins_Schacht_Name = get_value(SPACE.LookupParameter('TGA_RLT_InstallationsSchachtName'))

            if Ins_Schacht:
                for num in range(len(Schacht)):
                    if Ins_Schacht_Name == Schacht[num]:
                        Schacht_V = SPACE.LookupParameter('IGF_S_Schacht_PKW-VL_Menge')
                        Schacht_R = SPACE.LookupParameter('IGF_S_Schacht_PKW-RL_Menge')
                        SchachtV = SPACE.LookupParameter('IGF_S_Schacht_PKW-VL_Leistung')
                        SchachtR = SPACE.LookupParameter('IGF_S_Schacht_PKW-RL_Leistung')
                        Schacht_V.SetValueString(str(PKW_M_VL_Liste[num]))
                        Schacht_R.SetValueString(str(PKW_M_RL_Liste[num]))
                        SchachtV.SetValueString(str(PKW_L_VL_Liste[num]))
                        SchachtR.SetValueString(str(PKW_L_RL_Liste[num]))

        t1.Commit()


logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Schacht in MEP_Raumliste".format(len(Schacht)))

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
