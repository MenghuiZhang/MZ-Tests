# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction


__title__ = "4.10 VE Wasser"
__doc__ = """Schacht von Lüftung Übernehmen und VE Wasser von Schacht berechnen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces) \
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()


logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()



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
        vewasser_vl = get_value(Space.LookupParameter("IGF_S_Schacht_VEWasserVL"))
        vewasser_rl = get_value(Space.LookupParameter("IGF_S_Schacht_VEWasserRL"))
        tabel_Date_MEP.append([nummer,name,luft_zu,luft_ab,vewasser_vl,vewasser_rl])
        if ins_schacht:
            schacht.append([nummer,name,install_schacht_name,luft_zu,luft_ab,vewasser_vl,vewasser_rl])


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
             'Schacht_Vorlauf','Schacht_Rücklauf']
)

output.print_table(
    table_data=schacht,
    title="Schächte",
    columns=['nummer', 'name', 'Install_Name', 'Zuluft_Schacht','Abluft_Schacht',
            'Schacht_Vorlauf','Schacht_Rücklauf']
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
            VEW_VL_Schacht = Spaces.LookupParameter("IGF_S_Schacht_VEWasserVL")
            VEW_RL_Schacht = Spaces.LookupParameter("IGF_S_Schacht_VEWasserRL")
            if Zuluft_Schacht and Zuluft_Schacht != 'None':
                VEW_VL_Schacht.Set(str(Zuluft_Schacht))
            if Abluft_Schacht and Abluft_Schacht != 'None':
                VEW_RL_Schacht.Set(str(Abluft_Schacht))


        t.Commit()


class MEPRaum:

    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties für die
        Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],
            ['SchachtVL_Nr', 'IGF_S_Schacht_VEWasserVL'],
            ['SchachtRL_Nr', 'IGF_S_Schacht_VEWasserRL'],
            ['VE_Wasser', 'IGF_S_VEWasserMenge'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))


    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return get_value(param)



    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.SchachtVL_Nr,
            self.SchachtRL_Nr,
            self.VE_Wasser
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

# VEWasserSumme Liste erstellen
VEW_VL_Liste = []
VEW_RL_Liste = []

# Schacht Liste
Schacht = S0


for i in range(len(Schacht)):
    VEW_VL_Liste.append(0)
    VEW_RL_Liste.append(0)

output.print_table(
    table_data=table_data,
    title="Mep Räume",
    columns=['Nummer', 'Name', 'Schacht_VL', 'Schacht_RL', 'VE_Wasser', ]
)

# Wassermenge berechnen
with forms.ProgressBar(title='{value}/{max_value} Wärmeabgabe Berechnen',
                       cancellable=True, step=10) as pb3:
    n_3 = 0
    for space in spaces_collector:

        if pb3.cancelled:
            script.exit()

        n_3 += 1
        pb3.update_progress(n_3, len(spaces))

        Schacht_VL = get_value(space.LookupParameter('IGF_S_Schacht_VEWasserVL'))
        Schacht_RL = get_value(space.LookupParameter('IGF_S_Schacht_VEWasserRL'))
        VEWasser = get_value(space.LookupParameter('IGF_S_VEWasserMenge'))
        for ii in range(len(Schacht)):
            if Schacht_VL == Schacht[ii]:
                VEW_VL_Liste[ii] = VEWasser + VEW_VL_Liste[ii]
            if Schacht_RL == Schacht[ii]:
                VEW_RL_Liste[ii] = VEWasser + VEW_RL_Liste[ii]


table_data_Schacht = []
for Nr in range(len(Schacht)):
    table_data_Schacht.append([Schacht[Nr],VEW_VL_Liste[Nr],VEW_RL_Liste[Nr]])




output.print_table(
    table_data=table_data_Schacht,
    title="Schächte",
    columns=[ 'Schacht','VEWasser_VL', 'VEWasser_RL' ]
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
                        schacht_vl = SPACE.LookupParameter('IGF_S_Schacht_VEWasserMenge_VL')
                        schacht_vl.SetValueString(str(VEW_VL_Liste[num]))
                        schacht_rl = SPACE.LookupParameter('IGF_S_Schacht_VEWasserMenge_RL')
                        schacht_rl.SetValueString(str(VEW_RL_Liste[num]))

        t1.Commit()


logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Schacht in MEP_Raumliste".format(len(Schacht)))


