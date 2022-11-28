# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time


start = time.time()


__title__ = "3.Schächte"
__doc__ = """Schachtberechnung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()


logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()


class MEPRaum:
    Schacht_Liste = []

    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],
            ['Zu_Schacht_Nr', 'TGA_RLT_SchachtZuluft'],
            ['Zul', 'IGF_RLT_ZuluftminRaum'],
            ['Ab_Schacht_Nr', 'TGA_RLT_SchachtAbluft'],
            ['Abl', 'IGF_RLT_AbluftminRaum'],
            ['Ab_24h_Schacht_Nr', 'TGA_RLT_Schacht24hAbluft'],
            ['Abl_24h', 'IGF_RLT_AbluftminRaumL24h'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        # Schachtliste erstellen

        if self.Install_Schacht:
            self.Schacht_Liste.append(self.Install_Schacht_Name)


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

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.Install_Schacht,
            self.Install_Schacht_Name,
            self.Zu_Schacht_Nr,
            self.Zul,
            self.Ab_Schacht_Nr,
            self.Abl,
            self.Ab_24h_Schacht_Nr,
            self.Abl_24h,
            ]


    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


mepraum_liste = []
table_data = []
with forms.ProgressBar(title='{value}/{max_value} Raumliste',
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())

#Sortieren nach Raumnummer
table_data.sort()

P = []
for Da in table_data:
    if Da[4]== Da[6] == Da[8]:
        pass
    else:
        P.append(Da)

logger.info("{}".format(len(P)))

output.print_table(
    table_data=P,
    title="Mep Räume",
    columns=['name', 'nummer', 'Install_Schacht', 'Install_Schacht_Name', 'Zu_Schacht_Nr',
             'Zul', 'Ab_Schacht_Nr', 'Abl', 'Ab_24h_Schacht_Nr', 'Abl_24h', ]
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


# Luftsumme_Liste erstellen

Zuluft_Liste = []
Abluft_Liste = []
Abluft_24h_Liste = []

schacht = set(MEPRaum.Schacht_Liste)
Schacht = list(schacht)  # Schacht Liste
Schacht.sort()

for i in range(len(Schacht)):
    Zuluft_Liste.append(0)
    Abluft_Liste.append(0)
    Abluft_24h_Liste.append(0)


# Luftmenge berechnen
with forms.ProgressBar(title='{value}/{max_value} Luftmengen Berechnen',
                       cancellable=True, step=10) as pb1:
    n_0 = 0
    for space in spaces_collector:

        if pb1.cancelled:
            script.exit()

        n_0 += 1
        pb1.update_progress(n_0, len(spaces))

        Zu_Schacht_Nr = get_value(space.LookupParameter('TGA_RLT_SchachtZuluft'))
        Zul = get_value(space.LookupParameter('IGF_RLT_ZuluftminRaum'))
        Ab_Schacht_Nr = get_value(space.LookupParameter('TGA_RLT_SchachtAbluft'))
        Abl = get_value(space.LookupParameter('IGF_RLT_AbluftminRaum'))
        Ab_24h_Schacht_Nr = get_value(space.LookupParameter('TGA_RLT_Schacht24hAbluft'))
        Abl_24h = get_value(space.LookupParameter('IGF_RLT_AbluftminRaumL24h'))
        for ii in range(len(Schacht)):
            if Zu_Schacht_Nr == Schacht[ii]:
                Zuluft_Liste[ii] = Zul + Zuluft_Liste[ii]
        for jj in range(len(Schacht)):
            if Ab_Schacht_Nr == Schacht[jj]:
                Abluft_Liste[jj] = Abl + Abluft_Liste[jj]
        for kk in range(len(Schacht)):
            if Ab_24h_Schacht_Nr == Schacht[kk]:
                Abluft_24h_Liste[kk] = Abl_24h + Abluft_24h_Liste[kk]


table_data_Schacht = []
for Nr in range(len(Schacht)):
    table_data_Schacht.append([Schacht[Nr],Zuluft_Liste[Nr],Abluft_Liste[Nr],Abluft_24h_Liste[Nr]])


output.print_table(
    table_data=table_data,
    title="Mep Räume",
    columns=['name', 'nummer', 'Install_Schacht', 'Install_Schacht_Name', 'Zu_Schacht_Nr',
             'Zul', 'Ab_Schacht_Nr', 'Abl', 'Ab_24h_Schacht_Nr', 'Abl_24h', ]
)

output.print_table(
    table_data=table_data_Schacht,
    title="Luftmengen",
    columns=[ 'Schacht','Zuluft', 'Abluft', 'Abluft_24h', ]
)


# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):

    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',
                           cancellable=True, step=10) as pb2:
        n_1 = 0

        t = Transaction(doc, "Werteschreiben")
        t.Start()

        for Spaces in spaces_collector:

            if pb2.cancelled:
                script.exit()

            n_1 += 1

            pb2.update_progress(n_1, len(spaces))

            Ins_Schacht = get_value(Spaces.LookupParameter('TGA_RLT_InstallationsSchacht'))
            Ins_Schacht_Name = get_value(Spaces.LookupParameter('TGA_RLT_InstallationsSchachtName'))
            if Ins_Schacht:
                for num in range(len(Schacht)):
                    if Ins_Schacht_Name == Schacht[num]:
                        zulschacht = Spaces.LookupParameter('TGA_RLT_SchachtZuluftMenge')
                        zulschacht.SetValueString(str(Zuluft_Liste[num]))

                        ablschacht = Spaces.LookupParameter('TGA_RLT_SchachtAbluftMenge')
                        ablschacht.SetValueString(str(Abluft_Liste[num]))

                        abl24hschacht = Spaces.LookupParameter('TGA_RLT_Schacht24hAbluftMenge')
                        abl24hschacht.SetValueString(str(Abluft_24h_Liste[num]))


        t.Commit()




logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Schacht in MEP_Raumliste".format(len(Schacht)))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
