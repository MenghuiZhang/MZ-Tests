# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time

start = time.time()


__title__ = "5.Geschoss"
__doc__ = """Kühllastberechnung"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc


# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()


class MEPRaum:
    Ebene_Nr_Liste = []
    Schacht_Liste = []
    Zu_Anl_Nr_Liste = []
    Ab_Anl_Nr_Liste = []
    Ab_24h_Anl_Nr_Liste = []
    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties für die
        Kühlleistung Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['ebene', 'Ebene'],
            ['ebene_nr', 'IGF_RLT_Verteilung_EbenenSortieren'],
            ['Zu_Anl_Nr', 'TGA_RLT_AnlagenNrZuluft'],
            ['Ab_Anl_Nr', 'TGA_RLT_AnlagenNrAbluft'],
            ['Ab_Anl_Nr_24h', 'TGA_RLT_AnlagenNr24hAbluft'],
            ['Install_Schacht', 'TGA_RLT_InstallationsSchacht'],
            ['Install_Schacht_Name', 'TGA_RLT_InstallationsSchachtName'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))

        # Anlagenliste erstellen
        if self.Zu_Anl_Nr:
            self.Zu_Anl_Nr_Liste.append(self.Zu_Anl_Nr)

        if self.Ab_Anl_Nr:
            self.Ab_Anl_Nr_Liste.append(self.Ab_Anl_Nr)

        if self.Ab_Anl_Nr_24h:
            self.Ab_24h_Anl_Nr_Liste.append(self.Ab_Anl_Nr_24h)


        # Schachtliste erstellen

        if self.Install_Schacht:
            self.Schacht_Liste.append(self.Install_Schacht_Name)

        # EbeneNrliste erstellen
        if self.ebene:
            self.Ebene_Nr_Liste.append(self.ebene_nr)


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
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)





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




table_data = []
mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} Kühlleistung berechnen',
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())


# Luft_Liste erstellen

Zu = set(MEPRaum.Zu_Anl_Nr_Liste)
ZuAnlNrListe = list(Zu)
for i in range(len(ZuAnlNrListe)):
    Zu_Anl_int = int(ZuAnlNrListe[i])
    ZuAnlNrListe[i] = Zu_Anl_int
ZuAnlNrListe.sort()

Ab = set(MEPRaum.Ab_Anl_Nr_Liste)
AbAnlNrListe = list(Ab)
for ii in range(len(AbAnlNrListe)):
    Ab_Anl_int = int(AbAnlNrListe[ii])
    AbAnlNrListe[ii] = Ab_Anl_int
AbAnlNrListe.sort()


Ab_24h = set(MEPRaum.Ab_24h_Anl_Nr_Liste)
Ab24hAnlNrListe = list(Ab_24h)
for iii in range(len(Ab24hAnlNrListe)):
    Ab24h_Anl_int = int(Ab24hAnlNrListe[iii])
    Ab24hAnlNrListe[iii] = Ab24h_Anl_int
Ab24hAnlNrListe.sort()

Eb = set(MEPRaum.Ebene_Nr_Liste)
EbeneNrListe = list(Eb)
EbeneNrListe.sort()

Sc = set(MEPRaum.Schacht_Liste)
SchachtListe = list(Sc)
SchachtListe.sort()

EbenenListe = []
for i in range(len(EbeneNrListe)):
    EbenenListe.append('')

for S in spaces_collector:
    E_Nr = get_value(S.LookupParameter('IGF_RLT_Verteilung_EbenenSortieren'))
    E = get_value(S.LookupParameter('IGF_RLT_Verteilung_EbenenName'))
    for n in range(len(EbeneNrListe)):
        if E_Nr == EbeneNrListe[n]:
            EbenenListe[n] = E
# Luftverteilung_List


for f in range(len(SchachtListe)):
    Schacht_Zu = []
    Schacht_Ab = []
    Schacht_Ab24h = []

    for e in range(len(SchachtListe)):
        Ebene_zu = []
        Ebene_ab = []
        Ebene_ab24h = []

        for d in range(len(EbeneNrListe)):
            Anl_zu = []
            Anl_ab = []
            Anl_ab24h = []
            for a in range(len(ZuAnlNrListe)):
                Anl_zu.append(0)
            for b in range(len(AbAnlNrListe)):
                Anl_ab.append(0)
            for c in range(len(Ab24hAnlNrListe)):
                Anl_ab24h.append(0)

            Ebene_zu.append(Anl_zu)
            Ebene_ab.append(Anl_ab)
            Ebene_ab24h.append(Anl_ab24h)


        Schacht_Zu.append(Ebene_zu)
        Schacht_Ab.append(Ebene_ab)
        Schacht_Ab24h.append(Ebene_ab24h)



# berechnen
for Space in spaces_collector:
    Zu_Schacht = get_value(Space.LookupParameter('TGA_RLT_SchachtZuluft'))
    Ab_Schacht = get_value(Space.LookupParameter('TGA_RLT_SchachtAbluft'))
    Ab24h_Schacht = get_value(Space.LookupParameter('TGA_RLT_Schacht24hAbluft'))
    Zu_Anl = get_value(Space.LookupParameter('TGA_RLT_AnlagenNrZuluft'))
    Ab_Anl = get_value(Space.LookupParameter('TGA_RLT_AnlagenNrAbluft'))
    Ab24h_Anl = get_value(Space.LookupParameter('TGA_RLT_AnlagenNr24hAbluft'))
    EbenenNr = get_value(Space.LookupParameter('IGF_RLT_Verteilung_EbenenSortieren'))
    Zul = get_value(Space.LookupParameter('TGA_RLT_ZuluftminRaum'))
    Abl = get_value(Space.LookupParameter('TGA_RLT_AbluftminRaum'))
    Abl_24h = get_value(Space.LookupParameter('TGA_RLT_AbluftSumme24h'))
    for aa in range(len(SchachtListe)):
        if Zu_Schacht == SchachtListe[aa]:
            S_zu = Schacht_Zu[aa]
            for bb in range(len(EbeneNrListe)):
                if EbenenNr == EbeneNrListe[bb]:
                    E_zu = S_zu[bb]
                    for cc in range(len(ZuAnlNrListe)):
                        if Zu_Anl == str(ZuAnlNrListe[cc]):
                            E_zu[cc] = E_zu[cc] + Zul

        if Ab_Schacht == SchachtListe[aa]:
            S_ab = Schacht_Ab[aa]
            for dd in range(len(EbeneNrListe)):
                if EbenenNr == EbeneNrListe[dd]:
                    E_ab = S_ab[dd]
                    for ee in range(len(AbAnlNrListe)):
                        if Ab_Anl == str(AbAnlNrListe[ee]):
                            E_ab[ee] = E_ab[ee] + Abl

        if Ab24h_Schacht == SchachtListe[aa]:
            S_ab24h = Schacht_Ab24h[aa]
            for ff in range(len(EbeneNrListe)):
                if EbenenNr == EbeneNrListe[ff]:
                    E_ab24h = S_ab24h[ff]
                    for gg in range(len(Ab24hAnlNrListe)):
                        if Ab24h_Anl == str(Ab24hAnlNrListe[gg]):
                            E_ab24h[gg] = E_ab24h[gg] + Abl_24h


Zu_Verteilen = []
Ab_Verteilen = []
Ab24h_Verteilen = []
for z in range(len(SchachtListe)):
    Zu_Verteilen.append('')
    Ab_Verteilen.append('')
    Ab24h_Verteilen.append('')

for Spaces in spaces_collector:
    Ins_Schacht = get_value(Spaces.LookupParameter('TGA_RLT_InstallationsSchacht'))
    Ins_Schacht_Name = get_value(Spaces.LookupParameter('TGA_RLT_InstallationsSchachtName'))
    EbenenNummer = get_value(Space.LookupParameter('IGF_RLT_Verteilung_EbenenSortieren'))
    Ebenen = get_value(Space.LookupParameter('IGF_RLT_Verteilung_EbenenName'))
    luft_zu = ''
    luft_ab = ''
    luft_ab24h = ''
    if Ins_Schacht:
        for num in range(len(SchachtListe)):
            if Ins_Schacht_Name == SchachtListe[num]:
                Sch_zu = Schacht_Zu[num]
                Sch_ab = Schacht_Ab[num]
                Sch_ab24h = Schacht_Ab24h[num]

                for aaa in range(len(EbenenListe)):
                    ebene = EbenenListe[aaa]
                    ebe_zu = Sch_zu[aaa]
                    ebe_ab = Sch_ab[aaa]
                    ebe_ab24h = Sch_ab24h[aaa]

                    if any(ebe_zu):
                        luft_zu = luft_zu + str(ebene) +': '

                        for bbb in range(len(ZuAnlNrListe)):
                            if ebe_zu[bbb] > 0:
                                luft_zu = luft_zu + 'Anl ' + str(ZuAnlNrListe[bbb]) + '=' + str(ebe_zu[bbb]) + ' m3/h,'

                    if any(ebe_ab):
                        luft_ab = luft_ab + str(ebene) +': '

                        for ccc in range(len(AbAnlNrListe)):
                            if ebe_ab[ccc] > 0:
                                luft_ab = luft_ab + 'Anl ' + str(AbAnlNrListe[ccc]) + '=' + str(ebe_ab[ccc]) + ' m3/h,'

                    if any(ebe_ab24h):
                        luft_ab24h = luft_ab24h + str(ebene) +': '

                        for ddd in range(len(Ab24hAnlNrListe)):
                            if ebe_ab24h[ddd] > 0:
                                luft_ab24h = luft_ab24h + 'Anl ' + str(Ab24hAnlNrListe[ddd]) + '=' + str(ebe_ab24h[ddd]) + ' m3/h,'

                Zu_Verteilen[num] = luft_zu
                Ab_Verteilen[num] = luft_ab
                Ab24h_Verteilen[num] = luft_ab24h


        logger.info("{} Zuluftmengen ist {}".format(Ins_Schacht_Name, luft_zu))
        logger.info("{} Abluftmengen ist {}".format(Ins_Schacht_Name, luft_ab))
        logger.info("{} Abluft24hmengen ist {}".format(Ins_Schacht_Name, luft_ab24h))
        logger.info(30 * "=")




output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=['Nummer', 'Name', ]
)

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    t = Transaction(doc, "Werteschreiben")
    t.Start()
    for SPACE in spaces_collector:
        INS_Schacht = get_value(SPACE.LookupParameter('TGA_RLT_InstallationsSchacht'))
        INS_Schacht_Name = get_value(SPACE.LookupParameter('TGA_RLT_InstallationsSchachtName'))
        if INS_Schacht:
            for NUM in range(len(SchachtListe)):
                if INS_Schacht_Name == SchachtListe[NUM]:
                    Zul_Verteilung = SPACE.LookupParameter('IGF_RLT_VerteilungZuluft')
                    Zul_Verteilung.Set(Zu_Verteilen[NUM])

                    Abl_Verteilung = SPACE.LookupParameter('IGF_RLT_VerteilungAbluft')
                    Abl_Verteilung.Set(Ab_Verteilen[NUM])

                    Ab24h_Verteilung = SPACE.LookupParameter('IGF_RLT_Verteilung24hAbluft')
                    Ab24h_Verteilung.Set(Ab24h_Verteilen[NUM])


    t.Commit()



logger.info("{} Räume in MEP_Raumliste".format(len(mepraum_liste)))
logger.info("{} Schächte in Schächteliste".format(len(SchachtListe)))
logger.info("Schächte: {}".format(SchachtListe))
logger.info("{} Ebene in Ebeneliste".format(len(EbenenListe)))
logger.info("Ebene: {}".format(EbenenListe))
logger.info("{} Zuluftanlagen in Zuluftanlagenliste".format(len(ZuAnlNrListe)))
logger.info("Zuluftanlagen: {}".format(ZuAnlNrListe))
logger.info("{} Abluftanlagen in Abluftanlagenliste".format(len(AbAnlNrListe)))
logger.info("Abluftanlagen: {}".format(AbAnlNrListe))
logger.info("{} Abluft_24h_Anlagen in Abluft_24h_Anlagenliste".format(len(Ab24hAnlNrListe)))
logger.info("Abluft_24h_Anlagen: {}".format(Ab24hAnlNrListe))
logger.info("{}".format(Zu_Verteilen))
logger.info("{}".format(Ab_Verteilen))
logger.info("{}".format(Ab24h_Verteilen))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
