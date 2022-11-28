# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
from operator import itemgetter
import csv


start = time.time()


__title__ = "8.Luftauslass"
__doc__ = """Informationen der Luftauslaesse zeigen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

# Luftauslässe aus aktueller Projekt
DuctTerminal_collector = DB.FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(DB.BuiltInCategory.OST_DuctTerminal)
ducts = DuctTerminal_collector.ToElementIds()

phase = list(doc.Phases)[-1]

logger.info("{} Luftauslässe ausgewählt".format(len(ducts)))

if not ducts:
    logger.error("Keine Luftauslässe in aktueller Ansicht gefunden")
    script.exit()


class DUCT_Terminal:
    Auslass_Liste = []
    Auslass_Tag = []
    Raumnummer = []

    def __init__(self, element_id):
        """
        Definiert Luftauslässe Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr_Auslass = [
            ['S_ABZ', 'Systemabkürzung'],
            ['S_Name', 'Systemname'],
            ['S_Grupp', 'Systemklassifizierung'],
            ['KZ', 'Kennzeichen'],
            ['Vmin', 'IGF_RLT_AuslassVolumenstromMin'],
            ['Vmax', 'IGF_RLT_AuslassVolumenstromMax'],
            ['V_Nacht', 'IGF_RLT_AuslassVolumenstromNacht'],
#            ['V_TNacht', 'TGA_RLT_AbluftSumme24h'],
        ]

        attr_Raum = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['zuluft_min', 'TGA_RLT_ZuluftminRaum'],
            ['abluft_min', 'TGA_RLT_AbluftminRaum'],
            ['abluft_labor', 'TGA_RLT_AbluftSummeLabor'],
            ['abluft_24h', 'TGA_RLT_AbluftSumme24h'],
            ['TS_Zuluft', 'Tatsächlicher Zuluftstrom'],
            ['TS_Abluft', 'Tatsächlicher Abluftstrom'],
        ]

        for a in attr_Auslass:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr_Auslass(revit_name))

        for b in attr_Raum:
            python_name, revit_name = b
            setattr(self, python_name, self.__get_element_attr_Raum(revit_name))

        logger.info(50 * "=")
        logger.info("{}\t{}\t{}\t{}".format(self.nummer, self.name, self.S_Name, self.S_ABZ))

        self.Raumnummer.append(self.nummer)

        self.Auslass_Tag = [self.nummer, self.name, self.zuluft_min, self.abluft_min,
                      self.TS_Zuluft, self.TS_Abluft,self.abluft_labor,
                      self.abluft_24h, self.S_ABZ, self.S_Name, self.S_Grupp,
                      self.KZ, self.Vmin, self.Vmax, self.V_Nacht]

        self.Auslass_Liste.append(self.Auslass_Tag)


    def __get_element_attr_Auslass(self, param_name):
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


    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.zuluft_min,
            self.abluft_min,
            self.abluft_labor,
            self.abluft_24h,
            self.S_Name,
            self.S_ABZ,
            self.S_Grupp,
            self.KZ,
            self.Vmin,
            self.Vmax,
            self.V_Nacht,
        ]

    def __repr__(self):
        return "MEPRaum({})".format(self.element_id)

    def __str__(self):
        return '{}\t{}'.format(self.nummer, self.name)


table_data = []
Luftauslass_Liste = []

with forms.ProgressBar(title='{value}/{max_value} Luftmengenberechnung',
                       cancellable=True, step=10) as pb:

    for n, duct_id in enumerate(ducts):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(ducts))
        Auslass = DUCT_Terminal(duct_id)

        Luftauslass_Liste.append(Auslass)
        table_data.append(Auslass.table_row())

Raum_Nr = set(DUCT_Terminal.Raumnummer)
Raumnummer_Liste = list(Raum_Nr)

if not all(Raumnummer_Liste):
    Raumnummer_Liste.remove(None)

Raumnummer_Liste.sort()

al_Liste = DUCT_Terminal.Auslass_Liste
al_Liste = sorted(al_Liste,key=itemgetter(10,8),reverse= True)


AL_Liste = []
for RN in Raumnummer_Liste:
    for Al in al_Liste:
        if Al[0] == RN:
            AL_Liste.append(Al)


Raum_Liste = []

for RaumNR in Raumnummer_Liste:
    Einzel_Raum = []
    for AUSLASS in AL_Liste:
        if AUSLASS[0] == RaumNR:
            Einzel_Raum.append(AUSLASS)
    Raum_Liste.append(Einzel_Raum)

Raum_AUSlass_Liste = []
for RAUM in Raum_Liste:
    Raum_AUSlass = []
    if any(RAUM):
        Column = [str(RAUM[0][0]) + ' / ' + str(RAUM[0][1]),'V_min','Zu_Raum','Ab_Raum']
        Raum_AUSlass.append(Column)
        i = 0
        j = 0
        k = 0
        Summe_Zul = 0
        Summe_Abl = 0
        Summe_Abl_24h = 0
        Summe_AblUndAbl24h = 0
        Summe_Rauml = 0
        ZulAUS = ''
        AblAUS = ''
        Raum_Luft = ''

        for Ausl in RAUM:
            if Ausl[10] == 'Zuluft':
                i += 1
                Summe_Zul = Summe_Zul + int(round(Ausl[12]))
                Zul_AUS = ['Zu_' + str(i), Ausl[12], Ausl[2], Ausl[3]]
                Raum_AUSlass.append(Zul_AUS)
            elif Ausl[10] == 'Abluft':
                if Ausl[8] != 'ETA_24h':
                    j += 1
                    Abl_AUS = ['Ab_' + str(j), Ausl[12], Ausl[2], Ausl[3]]
                    Raum_AUSlass.append(Abl_AUS)
                    Summe_Abl = Summe_Abl + int(round(Ausl[12]))
                else:
                    k += 1
                    Abl24h_AUS = ['Ab24h_' + str(k), Ausl[12], Ausl[2], Ausl[3]]
                    Raum_AUSlass.append(Abl24h_AUS)
                    Summe_Abl_24h = Summe_Abl_24h + int(round(Ausl[12]))
        Summe_AblUndAbl24h =  Summe_Abl + Summe_Abl_24h
        Summe_Rauml = Summe_Zul - Summe_AblUndAbl24h

        if Summe_Zul == int(RAUM[0][2]):
            ZulAUS = 'OK'
        else:
            ZulAUS = 'FALSE'

        if Summe_Abl == int(RAUM[0][3]):
            AblAUS = 'OK'
        else:
            AblAUS = 'FALSE'

        if Summe_Rauml == 0 and ZulAUS == 'OK' and AblAUS == 'OK':
            Raum_Luft = 'OK'
        else:
            Raum_Luft = 'FALSE'


        Summe_ZUL = ['Summe_Zu', Summe_Zul, ZulAUS, ' ']
        Summe_ABL = ['Summe_Ab', Summe_Abl, RAUM[0][2], RAUM[0][3]]
        Summe_ABL24H = ['Summe_24h', Summe_Abl_24h, RAUM[0][2], RAUM[0][3]]
        Summe_ABLUND24H = ['Summe_AbUndAb24h', Summe_AblUndAbl24h, ' ', AblAUS]
        Summe_RaumL = ['Summe_Raum', Summe_Rauml, Raum_Luft, '' ]

        Raum_AUSlass.insert(i+1,Summe_ZUL)
        Raum_AUSlass.insert(i+j+2,Summe_ABL)
        Raum_AUSlass.append(Summe_ABL24H)
        Raum_AUSlass.append(Summe_ABLUND24H)
        Raum_AUSlass.append(Summe_RaumL)
    Raum_AUSlass_Liste.append(Raum_AUSlass)

output.print_table(
    table_data=table_data,
    title="Luftmengen_Unsortiert",
    columns=['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name',
             'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
             'V_Nacht_Auslass', ]
)


output.print_table(
    table_data = AL_Liste,
    title = "Luftmengen_Sortiert(nach Raumnummer, Systemklassifizierung, Systemabkürzung)",
    columns = ['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
             'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass', 'V_Nacht_Auslass', ]
)

Raum_Auslass_Liste = []
for RAL in Raum_AUSlass_Liste:
    for Raum_aus in RAL:
        Raum_Auslass_Liste.append(Raum_aus)

output.print_table(
    table_data = Raum_Auslass_Liste,
    title = "Luftmengen",
    columns = ['Raum_Nr/Name', 'V_min', 'Zu_Raum', 'Ab_Raum']
)


ops = ['Raumnummer', 'Systemname', 'Systemabkürzung', 'Schließen']
c = forms.CommandSwitchWindow.show(ops,message = 'Filter nach')
if c == 'Raumnummer':
    x = rpw.ui.forms.TextInput('Raumnummer: ', default = "0")
    while x != '0':
        al_liste = []
        logger.info("Raumnummer ist:{}".format(x))
        for Al in AL_Liste:
            if Al[0] == x:
                al_liste.append(Al)

        if any(al_Liste):
            output.print_table(
                table_data = al_liste,
                title = "Luftmengen",
                columns = ['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
                         'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass', 'V_Nacht_Auslass', ]
            )
        else:
            logger.error("Die Raumnummer ist falsch. Bitte erneut eingeben.")


        Raum_Auslass = []
        if any(al_liste):
            column = [str(al_liste[0][0]) + ' / ' + str(al_liste[0][1]),'V_min','Zu_Raum','Ab_Raum']
            Raum_Auslass.append(column)
            ii = 0
            jj = 0
            Kk = 0
            Summe_Zu = 0
            Summe_Ab = 0
            Summe_24h = 0
            Summe_AbUndAb24h = 0
            Summe_Raum = 0
            ZuAUS = ''
            AbAUS = ''
            AbRaum = ''

            for Aus in al_liste:
                if Aus[10] == 'Zuluft':
                    ii += 1
                    Summe_Zu = Summe_Zu + int(round(Aus[12]))
                    Zu_AUS = ['Zu_' + str(ii), Aus[12], Aus[2], Aus[3]]
                    Raum_Auslass.append(Zu_AUS)
                elif Aus[10] == 'Abluft':
                    if Aus[8] != 'ETA_24h':
                        jj += 1
                        Ab_AUS = ['Ab_' + str(jj), Aus[12], Aus[2], Aus[3]]
                        Raum_Auslass.append(Ab_AUS)
                        Summe_Ab = Summe_Ab + int(round(Aus[12]))
                    else:
                        Kk += 1
                        Ab24h_AUS = ['Ab24h_' + str(kk), Aus[12], Aus[2], Aus[3]]
                        Raum_Auslass.append(Ab24h_AUS)
                        Summe_24h = Summe_24h + int(round(Aus[12]))
            Summe_AbUndAb24h =  Summe_Ab + Summe_24h
            Summe_Raum = Summe_Zu - Summe_AbUndAb24h

            if Summe_Zu == int(al_liste[0][2]):
                ZuAUS = 'OK'
            else:
                ZuAUS = 'FALSE'

            if Summe_Ab == int(al_liste[0][3]):
                AbAUS = 'OK'
            else:
                AbAUS = 'FALSE'

            if Summe_Raum == 0 and ZuAUS == 'OK' and AbAUS == 'OK':
                AbRaum = 'OK'
            else:
                AbRaum = 'FALSE'


            Summe_ZU = ['Summe_Zu', Summe_Zu, ZuAUS, ' ']
            Summe_AB = ['Summe_Ab', Summe_Ab, al_liste[0][2], al_liste[0][3]]
            Summe_AB24H = ['Summe_24h', Summe_24h, al_liste[0][2], al_liste[0][3]]
            Summe_ABUND24H = ['Summe_AbUndAb24h', Summe_AbUndAb24h, ' ', AbAUS]
            Summe_Raum = ['Summe_Raum', Summe_Raum, AbRaum, '' ]

            Raum_Auslass.insert(ii+1,Summe_ZU)
            Raum_Auslass.insert(ii+jj+2,Summe_AB)
            Raum_Auslass.append(Summe_AB24H)
            Raum_Auslass.append(Summe_ABUND24H)
            Raum_Auslass.append(Summe_Raum)
            output.print_table(
                table_data = Raum_Auslass,
                title = "Luftmengen",
                columns = ['Raum_Nr/Name', 'V_min', 'Zu_Raum', 'Ab_Raum']
            )

        x = rpw.ui.forms.TextInput('Raumnummer: ', default = "0")
elif c == 'Systemname':
    x = rpw.ui.forms.TextInput('Systemname: ', default = "0")
    while x != '0':
        al_liste = []
        logger.info("Systemname ist:{}".format(x))
        for Al in AL_Liste:
            if Al[8] == x:
                al_liste.append(Al)

        if any(al_liste):
            output.print_table(
                table_data = al_liste,
                title = "Luftmengen",
                columns = ['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
                         'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass', 'V_Nacht_Auslass', ]
            )
        else:
            logger.error("Der Systemname ist falsch. Bitte erneut eingeben.")

        x = rpw.ui.forms.TextInput('Systemname: ', default = "0")



elif c == 'Systemabkürzung':
    x = rpw.ui.forms.TextInput('Systemabkürzung: ', default = "0")
    while x != '0':
        al_liste = []
        logger.info("Systemabkürzung ist:{}".format(x))
        for Al in AL_Liste:
            if Al[9] == x:
                al_liste.append(Al)

        if any(al_liste):
            output.print_table(
                table_data = al_liste,
                title = "Luftmengen",
                columns = ['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
                         'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass', 'V_Nacht_Auslass', ]
            )

        else:
            logger.error("Die Systemabkürzung ist falsch. Bitte erneut eingeben.")

        x = rpw.ui.forms.TextInput('Systemabkürzung: ', default = "0")

elif c == 'Schließen':
    pass





logger.info("{} Luftauslässe in Luftauslass_Liste".format(len(Luftauslass_Liste)))


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
