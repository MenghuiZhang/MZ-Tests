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
            ['V_TNacht', 'IGF_RLT_AuslassVolumenstromTiefeNacht'],
        ]

        attr_Raum = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['zuluft_min', 'IGF_RLT_ZuluftminRaum'],
            ['zuluft_max', 'IGF_RLT_ZuluftmaxRaum'],
            ['zuluft_nacht', 'IGF_RLT_ZuluftNachtRaum'],
            ['zuluft_Tnacht', 'IGF_RLT_ZuluftTieferNachtRaum'],
            ['abluft_min_ohne', 'IGF_RLT_AbluftminRaum'],
            ['abluft_min_mit', 'IGF_RLT_AbluftminRaumL24h'],
            ['abluft_max', 'IGF_RLT_AbluftmaxRaum'],
            ['abluft_nacht', 'IGF_RLT_AbluftNachtRaum'],
            ['abluft_Tnacht', 'IGF_RLT_AbluftTieferNachtRaum'],
            ['abluft_labor', 'IGF_RLT_AbluftminSummeLabor'],
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

        self.Auslass_Tag = [self.nummer, self.name, self.zuluft_min, self.zuluft_max,
                            self.zuluft_nacht, self.zuluft_Tnacht, self.abluft_min_ohne,
                            self.abluft_min_mit,self.abluft_max, self.abluft_nacht,
                            self.abluft_Tnacht, self.abluft_labor,
                            self.abluft_24h, self.S_ABZ, self.S_Name, self.S_Grupp,
                            self.KZ, self.Vmin, self.Vmax, self.V_Nacht, self.V_TNacht]

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
            self.zuluft_max,
            self.zuluft_nacht,
            self.zuluft_Tnacht,
            self.abluft_min_ohne,
            self.abluft_min_mit,
            self.abluft_max,
            self.abluft_nacht,
            self.abluft_Tnacht,
            self.abluft_labor,
            self.abluft_24h,
            self.S_Name,
            self.S_ABZ,
            self.S_Grupp,
            self.KZ,
            self.Vmin,
            self.Vmax,
            self.V_Nacht,
            self.V_TNacht
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
al_Liste = sorted(al_Liste,key=itemgetter(15,13),reverse= True)


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
        Column = [str(RAUM[0][0]) + ' / ' + str(RAUM[0][1]),'V_min','Zumin_Raum','Abmin_Raum',
                  'V_max','Zumax_Raum','Abmax_Raum','V_Nacht','ZuNacht_Raum','AbNacht_Raum',
                  'V_TiefNacht','ZuTiefNacht_Raum','AbTiefNacht_Raum']
        Raum_AUSlass.append(Column)
        i = 0
        j = 0
        k = 0
        Summe_Zul = [0,0,0,0]
        Summe_Abl = [0,0,0,0]
        Summe_Abl_24h = [0,0,0,0]
        Summe_AblUndAbl24h = [0,0,0,0]
        Summe_Rauml = [0,0,0,0]
        ZulAUS = ['','','','']
        AblAUS = ['','','','']
        Raum_Luft = ['','','','']

        for Ausl in RAUM:
            if Ausl[15] == 'Zuluft':
                i += 1
                Zul_AUS = ['Zu_' + str(i), Ausl[17], Ausl[2], Ausl[7],
                                           Ausl[18], Ausl[3], Ausl[8],
                                           Ausl[19], Ausl[4], Ausl[9],
                                           Ausl[20], Ausl[5], Ausl[10]]
                Raum_AUSlass.append(Zul_AUS)
                Summe_Zul[0] = Summe_Zul[0] + int(round(Ausl[17]))
                Summe_Zul[1] = Summe_Zul[1] + int(round(Ausl[18]))
                Summe_Zul[2] = Summe_Zul[2] + int(round(Ausl[19]))
                Summe_Zul[3] = Summe_Zul[3] + int(round(Ausl[20]))
            elif Ausl[15] == 'Abluft':
                if Ausl[13] != 'ETA_24h':
                    j += 1
                    Abl_AUS = ['Ab_' + str(j), Ausl[17], Ausl[2], Ausl[7],
                                               Ausl[18], Ausl[3], Ausl[8],
                                               Ausl[19], Ausl[4], Ausl[9],
                                               Ausl[20], Ausl[5], Ausl[10]]
                    Raum_AUSlass.append(Abl_AUS)
                    Summe_Abl[0] = Summe_Abl[0] + int(round(Ausl[17]))
                    Summe_Abl[1] = Summe_Abl[1] + int(round(Ausl[18]))
                    Summe_Abl[2] = Summe_Abl[2] + int(round(Ausl[19]))
                    Summe_Abl[3] = Summe_Abl[3] + int(round(Ausl[20]))
                else:
                    k += 1
                    Abl24h_AUS = ['Ab24h_' + str(k), Ausl[17], Ausl[2], Ausl[7],
                                                     Ausl[18], Ausl[3], Ausl[8],
                                                     Ausl[19], Ausl[4], Ausl[9],
                                                     Ausl[20], Ausl[5], Ausl[10]]
                    Raum_AUSlass.append(Abl24h_AUS)
                    Summe_Abl_24h[0] = Summe_Abl_24h[0] + int(round(Ausl[17]))
                    Summe_Abl_24h[1] = Summe_Abl_24h[1] + int(round(Ausl[18]))
                    Summe_Abl_24h[2] = Summe_Abl_24h[2] + int(round(Ausl[19]))
                    Summe_Abl_24h[3] = Summe_Abl_24h[3] + int(round(Ausl[20]))
        for l in range(len(Summe_Abl)):
            Summe_AblUndAbl24h[l] =  Summe_Abl[l] + Summe_Abl_24h[l]
            Summe_Rauml[l] = Summe_Zul[l] - Summe_AblUndAbl24h[l]

            if Summe_Zul[l] == int(RAUM[0][2+l]):
                ZulAUS[l] = 'OK'
            else:
                ZulAUS[l] = 'FALSE'

            if Summe_AblUndAbl24h[l] == int(RAUM[0][7+l]):
                AblAUS[l] = 'OK'
            else:
                AblAUS[l] = 'FALSE'

            if Summe_Rauml[l] == 0 and ZulAUS[l] == 'OK' and AblAUS[l] == 'OK':
                Raum_Luft[l] = 'OK'
            else:
                Raum_Luft[l] = 'FALSE'





        Summe_ZUL = ['Summe_Zu', Summe_Zul[0], ZulAUS[0], '',
                                 Summe_Zul[1], ZulAUS[1], '',
                                 Summe_Zul[2], ZulAUS[2], '',
                                 Summe_Zul[3], ZulAUS[3], '']
        Summe_ABL = ['Summe_Ab', Summe_Abl[0],'','',
                                 Summe_Abl[1],'','',
                                 Summe_Abl[2],'','',
                                 Summe_Abl[3],'','']
        Summe_ABL24H = ['Summe_24h', Summe_Abl_24h[0], '','',
                                     Summe_Abl_24h[1], '','',
                                     Summe_Abl_24h[2], '','',
                                     Summe_Abl_24h[3], '','']
        Summe_ABLUND24H = ['Summe_AbUndAb24h', Summe_AblUndAbl24h[0], ' ', AblAUS[0],
                                               Summe_AblUndAbl24h[1], ' ', AblAUS[1],
                                               Summe_AblUndAbl24h[2], ' ', AblAUS[2],
                                               Summe_AblUndAbl24h[3], ' ', AblAUS[3]]
        Summe_RaumL = ['Summe_Raum', Summe_Rauml[0], Raum_Luft[0], '',
                                     Summe_Rauml[1], Raum_Luft[1], '',
                                     Summe_Rauml[2], Raum_Luft[2], '',
                                     Summe_Rauml[3], Raum_Luft[3], '']

        Raum_AUSlass.insert(i+1,Summe_ZUL)
        Raum_AUSlass.insert(i+j+2,Summe_ABL)
        Raum_AUSlass.append(Summe_ABL24H)
        Raum_AUSlass.append(Summe_ABLUND24H)
        Raum_AUSlass.append(Summe_RaumL)
    Raum_AUSlass_Liste.append(Raum_AUSlass)

output.print_table(
    table_data=table_data,
    title="Luftmengen_Unsortiert",
    columns=['Nummer', 'Name', 'Zuluftmin_Raum',  'Zuluftmax_Raum', 'ZuluftNacht_Raum',
             'ZuluftTiefNacht_Raum', 'Abluftmin_Raum','Abluftmin_Raum(incl. Summe24h)',
             'Abluftmax_Raum','AbluftNacht_Raum',
             'AbluftTNacht_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name','Sys_Abkürzung',
             'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
             'V_Nacht_Auslass', 'V_TNacht_Auslass']
)


output.print_table(
    table_data = AL_Liste,
    title = "Luftmengen_Sortiert(nach Raumnummer, Systemklassifizierung, Systemabkürzung)",
    columns=['Nummer', 'Name', 'Zuluftmin_Raum',  'Zuluftmax_Raum', 'ZuluftNacht_Raum',
             'ZuluftTiefNacht_Raum', 'Abluftmin_Raum','Abluftmin_Raum(incl. Summe24h)',
             'Abluftmax_Raum','AbluftNacht_Raum',
             'AbluftTNacht_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name','Sys_Abkürzung',
             'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
             'V_Nacht_Auslass', 'V_TNacht_Auslass']
             )


for RAL in Raum_AUSlass_Liste:
    Co = RAL[0]
    table = RAL[:]
    del(table[0])
    output.print_table(
        table_data = table,
#        title = "Luftmengen",
        columns = Co
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
                columns = ['Nummer', 'Name', 'Zuluftmin_Raum',  'Zuluftmax_Raum', 'ZuluftNacht_Raum',
                         'ZuluftTiefNacht_Raum', 'Abluftmin_Raum','Abluftmin_Raum(incl. Summe24h)',
                         'Abluftmax_Raum','AbluftNacht_Raum',
                         'AbluftTNacht_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name','Sys_Abkürzung',
                         'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
                         'V_Nacht_Auslass', 'V_TNacht_Auslass']
            )
        else:
            logger.error("Die Raumnummer ist falsch. Bitte erneut eingeben.")

        NAME = str(al_liste[0][0]) + ' / ' + str(al_liste[0][1])
        logger.info('{}'.format(NAME))

        for RAum in Raum_AUSlass_Liste:
            if RAum[0][0] == NAME:
                print('True')
                CO = RAum[0]
                Table = RAum[:]
                del(Table[0])
                output.print_table(
                    table_data = Table,
            #        title = "Luftmengen",
                    columns = CO
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
