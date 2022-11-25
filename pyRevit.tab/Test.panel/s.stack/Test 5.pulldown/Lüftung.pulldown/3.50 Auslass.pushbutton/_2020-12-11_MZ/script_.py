# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
from operator import itemgetter



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
DuctTerminal_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctTerminal)\
    .WhereElementIsNotElementType()
ducts = DuctTerminal_collector.ToElementIds()

logger.info("{} Luftauslässe ausgewählt".format(len(ducts)))

if not ducts:
    logger.error("Keine Luftauslässe in aktueller Ansicht gefunden")
    script.exit()

# MEP Räume
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

# Werte ermitteln
def MEP_Raum(Familie,Familie_Id):
    Raum = []
    Title = '{value}/{max_value} Raumdaten Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))
            param = ['Nummer', 'Name', 'IGF_RLT_ZuluftminRaum',
                     'IGF_RLT_ZuluftmaxRaum', 'IGF_RLT_ZuluftNachtRaum',
                     'IGF_RLT_ZuluftTieferNachtRaum', 'IGF_RLT_AbluftminRaumGes',
                     'IGF_RLT_AbluftminRaumL24h', 'IGF_RLT_AbluftminRaum',
                     'IGF_RLT_AbluftmaxRaum', 'IGF_RLT_AbluftNachtRaum',
                     'IGF_RLT_AbluftTieferNachtRaum','IGF_RLT_ÜberströmungRaum',
                     'IGF_RLT_ÜberstromVerteilung']
            raum = []
            for i in param:
                Werte = get_value(Item.LookupParameter(i))
                raum.append(Werte)
            Raum.append(raum)
    Raum.sort()
    output.print_table(
        table_data = Raum,
        title = "MEP-Räume",
        columns=['Nummer', 'Name', 'ZULmin_Raum',  'ZULmax_Raum', 'ZULNacht_Raum',
                 'ZUL-TNacht_Raum', 'ABLges_Raum', 'ABL24h_Raum', 'ABLmin_Raum',
                 'ABL-TNacht_Raum', 'ABLmax_Raum', 'ABLNacht_Raum',
                 'Überströmung_Raum', 'Überströmung_Verteilung']
                 )
    return Raum

def Auslass(Familie,Familie_Id):
    Auslass = []
    RaumNr = []
    Title = '{value}/{max_value} Luftauslässedaten Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))
            param = ['Systemabkürzung', 'Systemname',
                     'Systemklassifizierung', 'IGF_RLT_Volumenstrom',
                     'IGF_RLT_AuslassVolumenstromMin',
                     'IGF_RLT_AuslassVolumenstromMax',
                     'IGF_RLT_AuslassVolumenstromNacht',
                     'IGF_RLT_AuslassVolumenstromTiefeNacht',]
            auslass = []
            Nummer = get_value(Item.LookupParameter('IGF_X-Einbauort'))
            Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) \
                      .AsValueString()
            if Nummer:
                auslass.append(Nummer)
                auslass.append(Typ)
                RaumNr.append(Nummer)
            for i in param:
                Werte = get_value(Item.LookupParameter(i))
                auslass.append(Werte)
            Auslass.append(auslass)
    Auslass.sort()
    output.print_table(
        table_data = Auslass,
        title = "Luftdurclässe",
        columns=['Raumnummer', 'Systemtyp','Systemabkürzung', 'Systemname',
                 'Systemklassifizierung', 'Volumenstrom', 'Vol_min', 'Vol_max',
                 'Vol_Nacht', 'Vol_Tiefnacht']
                 )

    RaumNr = set(RaumNr)
    RaumNr = list(RaumNr)
    RaumNr.sort()
    return Auslass, RaumNr

def Raum_Auslass(Raum_NR,Auslaesseliste,Raumliste):
    All_Raum = []
    N_Auslaesseliste = sorted(Auslaesseliste,key=itemgetter(2,4),reverse= True)
    for RN in Raum_NR:
        Raum = []
        for Al in N_Auslaesseliste:
            if Al[0] == RN:
                Raum.append(Al)
        for Rm in Raumliste:
            if Rm[0] == RN:
                Raum.append(Rm)
        All_Raum.append(Raum)
    return All_Raum

Raum = MEP_Raum(spaces_collector,spaces)
Luftdurchlass, RaumNr = Auslass(DuctTerminal_collector,ducts)
Raum_Durchlass = Raum_Auslass(RaumNr,Luftdurchlass,Raum)


#########################################################################








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
    table_data = AL_Liste,
    title = "Luftmengen",
    columns=['Nummer', 'Name', 'Zuluftmin_Raum',  'Zuluftmax_Raum', 'ZuluftNacht_Raum',
             'ZuluftTiefNacht_Raum', 'Abluftmin_Raum','Abluftmin_Raum(incl. Summe24h)',
             'Abluftmax_Raum','AbluftNacht_Raum',
             'AbluftTNacht_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name','Sys_Abkürzung',
             'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
             'V_Nacht_Auslass', 'V_TNacht_Auslass']
             )

typ =[]
familie = []
kuz = []
for i in AL_Liste:
    typ.append(i[])

for RAL in Raum_AUSlass_Liste:
    Co = RAL[0]
    table = RAL[:]
    del(table[0])
    output.print_table(
        table_data = table,
#        title = "Luftmengen",
        columns = Co
    )






# ops = ['Raumnummer', 'Systemname', 'Systemabkürzung', 'Schließen']
# c = forms.CommandSwitchWindow.show(ops,message = 'Filter nach')
# if c == 'Raumnummer':
#     x = rpw.ui.forms.TextInput('Raumnummer: ', default = "0")
#     while x != '0':
#         al_liste = []
#         logger.info("Raumnummer ist:{}".format(x))
#         for Al in AL_Liste:
#             if Al[0] == x:
#                 al_liste.append(Al)
#
#         if any(al_Liste):
#             output.print_table(
#                 table_data = al_liste,
#                 title = "Luftmengen",
#                 columns = ['Nummer', 'Name', 'Zuluftmin_Raum',  'Zuluftmax_Raum', 'ZuluftNacht_Raum',
#                          'ZuluftTiefNacht_Raum', 'Abluftmin_Raum','Abluftmin_Raum(incl. Summe24h)',
#                          'Abluftmax_Raum','AbluftNacht_Raum',
#                          'AbluftTNacht_Raum', 'Abluft_Labor','Abluft_24h','Sys_Name','Sys_Abkürzung',
#                          'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass',
#                          'V_Nacht_Auslass', 'V_TNacht_Auslass']
#             )
#         else:
#             logger.error("Die Raumnummer ist falsch. Bitte erneut eingeben.")
#
#         NAME = str(al_liste[0][0]) + ' / ' + str(al_liste[0][1])
#         logger.info('{}'.format(NAME))
#
#         for RAum in Raum_AUSlass_Liste:
#             if RAum[0][0] == NAME:
#                 print('True')
#                 CO = RAum[0]
#                 Table = RAum[:]
#                 del(Table[0])
#                 output.print_table(
#                     table_data = Table,
#             #        title = "Luftmengen",
#                     columns = CO
#                 )
#
#
#         x = rpw.ui.forms.TextInput('Raumnummer: ', default = "0")
# elif c == 'Systemname':
#     x = rpw.ui.forms.TextInput('Systemname: ', default = "0")
#     while x != '0':
#         al_liste = []
#         logger.info("Systemname ist:{}".format(x))
#         for Al in AL_Liste:
#             if Al[8] == x:
#                 al_liste.append(Al)
#
#         if any(al_liste):
#             output.print_table(
#                 table_data = al_liste,
#                 title = "Luftmengen",
#                 columns = ['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
#                          'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass', 'V_Nacht_Auslass', ]
#             )
#         else:
#             logger.error("Der Systemname ist falsch. Bitte erneut eingeben.")
#
#         x = rpw.ui.forms.TextInput('Systemname: ', default = "0")
#
#
#
# elif c == 'Systemabkürzung':
#     x = rpw.ui.forms.TextInput('Systemabkürzung: ', default = "0")
#     while x != '0':
#         al_liste = []
#         logger.info("Systemabkürzung ist:{}".format(x))
#         for Al in AL_Liste:
#             if Al[9] == x:
#                 al_liste.append(Al)
#
#         if any(al_liste):
#             output.print_table(
#                 table_data = al_liste,
#                 title = "Luftmengen",
#                 columns = ['Nummer', 'Name', 'Zuluft_Raum', 'Abluft_Raum', 'Ts_Zuluft', 'TS_Abluft', 'Abluft_Labor','Abluft_24h','Sys_Name',
#                          'Sys_Abkürzung', 'Sys_Klassifizierung', 'Kennzeichen', 'Vmin_Auslass', 'Vmax_Auslass', 'V_Nacht_Auslass', ]
#             )
#
#         else:
#             logger.error("Die Systemabkürzung ist falsch. Bitte erneut eingeben.")
#
#         x = rpw.ui.forms.TextInput('Systemabkürzung: ', default = "0")
#
# elif c == 'Schließen':
#     pass



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
