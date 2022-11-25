# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
from operator import itemgetter



start = time.time()


__title__ = "3.5 Luftauslass"
__doc__ = """Informationen der Luftauslaesse zeigen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView
from pyIGF_logInfo import getlog
getlog(__title__)

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
# Werte ermitteln
def MEP_Raum(Familie,Familie_Id):
    Raum = []
    Schacht = []
    Title = '{value}/{max_value} Raumdaten Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=10) as pb:
        n = 0
        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))

            param = ['Nummer', 'Name', 'IGF_RLT_ZuluftminRaum',
                     'IGF_RLT_ZuluftmaxRaum', 'IGF_RLT_ZuluftNachtRaum',
                     'IGF_RLT_ZuluftTieferNachtRaum', 'IGF_RLT_AbluftminRaumOhneÜber',
                     'IGF_RLT_AbluftminRaum24h', 'IGF_RLT_AbluftminRaum',
                     'IGF_RLT_AbluftmaxRaum', 'IGF_RLT_AbluftNachtRaum',
                     'IGF_RLT_AbluftTieferNachtRaum','IGF_RLT_ÜberströmungRaum',
                     'IGF_RLT_ÜberstromVerteilung']
            para = ['Nummer', 'Name','TGA_RLT_SchachtZuluftMenge',
                    'TGA_RLT_SchachtAbluftMenge','TGA_RLT_Schacht24hAbluftMenge']
            if get_value(Item.LookupParameter('TGA_RLT_InstallationsSchacht')):
                schacht = []
                for i in para:
                    Werte = get_value(Item.LookupParameter(i))
                    schacht.append(Werte)
                Schacht.append(schacht)
            else:
                raum = []
                for i in param:
                    Werte = get_value(Item.LookupParameter(i))
                    raum.append(Werte)
                Raum.append(raum)

    Raum.sort()
    Schacht.sort()
    output.print_table(
        table_data = Raum,
        title = "MEP-Räume",
        columns=['Nummer', 'Name', 'ZULmin_Raum',  'ZULmax_Raum', 'ZULNacht_Raum',
                 'ZUL-TNacht_Raum', 'ABmin+Ab24h', 'ABL24h_Raum', 'ABLmin_Raum',
                 'ABLmax_Raum', 'ABLNacht_Raum', 'ABL-TNacht_Raum',
                 'Überströmung_Raum', 'Überströmung_Verteilung']
                 )
    output.print_table(
        table_data = Schacht,
        title = "Schächte",
        columns=['Nummer', 'Name', 'Zuluft',  'Abluft', 'Abluft_24h',]
                 )
    return Raum,Schacht

def Auslass(Familie,Familie_Id):
    Auslass = []
    RaumNr = []
    Title = '{value}/{max_value} Luftauslässedaten Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=10) as pb:
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
            Nummer = get_value(Item.LookupParameter('IGF_X_Einbauort'))
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
        title = "Luftdurchlässe",
        columns=['Raumnummer', 'Systemtyp','Systemabkürzung', 'Systemname',
                 'Systemklassifizierung', 'Volumenstrom', 'Vol_min', 'Vol_max',
                 'Vol_Nacht', 'Vol_Tiefnacht']
                 )

    RaumNr = set(RaumNr)
    RaumNr = list(RaumNr)
    RaumNr.sort()
    return Auslass, RaumNr

def Raum_Auslass(Auslaesseliste,Raumliste):
    All_Raum = []
    N_Auslaesseliste = sorted(Auslaesseliste,key=itemgetter(2,4),reverse= True)
    for Rm in Raumliste:
        Raum = []
        for Al in N_Auslaesseliste:
            if Al[0] == Rm[0]:
                Raum.append(Al)
        if len(Raum) > 0:
            Raum.append(Rm)
            All_Raum.append(Raum)

    return All_Raum

def Auslass_Sort(Auslass_Liste):
    def summe(liste,kz):
        Summe = [kz,0,0,0,0,0]
        if len(liste) > 0:
            for ele in liste:
                for i in range(1,len(Summe)):
                    Summe[i] = Summe[i] + ele[i]
        return Summe

    def ab_summe(In_ab,In_ab24h):
        Summe = ['Ab+Ab24h',0,0,0,0,0]
        if len(In_ab) > 0:
            if len(In_ab24h) > 0:
                for i in range(1,len(Summe)):
                    Summe[i] = In_ab[i] + In_ab24h[i]
            else:
                for i in range(1,len(Summe)):
                    Summe[i] = In_ab[i]
        else:
            if len(In_ab24h) > 0:
                for i in range(1,len(Summe)):
                    Summe[i] = In_ab24h[i]
            else:
                pass
        return Summe

    def Bilanz_Pru(gegeben,berechneten):
        ergebnis = ['Prüfung','--','--','--','--','--']
        for i in range(1,len(gegeben)):
            if gegeben[i] != '--':
                if gegeben[i] == berechneten[i]:
                    ergebnis[i] = 'OK'
                else:
                    ergebnis[i] = 'FEHLER!'
            else:
                pass

        return ergebnis

    def Druck_Pru(zuluft,abluft,ueberstrom):
        ergebnis = ['Druckstufe','--','--','--','--','--']
        dr = zuluft[1] - abluft[1] + ueberstrom[1]
        if dr > 0:
            ergebnis[1] = 'Überdruck'
        elif dr < 0:
            ergebnis[1] = 'Unterdruck'
        else:
            ergebnis[1] = 'Balance'
        return ergebnis

    n_Raum_Auslass = []
    for auslass in Auslass_Liste:
        Zu,Ab,Ab_24h,Ueber = [],[],[],[]
        Zu_Nr,Ab_Nr,Ab_24h_Nr,Ueber_Nr = 0,0,0,0
        Raum_Nr = auslass[0][0]
        for i in range(0,len(auslass)-1):
            ele = auslass[i][1:]
            if ele[0] != '31_Überstromluft':
                if ele[1] == 'ETA_24h':
                    Ab_24h_Nr += 1
                    text = 'Ab24h-' + str(Ab_24h_Nr)
                    neu_ele = ele[4:9]
                    neu_ele.insert(0,text)
                    Ab_24h.append(neu_ele)
                else:
                    if ele[3] == 'Zuluft':
                        Zu_Nr += 1
                        text = 'Zu-' + str(Zu_Nr)
                        neu_ele = ele[4:9]
                        neu_ele.insert(0,text)
                        Zu.append(neu_ele)
                    else:
                        Ab_Nr += 1
                        text = 'Ab-' + str(Ab_Nr)
                        neu_ele = ele[4:9]
                        neu_ele.insert(0,text)
                        Ab.append(neu_ele)
            else:
                Ueber_Nr += 1
                text = 'Über-' + str(Ueber_Nr)
                if ele[3] == 'Zuluft':
                    neu_ele = ele[4:9]
                    neu_ele.insert(0,text)
                    Ueber.append(neu_ele)
                else:
                    neu_ele = [text,0-ele[4],0-ele[5],0-ele[6],0-ele[7],0-ele[8]]
                    Ueber.append(neu_ele)

        Summe_Zu = summe(Zu,'Zuluft_Summe')
        Summe_Ab = summe(Ab,'Abluft_Summe')
        Summe_Ab24h = summe(Ab_24h,'Ab24h_Summe')
        Summe_AbUAb24h = ab_summe(Summe_Ab,Summe_Ab24h)
        Summe_Ueber = summe(Ueber,'Über_Summe')
        Raum = auslass[len(auslass)-1]
        Zu_AusRaum = ['Zuluft aus Raum',Raum[2],Raum[2],Raum[3],Raum[4],Raum[5]]
        Ab_AusRaum = ['Abluft aus Raum',Raum[8],Raum[8],Raum[9],Raum[10],Raum[11]]
        Ab24h_AusRaum = ['Abluft24h aus Raum',Raum[7],Raum[7],'--','--','--']
        AbUnd24h_AusRaum = ['Ab+Ab24h aus Raum',Raum[6],Raum[6],'--','--','--']
        Ueber_AusRaum = ['Überströmung aus Raum',Raum[12],Raum[12],'--','--','--']
        zu_pru = Bilanz_Pru(Zu_AusRaum,Summe_Zu)
        ab_pru = Bilanz_Pru(Ab_AusRaum,Summe_Ab)
        ab24h_pru = Bilanz_Pru(Ab24h_AusRaum,Summe_Ab24h)
        abU24h_pru = Bilanz_Pru(AbUnd24h_AusRaum,Summe_AbUAb24h)
        ueber_pru = Bilanz_Pru(Ueber_AusRaum,Summe_Ueber)
        druck_pru = Druck_Pru(Summe_Zu,Summe_AbUAb24h,Summe_Ueber)
        Zu.append(Summe_Zu)
        Zu.append(Zu_AusRaum)
        Zu.append(zu_pru)
        Ab.append(Summe_Ab)
        Ab.append(Ab_AusRaum)
        Ab.append(ab_pru)
        Ab_24h.append(Summe_Ab24h)
        Ab_24h.append(Ab24h_AusRaum)
        Ab_24h.append(abU24h_pru)
        Ab_24h.append(Summe_AbUAb24h)
        Ab_24h.append(AbUnd24h_AusRaum)
        Ab_24h.append(abU24h_pru)
        Ueber.append(Summe_Ueber)
        Ueber.append(Ueber_AusRaum)
        Ueber.append(ueber_pru)
        Ueber.append(druck_pru)
        raum_auslass = [Raum_Nr,Zu,Ab,Ab_24h,Ueber,Raum]
        n_Raum_Auslass.append(raum_auslass)
    return n_Raum_Auslass

def final_liste(liste):
    Final = []
    for item in liste:
        final = [item[0]]
        for i in range(1,5):
            for ele in item[i]:
                if any(ele):
                    final.append(ele)
        Final.append(final)
        output.print_table(
            table_data = final[1:],
            title = final[0],
            columns=['Name', 'Volumenstrom', 'Vol_min', 'Vol_max',
                     'Vol_Nacht', 'Vol_Tiefnacht']
                     )
    return Final

Raum,Schacht = MEP_Raum(spaces_collector,spaces)
Luftdurchlass, RaumNr = Auslass(DuctTerminal_collector,ducts)
Raum_Durchlass = Raum_Auslass(Luftdurchlass,Raum)
Schacht_Durchlass = Raum_Auslass(Luftdurchlass,Schacht)
Raum_Durchlass_mitSumme = Auslass_Sort(Raum_Durchlass)
# Schahct_Durchlass_mitSumme = Auslass_Sort(Schacht_Durchlass)
Final_Liste = final_liste(Raum_Durchlass_mitSumme)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
