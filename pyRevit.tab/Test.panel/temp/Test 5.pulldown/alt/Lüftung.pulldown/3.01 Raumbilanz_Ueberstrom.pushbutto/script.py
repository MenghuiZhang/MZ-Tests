# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr

start = time.time()

__title__ = "3.01 Überströmung_Raum"
__doc__ = """Raumbilanz_Überstrom (MEP-Räume)"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
from pyIGF_logInfo import getlog
getlog(__title__)

# Baugruppen
Baugruppen_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.AssemblyInstance)) \
    .WhereElementIsNotElementType()

Baugruppen = Baugruppen_collector.ToElementIds()

logger.info("{} Baugruppen ausgewählt".format(len(Baugruppen)))

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

phase = list(doc.Phases)[-1]
# Werte ermitteln
def Werte_ermitteln(Familie,Familie_Id):
    Raum = []
    Title = '{value}/{max_value} Daten Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0

        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))

            ElementIds = Item.GetMemberIds()
            i = 0
            for id in ElementIds:
                System_Typ = doc.GetElement(id) \
                .get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) \
                .AsValueString()
                if System_Typ == '31_Überstromluft':
                    i += 1
            if i > 0:
                Eingangsraum = get_value(Item.LookupParameter('IGF_RLT_Überströmung_Eingang'))
                Ausgangsraum = get_value(Item.LookupParameter('IGF_RLT_Überströmung_Ausgang'))
                Volumenstrom = get_value(Item.LookupParameter('IGF_RLT_Überströmung'))

                Raum.append([Ausgangsraum,Eingangsraum,Volumenstrom])

    return Raum


def Liste_Bearbeiten(input_list):
    Eingang = []
    Ausgang = []
    Ein_Raum = []
    Aus_Raum = []
    Raum = []
    for item in input_list:
        Eingang.append([item[0],item[1],item[2]])
        Ausgang.append([item[1],item[0],item[2]])
        Ein_Raum.append(item[0])
        Aus_Raum.append(item[1])
        Raum.append(item[0])
        Raum.append(item[1])
    Ein_Raum = set(Ein_Raum)
    Ein_Raum = list(Ein_Raum)
    Aus_Raum = set(Aus_Raum)
    Aus_Raum = list(Aus_Raum)
    Raum = set(Raum)
    Raum = list(Raum)
    return Eingang,Ausgang,Ein_Raum,Aus_Raum,Raum

def Raum_Daten(Eingang,Ausgang,Ein_Raum,Aus_Raum,Raum):
    Eingangsdaten = []
    Ausgangsdaten = []
    Raumdaten = []
    Raum_Text = []
    for item in Ein_Raum:
        raum = [item]
        for daten in Eingang:
            if daten[0] == item:
                raum.append([daten[1],daten[2]])
        Eingangsdaten.append(raum)
    for item in Aus_Raum:
        raum = [item]
        for daten in Ausgang:
            if daten[0] == item:
                raum.append([daten[1],daten[2]])
        Ausgangsdaten.append(raum)
    for item in Raum:
        raum_ein = []
        raum_aus = []
        for ein in Eingangsdaten:
            if ein[0] == item:
                raum_ein = ein[1:]

        for aus in Ausgangsdaten:
            if aus[0] == item:
                raum_aus = aus[1:]
        Raumdaten.append([item,raum_ein,raum_aus])

    for item in Raumdaten:
        if len(item[1]) > 0:
            output.print_table(
                table_data=item[1],
                title=item[0],
                columns=['Aus Raum', 'Volumen']
            )
            if len(item[2]) > 0:
                output.print_table(
                    table_data=item[2],
                    columns=['In Raum', 'Volumen'])
        else:
            if len(item[2]) > 0:
                output.print_table(
                    table_data=item[2],
                    title=item[0],
                    columns=['In Raum', 'Volumen'])

    for item in Raumdaten:
        text_raum = ''
        text_ein = ''
        text_aus = ''
        if len(item[1]) > 0:
            text_ein = 'Aus: [m3/h]--'
            for i in item[1]:
                text_ein = text_ein + i[0] + ': ' + str(int(i[1])) + ', '
        if len(item[2]) > 0:
            text_aus = 'In: [m3/h]--'
            for i in item[2]:
                text_aus = text_aus + i[0] + ': ' + str(int(i[1])) + ', '
        text_raum = text_ein + text_aus
        Raum_Text.append([item[0],text_raum])

    return Eingangsdaten, Ausgangsdaten, Raumdaten, Raum_Text

def Luftsumme(Raum_Daten):
    out = []
    for item in Raum_Daten:
        ein = 0
        aus = 0
        ges = 0
        if len(item[1]) > 0:
            for i in item[1]:
                ein = ein + i[1]
        if len(item[2]) > 0:
            for i in item[2]:
                aus = aus + i[1]
        ges = ein - aus
        out.append([item[0],ein,aus,ges])

    return out

def gesamt_list(luftsumme,luftverteilung):
    out = []
    for item in luftsumme:
        for item2 in luftverteilung:
            if str(item[0]) == str(item2[0]):
                out.append([item[0],item[1],item[2],item[3],item2[1]])
    output.print_table(
        table_data=out,
        title='Überstrom',
        columns=['Raum', 'Überstrom_In','Überstrom_Aus','ÜberstromSumme','Luftverteilung'])

    return out

def Datenschreiben(Raumlist,Familie):
    Title = '{value}/{max_value} Daten schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Daten schreiben")
        t.Start()

        for item in Raumlist:

            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Raumlist))

            Item = Familie[item[0]]

            Item.LookupParameter('IGF_RLT_ÜberströmungSummeIn').SetValueString(str(item[1]))
            Item.LookupParameter('IGF_RLT_ÜberströmungSummeAus').SetValueString(str(item[2]))
            Item.LookupParameter('IGF_RLT_ÜberströmungRaum').SetValueString(str(item[3]))
            Item.LookupParameter('IGF_RLT_ÜberstromVerteilung').Set(str(item[4]))


        t.Commit()


Kanaldaten = Werte_ermitteln(Baugruppen_collector,Baugruppen)
Eingang_liste,Ausgang_liste,Ein_Raum_liste,Aus_Raum_liste,Raum_liste = Liste_Bearbeiten(Kanaldaten)
Eingangsdaten, Ausgangsdaten, Raumdaten, Raum_Text = Raum_Daten(Eingang_liste,Ausgang_liste,Ein_Raum_liste,Aus_Raum_liste,Raum_liste)
Luftsumme_liste = Luftsumme(Raumdaten)
gesamt_list = gesamt_list(Luftsumme_liste,Raum_Text)
Raumdic = {}
for el in spaces_collector:
    Raumnummer = get_value(el.LookupParameter('Nummer'))
    Raumdic[Raumnummer] = el



if forms.alert('Daten schreiben?', ok=False, yes=True, no=True):
    Datenschreiben(gesamt_list, Raumdic)



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
