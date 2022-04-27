# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr

start = time.time()

__title__ = "Überströmung_Bauteil"
__doc__ = """Raumnummer und Volumen von Überstrom in Bauteile schreiben"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# Luftauslässe
Luftauslaesse_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctTerminal)\
    .WhereElementIsNotElementType()
Luftauslaesse = Luftauslaesse_collector.ToElementIds()

logger.info("{} Luftauslässe ausgewählt".format(len(Luftauslaesse)))
# Rohr
Kanal_collector = DB.FilteredElementCollector(doc) \
    .OfClass(clr.GetClrType(DB.Mechanical.Duct)) \
    .WhereElementIsNotElementType()

Kanal = Kanal_collector.ToElementIds()

logger.info("{} Rohre ausgewählt".format(len(Kanal)))

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
# Werte schreiben
def Raum_ermitteln(Familie,Familie_Id):
    Raum = []
    Title = '{value}/{max_value} Raumnummer Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=10) as pb:
        n = 0

        for Item in Familie:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))

            Raumnummer = None
            if Item.Space[phase]:
                System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
                if System_Typ == '31_Überstromluft':
                    Raumnummer = Item.Space[phase].Number
                    System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
                    Familie_Typ = Item.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                    Vol = get_value(Item.LookupParameter('IGF_RLT_Überströmung'))
                    Raum.append([System_Name,Familie_Typ,Raumnummer,Vol])

                    if Familie_Typ == 'CAx Zuluft Trassenteilstrecke RE':
                        logger.info('Raum: {}, Überströmung {}'.format(Raumnummer,Vol))
                    elif Familie_Typ == 'CAx Abluft Trassenteilstrecke RE':
                        logger.info('Raum: {}, Überströmung {}'.format(Raumnummer,0-Vol))

    return Raum

def Vol_ermitteln_Kanal(Familie,Familie_Id):
    Vol = []
    Title = '{value}/{max_value} Volumen Ermitteln'
    with forms.ProgressBar(title=Title, cancellable=True, step=100) as pb:
        n = 0

        for Item in Familie:

            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(Familie_Id))

            System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
            System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
            vol = get_value(Item.LookupParameter('IGF_RLT_Überströmung'))
            if System_Typ == '31_Überstromluft' and vol > 0:
                Vol.append([System_Name,vol])


    return Vol

def Liste(input_list_Raum, input_list_Vol):
    out_liste = []
    sys_liste = []
    for item in input_list_Raum:
        sys_liste.append(item[0])
    sys_liste = set(sys_liste)
    sys_liste = list(sys_liste)
    sys_liste.sort()
    for item in sys_liste:
        out_liste.append([item,'','',0])
    for sys in out_liste:
        for item in input_list_Raum:
            if item[0] == sys[0]:
                if item[1] == 'CAx Abluft Trassenteilstrecke RE':
                    sys[1] = item[2]
                    if item[3] > 0:
                        sys[3] = item[3]
                elif item[1] == 'CAx Zuluft Trassenteilstrecke RE':
                    sys[2] = item[2]
                    if item[3] > 0:
                        sys[3] = item[3]
    if len(input_list_Vol) > 0:
        for sys in out_liste:
            for item in input_list_Vol:
                if item[0] == sys[0] and sys[3] == 0 and item[1] > 0:
                    sys[3] = item[1]


    output.print_table(
        table_data=out_liste,
        title="System",
        columns=['Systemname', 'Eingangsraum', 'Ausgangsraum','Volumenstrom']
    )
    return out_liste


def Datenschreiben_Kanal(input_list,Familie,Familie_Id):
    Title = '{value}/{max_value} Daten schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Ein-/Ausgangsraum und Überstrom in Kanal schreiben")
        t.Start()


        for item in input_list:

            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(input_list))

            for Item in Familie:
                System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
                System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
                if System_Name == item[0] and System_Typ == '31_Überstromluft':
                    Item.LookupParameter('IGF_RLT_Überströmung_Eingang').Set(str(item[1]))
                    Item.LookupParameter('IGF_RLT_Überströmung_Ausgang').Set(str(item[2]))
                    Item.LookupParameter('IGF_RLT_Überströmung').SetValueString(str(item[3]))

        t.Commit()

def Datenschreiben_Auslass(input_list,Familie,Familie_Id):
    Title = '{value}/{max_value} Daten schreiben'
    with forms.ProgressBar(title=Title, cancellable=True, step=1) as pb:
        n = 0
        t = Transaction(doc, "Überstrom in Auslass schreiben")
        t.Start()

        for item in input_list:
            if pb.cancelled:
                script.exit()
            n += 1
            pb.update_progress(n, len(input_list))

            for Item in Familie:
                System_Typ = Item.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
                System_Name = Item.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
                if System_Name == item[0] and System_Typ == '31_Überstromluft':
                    Item.LookupParameter('IGF_RLT_Überströmung').SetValueString(str(item[3]))

        t.Commit()

Raum_Liste = Raum_ermitteln(Luftauslaesse_collector,Luftauslaesse)

Vol_Liste = Vol_ermitteln_Kanal(Kanal_collector,Kanal)
Sys_Liste = Liste(Raum_Liste,Vol_Liste)

if forms.alert('Daten schreiben?', ok=False, yes=True, no=True):
    Datenschreiben_Kanal(Sys_Liste, Kanal_collector,Kanal)
    Datenschreiben_Auslass(Sys_Liste,Luftauslaesse_collector,Luftauslaesse)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
