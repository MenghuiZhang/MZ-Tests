# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "2.70 Massenstrom 6-Wege-Ventil"
__doc__ = """Massenstrom 6-Wege-Ventil. Max. Abstand zwischen Ventil und Verbindung ist 305 mm."""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
active_view = uidoc.ActiveView

from pyIGF_logInfo import getlog
getlog(__title__)

# Rohrzubehör aus aktueller Ansicht
PipeAccessory_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_PipeAccessory)\
    .WhereElementIsNotElementType()

def get_value(param):

        value = revit.query.get_param_value(param)

        try:
            unit = param.DisplayUnitType

            value = DB.UnitUtils.ConvertFromInternalUnits(
                value,
                unit)

        except Exception as e:
            pass

        return value

Vorlauf = []
Ruecklauf = []
Verbindung = []
for rohr in PipeAccessory_collector:
    familie = rohr.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
    if familie == '_HK_IGF_481_6-Wege-Regelkugelhahn-K-VL':
        Vorlauf.append(rohr)
    elif familie == '_HK_IGF_481_6-Wege-Regelkugelhahn-K-RL':
        Ruecklauf.append(rohr)
    elif familie == 'SBI_481_G_6_Wege_Verbindung':
        Verbindung.append(rohr)

Kombi = []
for Vor in Vorlauf:
    x_vor = Vor.Location.Point.X
    y_vor = Vor.Location.Point.Y
    z_vor = Vor.Location.Point.Z
    for Ver in Verbindung:
        x_ver = Ver.Location.Point.X
        y_ver = Ver.Location.Point.Y
        z_ver = Ver.Location.Point.Z
        if (x_vor-x_ver)**2 + (y_vor-y_ver)**2 + (z_vor-z_ver)**2 < 1:
            Kombi.append([Vor,Ver])
            break

Vor_dic = {}
for el in Vorlauf:
    connectors = el.MEPModel.ConnectorManager.Connectors
    for conn in connectors:
        if not conn.IsConnected:
            continue
        connectorSet = conn.AllRefs
        iterator = connectorSet.ForwardIterator()
        while iterator.MoveNext():
            currentConnector = iterator.Current
            if currentConnector:
                vol = get_value(currentConnector.Owner.LookupParameter('Volumenstrom'))
                if vol >= 0.005:
                    Vor_dic[el] = round(vol,2)
                    break


Rueck_dic = {}
for el in Ruecklauf:
    connectors = el.MEPModel.ConnectorManager.Connectors
    for conn in connectors:
        if not conn.IsConnected:
            continue
        connectorSet = conn.AllRefs
        iterator = connectorSet.ForwardIterator()
        while iterator.MoveNext():
            currentConnector = iterator.Current
            if currentConnector:
                vol = get_value(currentConnector.Owner.LookupParameter('Volumenstrom'))
                if vol >= 0.005:
                    Rueck_dic[el] = round(vol,2)
                    break
print(Vor_dic)
print(30*'-')
print(Rueck_dic)
print(len(Vorlauf),len(Vor_dic),len(Kombi))
print(len(Rueck_dic),len(Ruecklauf))

t = Transaction(doc, "Volumenstrom schreiben")
t.Start()
with forms.ProgressBar(title='{value}/{max_value} Vorlauf schreiben',
                           cancellable=True, step=1) as pb:

    n_1 = 0
    for el in Kombi:
        if pb.cancelled:
            script.exit()
        n_1 += 1
        pb.update_progress(n_1,len(Kombi))
        vol = Vor_dic[el[0]]
        el[0].LookupParameter('Volumenstrom').SetValueString(str(vol))
        el[1].LookupParameter('SBI_Volumenstrom').SetValueString(str(vol))


with forms.ProgressBar(title='{value}/{max_value} Rücklauf schreiben',
                           cancellable=True, step=1) as pb:

    n_1 = 0
    for el in Ruecklauf:
        if pb.cancelled:
            script.exit()
        n_1 += 1
        pb.update_progress(n_1,len(Ruecklauf))
        vol = Rueck_dic[el]
        el.LookupParameter('Volumenstrom').SetValueString(str(vol))

t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
