# coding: utf8
from rpw import revit,DB,UI
from pyrevit import forms,script
from pyIGF_logInfo import getlog
from pyIGF_class import tempsystemtypclass,tempclass_id
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights

__title__ = "3. System trennen"
__doc__ = """
Es wird für jedes Netz ein neues System angelegt.

Parameter: 
ConnectorID_Verbinder,
ConnectorID_Rohre,
UniqueId_Rohre


[2021.09.23]
Version 1.0 """
__author__ = "Menghui Zhang"
__context__ = 'Selection'

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

elem = doc.GetElement(uidoc.Selection.GetElementIds()[0])

# Prüfen
conn_l_nr = None
conn_v_nr = None
elem_id = None

if elem.LookupParameter('ConnectorID_Rohre'):
    conn_l_nr = elem.LookupParameter('ConnectorID_Rohre').AsInteger()
if elem.LookupParameter('ConnectorID_Verbinder'):
    conn_v_nr = elem.LookupParameter('ConnectorID_Verbinder').AsInteger()
if elem.LookupParameter('UniqueId_Rohre'):
    elem_id = elem.LookupParameter('UniqueId_Rohre').AsString()

if not (conn_l_nr and conn_v_nr and elem_id):
    UI.TaskDialog.Show(__title__,'Parameter nicht Vollständig')
    script.exit()

# Trennen
t = DB.Transaction(doc,'System trennen')
t.Start()
leitung = doc.GetElement(elem_id)
system = leitung.MEPSystem
dict_param = {}
# Daten von System werden gespeicht.
for Param in system.Parameters:
    if Param.Definition.Name == 'Systemname':
        continue
    if not Param.IsReadOnly:
        if Param.StorageType == DB.StorageType.Double:
            value = Param.AsDouble()
            if value:
                dict_param[Param.Definition.Name] = value
        elif Param.StorageType == DB.StorageType.Integer:
            value = Param.AsInteger()
            if value:
                dict_param[Param.Definition.Name] = value
        elif Param.StorageType == DB.StorageType.String:
            value = Param.AsString()
            if value:
                dict_param[Param.Definition.Name] = value
        elif Param.StorageType == DB.StorageType.ElementId:
            value = Param.AsElementId()
            if value:
                dict_param[Param.Definition.Name] = value

# System wird getrennt
System_liste = []

try:
    System_liste = leitung.MEPSystem.DivideSystem(doc)
except:
    logger.error('Fehler beim Trennen des Systems')
    t.RollBack()
    script.exit()
doc.Regenerate()
# Daten schreiben
try:
    for el in System_liste:
        sys = doc.GetElement(el)
        sys_name = sys.LookupParameter('Systemname').AsString()
        for param_key in dict_param.keys():
            try:
                sys.LookupParameter(param_key).Set(dict_param[param_key])
            except:
                logger.error('Fehler beim Daten-Schreiben. Parameter: {}, Werte: {}, System: {}'.format(param_key,dict_param[param_key],sys_name))
except:
    logger.error('Fehler beim Datenschreiben')
    t.RollBack()
    script.exit()

t.Commit()