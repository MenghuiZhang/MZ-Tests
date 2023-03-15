# coding: utf8
from rpw import revit,DB,UI
from pyrevit import script


__title__ = "3. Verbinden"
__doc__ = """...

[2021.09.23]
Version: 1.0"""
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
t = DB.Transaction(doc)
t.Start('System verbinden')
leitung = doc.GetElement(elem_id)

# Systeme werden verbunden
try:
    for conn in elem.MEPModel.ConnectorManager.Connectors: 
        if conn.Id == conn_v_nr: conn_v = conn
    for conn in leitung.ConnectorManager.Connectors: 
        if conn.Id == conn_l_nr: conn_l = conn
    conn_v.ConnectTo(conn_l)
    doc.Regenerate()
except Exception as e:
    logger.error(e)
    logger.error('Fehler beim Verbinden der Systeme')
    t.RollBack()
    script.exit()
doc.Regenerate()
t.Commit()
# Umbenennen
t.Start('System umbenennen')
try:
    system = leitung.MEPSystem
    sys_name = system.LookupParameter('Systemname').AsString()
    system.LookupParameter('Systemname').Set(str(sys_name[:-4]))
    doc.Regenerate()
except:
    logger.error('Fehler beim Umbenennen des Systems')

t.Commit()

