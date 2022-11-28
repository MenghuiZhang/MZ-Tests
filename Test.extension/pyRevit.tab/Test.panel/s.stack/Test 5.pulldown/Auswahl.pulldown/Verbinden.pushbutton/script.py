# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit,DB,UI
from pyrevit import script,forms
from IGF_log import getlog


__title__ = "2. Verbinden"
__doc__ = """
Verbinden

Parameter: 
IGF_X_AnschlussId_Verbinder,
IGF_X_AnschlussId_Leitung,
IGF_X_UniqueId_Leitung,
IGF_X_Hauptanlage

[2021.09.27]
Version: 1.1"""
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

selection = uidoc.Selection.GetElementIds()
logger.info('{} Baueile ausgewählt.'.format(len(selection)))
if len(selection) == 0:
    UI.TaskDialog.Show(__title__, 'Kein Bauteil ausgewählt')
    script.exit()

def SystemDaten(system):
    dict_param = {}
    # Daten von System werden gespeicht.
    for Param in system.Parameters:       
        if not Param.IsReadOnly:
            # if Param.Definition.Name == 'Systemname':
            #     continue
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
    return dict_param

def Datenschreiben(daten,system_neu):
    system_neu_name = system_neu.LookupParameter('Systemname').AsString()
    for param_key in daten.keys():
        try:
            system_neu.LookupParameter(param_key).Set(daten[param_key])
        except:
            logger.error('Fehler beim Daten-Schreiben. Parameter: {}, Werte: {}, System: {}'.format(param_key,daten[param_key],system_neu_name))

with forms.ProgressBar(title="{value}/{max_value} Bauteile ausgewählt", cancellable=True, step=1) as pb:
    for n, elemid in enumerate(selection):
        t = DB.Transaction(doc)
        
        if pb.cancelled:
            t.RollBack()
            script.exit()
        pb.update_progress(n + 1, len(selection))

        elem = doc.GetElement(elemid)
        # Prüfen
        conn_l_nr = None
        conn_v_nr = None
        elem_id = None

        if elem.LookupParameter('IGF_X_AnschlussId_Leitung'):
            conn_l_nr = elem.LookupParameter('IGF_X_AnschlussId_Leitung').AsInteger()
        if elem.LookupParameter('IGF_X_AnschlussId_Verbinder'):
            conn_v_nr = elem.LookupParameter('IGF_X_AnschlussId_Verbinder').AsInteger()
        if elem.LookupParameter('IGF_X_UniqueId_Leitung'):
            elem_id = elem.LookupParameter('IGF_X_UniqueId_Leitung').AsString()
        if not (conn_l_nr and conn_v_nr and elem_id):
            logger.error('Parameter nicht Vollständig, ElementId: {}'.format(elemid.ToString()))
            UI.TaskDialog.Show(__title__,'Parameter nicht Vollständig, ElementId: {}'.format(elemid.ToString()))
            continue
        
        try:
            leitung = doc.GetElement(elem_id)
            sys_leitung = leitung.MEPSystem
            sysname_leitung = sys_leitung.LookupParameter('Systemname').AsString()
            hauptanlage_leitung = sys_leitung.LookupParameter('IGF_X_Hauptanlage').AsInteger()
            for conn in elem.MEPModel.ConnectorManager.Connectors: 
                if conn.Id == conn_v_nr: conn_v = conn
            for conn in leitung.ConnectorManager.Connectors: 
                if conn.Id == conn_l_nr: conn_l = conn
            sys_verbinder = conn_v.MEPSystem
            sysname_verbinder = sys_verbinder.LookupParameter('Systemname').AsString()
            hauptanlage_verbinder = sys_verbinder.LookupParameter('IGF_X_Hauptanlage').AsInteger()
            
            if sysname_verbinder != sysname_leitung:
                t.Start('Verbinden:{} -- {}'.format(sysname_verbinder,sysname_leitung))
                if hauptanlage_leitung != 0:
                    daten = SystemDaten(sys_leitung)
                    conn_v.ConnectTo(conn_l)
                    system_neu = leitung.MEPSystem
                    doc.Regenerate()
                    t.Commit()
                    system_neu_name = system_neu.LookupParameter('Systemname').AsString()
                    t.Start('Daten schreiben, '+system_neu_name)
                    Datenschreiben(daten,system_neu)
                    t.Commit()
                    continue
                
                if hauptanlage_verbinder != 0:
                    daten = SystemDaten(sys_verbinder)
                    conn_v.ConnectTo(conn_l)
                    doc.Regenerate()
                    t.Commit()
                    system_neu = leitung.MEPSystem
                    system_neu_name = system_neu.LookupParameter('Systemname').AsString()
                    t.Start('Daten schreiben, '+system_neu_name)
                    Datenschreiben(daten,system_neu)
                    t.Commit()
                    continue
                if sysname_verbinder.find('(') != -1 and sysname_verbinder.find(')') != -1:
                    daten = SystemDaten(sys_verbinder)
                    conn_v.ConnectTo(conn_l)
                    doc.Regenerate()
                    t.Commit()
                    system_neu = leitung.MEPSystem
                    system_neu_name = system_neu.LookupParameter('Systemname').AsString()
                    t.Start('Daten schreiben, '+system_neu_name)
                    Datenschreiben(daten,system_neu)
                    t.Commit()
                    continue
                if sysname_leitung.find('(') != -1 and sysname_leitung.find(')') != -1:
                    daten = SystemDaten(sys_leitung)
                    conn_v.ConnectTo(conn_l)
                    doc.Regenerate()
                    t.Commit()
                    system_neu = leitung.MEPSystem
                    system_neu_name = system_neu.LookupParameter('Systemname').AsString()
                    t.Start('Daten schreiben, '+system_neu_name)
                    Datenschreiben(daten,system_neu)
                    t.Commit()
                    continue
                else:
                    conn_v.ConnectTo(conn_l)
                    doc.Regenerate()
                    t.Commit()
            else:
                pass

        except Exception as e:
            logger.error(e)
            logger.error('ElementId: {}'.format(elemid.ToString()))
            if t.HasStarted():
                t.RollBack()    