# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit,DB,UI
from pyrevit import script,forms
from IGF_log import getlog

__title__ = "1. Trennen"
__doc__ = """
Daten werden in folgenden drei Parameter schreiben.
Das physische Netz im System wird aufgeteilt.
Es wird für jedes Netz ein neues System angelegt.

Parameter: 
IGF_X_AnschlussId_Verbinder,
IGF_X_AnschlussId_Leitung,
IGF_X_UniqueId_Leitung,
IGF_X_Hauptanlage

[2021.09.27]
Version 1.1 """

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

class Verbinder(object):
    def __init__(self,elementid):
        self.elem_id = elementid
        self.elem = doc.GetElement(self.elem_id)
        self.connid_Leitung = ''
        self.connid_verbinder = ''
        self.elemid_Leitung = ''
        if not self.Werte_Pruefen():
            self.Daten_ermitteln()
    
    @property
    def conns(self):
        return self.elem.MEPModel.ConnectorManager.Connectors
    
    @property
    def connid_verbinder(self):
        return self._connid_verbinder
    @connid_verbinder.setter
    def connid_verbinder(self,value):
        self._connid_verbinder = value
    @property
    def connid_Leitung(self):
        return self._connid_Leitung
    @connid_Leitung.setter
    def connid_Leitung(self,value):
        self._connid_Leitung = value
    @property
    def elemid_Leitung(self):
        return self._elemid_Leitung
    @elemid_Leitung.setter
    def elemid_Leitung(self,value):
        self._elemid_Leitung = value
    
    
    
    def Parameter_pruefen(self):
        if self.elem.LookupParameter('IGF_X_AnschlussId_Leitung') and self.elem.LookupParameter('IGF_X_AnschlussId_Verbinder') and self.elem.LookupParameter('IGF_X_UniqueId_Leitung'):
            return True
        else:
            return False
    
    def Werte_Pruefen(self):
        if self.Parameter_pruefen():
            if not self.elem.LookupParameter('IGF_X_AnschlussId_Leitung').AsInteger():
                return False
            else:
                self.connid_Leitung = self.elem.LookupParameter('IGF_X_AnschlussId_Leitung').AsInteger()
            if not self.elem.LookupParameter('IGF_X_AnschlussId_Verbinder').AsInteger():
                self.connid_Leitung = None
                return False
            else:
                self.connid_verbinder = self.elem.LookupParameter('IGF_X_AnschlussId_Verbinder').AsInteger()
            if not self.elem.LookupParameter('IGF_X_UniqueId_Leitung').AsString():
                self.connid_Leitung = None
                self.connid_verbinder = None
                return False
            else:
                self.elemid_Leitung = self.elem.LookupParameter('IGF_X_UniqueId_Leitung').AsString()
                return True
        
    def Daten_ermitteln(self):
        for conn in self.conns:
            self.connid_verbinder = conn.Id
            for ref in conn.AllRefs:
                try:
                    if ref.Owner.Category.Id.ToString() in ['-2008000','-2008044']:
                        self.connid_Leitung = ref.Id
                        self.elemid_Leitung = ref.Owner.UniqueId
                        break
                except:
                    pass
            if self.connid_Leitung:
                break
    
    def Daten_pruefen(self):
        if self.connid_Leitung and self.connid_verbinder and self.elemid_Leitung:
            return True

    def werte_schreiben(self):
        def wert_schreiben(param,value):
            try:
                self.elem.LookupParameter(param).Set(value)
            except:
                print('Fehler beim Werte-Schreiben in Parameter "{}".format(param)')
        wert_schreiben('IGF_X_AnschlussId_Leitung',int(self.connid_Leitung))
        wert_schreiben('IGF_X_AnschlussId_Verbinder',int(self.connid_verbinder))
        wert_schreiben('IGF_X_UniqueId_Leitung',str(self.elemid_Leitung))

selection = uidoc.Selection.GetElementIds()
logger.info('{} Baueile ausgewählt.'.format(len(selection)))

if len(selection) == 0:
    UI.TaskDialog.Show(__title__, 'Kein Bauteil ausgewählt')
    script.exit()

system_liste = []

t = DB.Transaction(doc)
t.Start('physische Netze trennen')
with forms.ProgressBar(title="{value}/{max_value} Bauteile ausgewählt", cancellable=True, step=1) as pb:
    for n, elemid in enumerate(selection):
        if pb.cancelled:
            t.RollBack()
            script.exit()
        pb.update_progress(n + 1, len(selection))

        elem = Verbinder(elemid)
        if elem.Daten_pruefen():
            try:
                elem.werte_schreiben()
            except:
                logger.error('Fehler beim Daten-Schreiben der Beiteile {}'.format(elemid.ToString()))
                continue
        
        try:
            leitung = doc.GetElement(elem.elemid_Leitung)
            if not leitung.MEPSystem.Id.ToString() in system_liste:
                system_liste.append(leitung.MEPSystem.Id.ToString())
            
            for conn in elem.elem.MEPModel.ConnectorManager.Connectors: 
                if conn.Id == elem.connid_verbinder: conn_v = conn
            for conn in leitung.ConnectorManager.Connectors: 
                if conn.Id == elem.connid_Leitung: conn_l = conn
            try:
                conn_v.DisconnectFrom(conn_l)
                
            except Exception as e:
                print(30*'-')
                logger.error(e)
                logger.error('ElementId: {}'.format(elemid.ToString()))
                print(30*'-')
        except:
            pass

t.Commit()


with forms.ProgressBar(title="{value}/{max_value} Systeme werden getrennt", cancellable=True, step=1) as pb_t:
    if len(system_liste) == 1:
        pb_t.title = "{value}/{max_value} System wird getrennt"
    for n, systemid in enumerate(system_liste):
        if pb_t.cancelled:
            try:
                t.RollBack()
            except:
                pass
            script.exit()
        pb_t.update_progress(n + 1, len(system_liste))
        system = doc.GetElement(DB.ElementId(int(systemid)))
        t.Start('System trennen')
        dict_param = {}
        hauptanlage = 0
        System_liste_Neu = []
        # Daten von System werden gespeicht.
        for Param in system.Parameters:
            if Param.Definition.Name == 'Systemname':
                continue
            if Param.Definition.Name == 'IGF_X_Hauptanlage':
                hauptanlage = Param.AsInteger()
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
        
        try:
            System_liste_Neu = system.DivideSystem(doc)
        except Exception as e:
            logger.error(e)

        doc.Regenerate()
        t.Commit()

        t.Start('Systemdaten schreiben')
        
        for el in System_liste_Neu:
            try:
                system_neu = doc.GetElement(el)
                system_neu_name = system_neu.LookupParameter('Systemname').AsString()
                print(system_neu_name,el.ToString())
                if system_neu_name[-3:] == '(1)' and hauptanlage != 0:
                    system_neu.LookupParameter('IGF_X_Hauptanlage').Set(1)
                for param_key in dict_param.keys():
                    try:
                        system_neu.LookupParameter(param_key).Set(dict_param[param_key])
                    except:
                        logger.error('Fehler beim Daten-Schreiben. Parameter: {}, Werte: {}, System: {}'.format(param_key,dict_param[param_key],system_neu_name))
            except Exception as e:
                print(30*'-')
                logger.error(e)
                logger.error('ElementId: {}'.format(el.ToString()))         
                print(30*'-')
        doc.Regenerate()
        t.Commit()