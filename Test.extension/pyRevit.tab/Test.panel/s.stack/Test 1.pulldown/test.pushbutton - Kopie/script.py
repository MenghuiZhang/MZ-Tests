# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit import script, forms
from rpw import revit,DB


__title__ = "Parameter ändern(Systeme)"
__doc__ = """"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
doc = revit.doc

try:
    getlog(__title__)
except:
    pass

class MEPSystem:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.elemtyp = doc.GetElement(self.elem.GetTypeId())
        try:
            self.anzahl_alt = self.elem.LookupParameter('TGA_RLT_AnlagenGeräteAnzahl').AsValueString()
        except:
            self.anzahl_alt = None
        try:
            self.genr_alt = self.elem.LookupParameter('TGA_RLT_AnlagenGeräteNr').AsValueString()
        except:
            self.genr_alt = None
        try:
            self.annr_alt = self.elem.LookupParameter('TGA_RLT_AnlagenNr').AsValueString()
        except:
            self.annr_alt = None
        try:
            self.sys_alt = self.elem.LookupParameter('IGF_RLT_SystemKürzel').AsString()
        except:
            self.sys_alt = None
        try:
            self.name_alt = self.elem.LookupParameter('IGF_RLT_SystemName').AsString()
        except:
            self.name_alt = None


        
        self.anzahl_neu = self.elemtyp.LookupParameter('IGF_X_AnlagenGeräteAnzahl')
        self.genr_neu = self.elem.LookupParameter('IGF_X_AnlagenGeräteNr')
        self.annr_neu = self.elemtyp.LookupParameter('IGF_X_AnlagenNr')
        self.sys_neu = self.elemtyp.LookupParameter('IGF_X_SystemKürzel')
        self.name_neu = self.elemtyp.LookupParameter('IGF_X_SystemName')
    
    def wert_schreiben(self,param,wert):
        if wert:
            param.Set(wert)
    def wert_schreiben1(self,param,wert):
        if wert  and wert != '0':
            param.SetValueString(wert)
    def werte_schreiben(self):
        self.wert_schreiben1(self.anzahl_neu,self.anzahl_alt)
        self.wert_schreiben1(self.genr_neu,self.genr_alt)
        self.wert_schreiben1(self.annr_neu,self.annr_alt)
        self.wert_schreiben(self.sys_neu,self.sys_alt)
        self.wert_schreiben(self.name_neu,self.name_alt)


system_Luft = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
system_Rohr = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()

t = DB.Transaction(doc,'Daten schreiben (System)')
t.Start()
for el in system_Luft:
    mepsystem = MEPSystem(el.Id)
    mepsystem.werte_schreiben()

for el in system_Rohr:
    mepsystem = MEPSystem(el.Id)
    mepsystem.werte_schreiben()
t.Commit()