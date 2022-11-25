# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms

__title__ = "1.31 Architektur-Raumzustandsprüfung"
__doc__ = """ Das alt Raumzustand wird mit dem aktuellen Raumzustand vergliechen.


input Parameter: 
IGF_A_Volumen_Architektur,
Volumen, 
IGF_A_Nummer_Architektur, 
Nummer

output Parameter:
IGF_A_Änderung_Architektur
IGF_A_Bearbeitet_Architektur

[2021.11.25]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass
uidoc = revit.uidoc
doc = revit.doc
logger = script.get_logger()
output = script.get_output()


spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()


if not spaces:
    logger.error("Keine Räume in aktueller Projekt gefunden")
    script.exit()


class Raum:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.volumen = round(get_value(self.elem.LookupParameter('Volumen')),3)
        self.nummer = get_value(self.elem.LookupParameter('Nummer'))

        self.volumen_alt = round(get_value(self.elem.LookupParameter('IGF_A_Volumen_Architektur')),3)
        self.nummer_alt = get_value(self.elem.LookupParameter('IGF_A_Nummer_Architektur'))
        self.vergleich = ''
        self.Analyse()

    
    def wert_schreiben(self,para,wert):
        param = self.elem.LookupParameter(para)
        if param:
            if param.StorageType.ToString() == 'Double':
                param.SetValueString(str(wert))
            else:
                param.Set(wert)

    def Analyse(self):
        if self.volumen > 0 and self.volumen_alt > 0:
            abgleich = self.volumen - self.volumen_alt
            if abgleich > 1:
                self.vergleich += 'größer als zuvor, Abweichung: ' + str(round(abgleich)) + '. '
            elif abgleich < -1:
                self.vergleich += 'kleiner als zuvor, Abweichung: ' + str(round(abgleich)) + '. '
        else:
            if self.volumen <= 0:
                self.vergleich += 'Das aktuelle Raumvolumen ist kleiner als 0. '
            if self.volumen_alt <= 0:
                self.vergleich += 'Das alte Raumvolumen ist kleiner als 0. '

        if self.nummer != self.nummer_alt:
            self.vergleich += 'Raumnummerabweichung. '    
    
    def werte_schreiben(self):
        if self.vergleich:
            self.wert_schreiben('IGF_A_Bearbeitet_Architektur',False)
        else:
            self.wert_schreiben('IGF_A_Bearbeitet_Architektur',True)

        self.wert_schreiben('IGF_A_Änderung_Architektur',self.vergleich)


raum_liste = []
with forms.ProgressBar(title='{value}/{max_value} Räume', cancellable=True, step=10) as pb:
    for n, spaceid in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        raum = Raum(spaceid)
        raum_liste.append(raum)


if forms.alert("Prüfungsergebnis schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP Räume",cancellable=True, step=10) as pb1:
        t = DB.Transaction(doc)
        t.Start('Raumzustand prüfen')
        for n,raum in enumerate(raum_liste):
            if pb1.cancelled:
                t.RollBack()
                script.exit()
            pb1.update_progress(n + 1, len(raum_liste))
            raum.werte_schreiben()
        t.Commit()
        t.Dispose()