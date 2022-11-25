# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms
import time

start = time.time()
Zustand = time.strftime("%d.%m.%Y", time.localtime())

__title__ = "1.11 MEP-Raumzustandsprüfung"
__doc__ = """
Das alt Raumzustand wird mit dem aktuellen Raumzustand vergliechen.
Wenn ein Raum geändert hat, wird die Information in IGF_A_Änderung geschrieben.

Zustand 01-04.
IGF_A_Änderung_Zustand_01
IGF_A_Bearbeitet_Zustand_01

[2021.11.22]
Version: 1.1
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass
# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()

class MEPRaum:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.volumen = round(get_value(self.elem.LookupParameter('Volumen')),3)
        self.name = get_value(self.elem.LookupParameter('Name'))
        self.nummer = get_value(self.elem.LookupParameter('Nummer'))

        self.volumen_alt = 0
        self.name_alt = ''
        self.nummer_alt = ''
        self.vergleich = ''

    
    def wert_schreiben(self,para,wert):
        param = self.elem.LookupParameter(para)
        if param:
            if param.StorageType.ToString() == 'Double':
                param.SetValueString(str(wert))
            else:
                param.Set(wert)
    
    def getaltdaten_01(self):
        self.nummer_alt = get_value(self.elem.LookupParameter('IGF_A_Raumnummer_Zustand_01'))
        self.name_alt = get_value(self.elem.LookupParameter('IGF_A_Raumname_Zustand_01'))
        self.volumen_alt = round(get_value(self.elem.LookupParameter('IGF_A_Volumen_Zustand_01')),3)

    def getaltdaten_02(self):
        self.nummer_alt = get_value(self.elem.LookupParameter('IGF_A_Raumnummer_Zustand_02'))
        self.name_alt = get_value(self.elem.LookupParameter('IGF_A_Raumname_Zustand_02'))
        self.volumen_alt = round(get_value(self.elem.LookupParameter('IGF_A_Volumen_Zustand_02')),3)

    def getaltdaten_03(self):
        self.nummer_alt = get_value(self.elem.LookupParameter('IGF_A_Raumnummer_Zustand_03'))
        self.name_alt = get_value(self.elem.LookupParameter('IGF_A_Raumname_Zustand_03'))
        self.volumen_alt = round(get_value(self.elem.LookupParameter('IGF_A_Volumen_Zustand_03')),3)

    def getaltdaten_04(self):
        self.nummer_alt = get_value(self.elem.LookupParameter('IGF_A_Raumnummer_Zustand_04'))
        self.name_alt = get_value(self.elem.LookupParameter('IGF_A_Raumname_Zustand_04'))
        self.volumen_alt = round(get_value(self.elem.LookupParameter('IGF_A_Volumen_Zustand_04')),3)

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
        if self.name != self.name_alt:
            self.vergleich += 'Raumname geändert. '
        if self.nummer != self.nummer_alt:
            self.vergleich += 'Raumnummer geändert. '    
    
    def werte_schreiben_01(self):
        if self.vergleich:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_01',False)
        else:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_01',True)
        self.wert_schreiben('IGF_A_Änderung_Zustand_01',self.vergleich)
    
    def werte_schreiben_02(self):
        if self.vergleich:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_02',False)
        else:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_02',True)

        self.wert_schreiben('IGF_A_Änderung_Zustand_02',self.vergleich)
    
    def werte_schreiben_03(self):
        if self.vergleich:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_03',False)
        else:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_03',True)

        self.wert_schreiben('IGF_A_Änderung_Zustand_03',self.vergleich)
    
    def werte_schreiben_04(self):
        if self.vergleich:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_04',False)
        else:
            self.wert_schreiben('IGF_A_Bearbeitet_Zustand_04',True)

        self.wert_schreiben('IGF_A_Änderung_Zustand_04',self.vergleich)

ops = ['1', '2', '3', '4']
Nachfrage = forms.CommandSwitchWindow.show(ops,message = 'Mit welchen Raumzustandsparameter vergliechen werden?')

if not Nachfrage in ops:
    logger.error('Bitte 1, 2, 3, 4 auswählen')
    script.exit()

mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} MEP Räume', cancellable=True, step=10) as pb:
    for n, spaceid in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(spaceid)
        if Nachfrage == '1':
           mepraum.getaltdaten_01()
        elif Nachfrage == '2':
           mepraum.getaltdaten_02()
        elif Nachfrage == '3':
           mepraum.getaltdaten_03() 
        elif Nachfrage == '4':
           mepraum.getaltdaten_04()
        mepraum.Analyse()
        mepraum_liste.append(mepraum)

if forms.alert("Prüfungsergebnis schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP Räume",cancellable=True, step=10) as pb1:
        t = DB.Transaction(doc)
        t.Start('Raumzustand prüfen')
        for n,mepraum in enumerate(mepraum_liste):
            if pb1.cancelled:
                t.RollBack()
                script.exit()
            pb1.update_progress(n+1, len(mepraum_liste))
            if Nachfrage == '1':
                mepraum.werte_schreiben_01()
            elif Nachfrage == '2':
                mepraum.werte_schreiben_02()
            elif Nachfrage == '3':
                mepraum.werte_schreiben_03()
            elif Nachfrage == '4':
                mepraum.werte_schreiben_04()

        t.Commit()