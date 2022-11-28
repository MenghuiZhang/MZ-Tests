# coding: utf8
import time
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from IGF_forms import Texteingeben
from rpw import revit, DB
from pyrevit import script, forms

datum = time.strftime("%d.%m.%Y", time.localtime())

__title__ = "1.10 MEP-Raumzustandsänderung"
__doc__ = """
schreibt Raumzustandsdaten in Zustandsparameter.
Bezug auf Volumen, Umfang, Licht Höhe, Fläche, Name und Nummer.

Parameter:
Zustand 01-04.
IGF_A_Wichtigkeit_Zustand_01
IGF_A_Raumnummer_Zustand_01
IGF_A_Raumname_Zustand_01
IGF_A_Fläche_Zustand_01
IGF_A_Volumen_Zustand_01
IGF_A_LichteHöhe_Zustand_01
IGF_A_Umfang_Zustand_01
IGF_A_Datum_Zustand_01

[2021.11.22]
Version: 1.1
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()

class MEPRaum:
    def __init__(self,elemid,wichtigkeit):
        self.elemid = elemid
        self.wichtig = wichtigkeit
        self.elem = doc.GetElement(self.elemid)
        self.flaeche = get_value(self.elem.LookupParameter('Fläche'))
        self.volumen = get_value(self.elem.LookupParameter('Volumen'))
        self.hoehe = get_value(self.elem.LookupParameter('Lichte Höhe'))
        self.umfang = get_value(self.elem.LookupParameter('Umfang'))
        self.name = get_value(self.elem.LookupParameter('Name'))
        self.nummer = get_value(self.elem.LookupParameter('Nummer'))
    
    def wert_schreiben(self,para,wert):
        param = self.elem.LookupParameter(para)
        if param:
            if param.StorageType.ToString() == 'Double':
                param.SetValueString(str(wert))
            else:
                param.Set(wert)
    
    def werte_schreiben_01(self):
        self.wert_schreiben('IGF_A_Wichtigkeit_Zustand_01',self.wichtig)
        self.wert_schreiben('IGF_A_Raumnummer_Zustand_01',self.nummer)
        self.wert_schreiben('IGF_A_Raumname_Zustand_01',self.name)
        self.wert_schreiben('IGF_A_Fläche_Zustand_01',self.flaeche)
        self.wert_schreiben('IGF_A_Volumen_Zustand_01',self.volumen)
        self.wert_schreiben('IGF_A_LichteHöhe_Zustand_01',self.hoehe)
        self.wert_schreiben('IGF_A_Umfang_Zustand_01',self.umfang)
        self.wert_schreiben('IGF_A_Datum_Zustand_01',datum)
    
    def werte_schreiben_02(self):
        self.wert_schreiben('IGF_A_Wichtigkeit_Zustand_02',self.wichtig)
        self.wert_schreiben('IGF_A_Raumnummer_Zustand_02',self.nummer)
        self.wert_schreiben('IGF_A_Raumname_Zustand_02',self.name)
        self.wert_schreiben('IGF_A_Fläche_Zustand_02',self.flaeche)
        self.wert_schreiben('IGF_A_Volumen_Zustand_02',self.volumen)
        self.wert_schreiben('IGF_A_LichteHöhe_Zustand_02',self.hoehe)
        self.wert_schreiben('IGF_A_Umfang_Zustand_02',self.umfang)
        self.wert_schreiben('IGF_A_Datum_Zustand_02',datum)
    
    def werte_schreiben_03(self):
        self.wert_schreiben('IGF_A_Wichtigkeit_Zustand_03',self.wichtig)
        self.wert_schreiben('IGF_A_Raumnummer_Zustand_03',self.nummer)
        self.wert_schreiben('IGF_A_Raumname_Zustand_03',self.name)
        self.wert_schreiben('IGF_A_Fläche_Zustand_03',self.flaeche)
        self.wert_schreiben('IGF_A_Volumen_Zustand_03',self.volumen)
        self.wert_schreiben('IGF_A_LichteHöhe_Zustand_03',self.hoehe)
        self.wert_schreiben('IGF_A_Umfang_Zustand_03',self.umfang)
        self.wert_schreiben('IGF_A_Datum_Zustand_03',datum)
    
    def werte_schreiben_04(self):
        self.wert_schreiben('IGF_A_Wichtigkeit_Zustand_04',self.wichtig)
        self.wert_schreiben('IGF_A_Raumnummer_Zustand_04',self.nummer)
        self.wert_schreiben('IGF_A_Raumname_Zustand_04',self.name)
        self.wert_schreiben('IGF_A_Fläche_Zustand_04',self.flaeche)
        self.wert_schreiben('IGF_A_Volumen_Zustand_04',self.volumen)
        self.wert_schreiben('IGF_A_LichteHöhe_Zustand_04',self.hoehe)
        self.wert_schreiben('IGF_A_Umfang_Zustand_04',self.umfang)
        self.wert_schreiben('IGF_A_Datum_Zustand_04',datum)

ops = ['1', '2', '3', '4']
Nachfrage = forms.CommandSwitchWindow.show(ops,message = 'In welchen Zustandparanmeter sollen die ZustandsänderungsDaten schreiben?')
if not Nachfrage in ops:
    logger.error('Bitte 1, 2, 3, 4 auswählen')
    script.exit()

wichtigkeit = Texteingeben(label='Wichtigkeit', text = "")
wichtigkeit.Titel = __title__
wichtigkeit.show_dialog()
wichtig = wichtigkeit.text.Text

mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} MEP Räume', cancellable=True, step=10) as pb:
    for n, spaceid in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(spaceid,wichtig)
        mepraum_liste.append(mepraum)

if forms.alert("Raumzustand übernehmen?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP Räume",cancellable=True, step=10) as pb1:
        t = DB.Transaction(doc)
        t.Start('Raumzustand')
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