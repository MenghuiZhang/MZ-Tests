# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from pyrevit import script, forms
from rpw import revit,DB
from IGF_log import getlog
from IGF_lib import get_value

__title__ = "3.10 Summe-Laboranschlüsse"
__doc__ = """summiert Luftmengen von Laboranschlüsse

[30.09.2021]
Version: 1.0"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

Liste_24h = []
Liste_Labor = []

elem = spaces[0]

for param in doc.GetElement(elem).Parameters:
    if param.Definition.Name.find('IGF_RLT_Laboranschluss_LAB') != -1:
        Liste_Labor.append(param.Definition.Name)
    elif param.Definition.Name.find('IGF_RLT_Laboranschluss_24h') != -1:
        Liste_24h.append(param.Definition.Name)

class MEPRaum:
    def __init__(self, elem_id):
        self.elem_id = elem_id
        self.elem = doc.GetElement(self.elem_id)
        self.Name = self.elem.LookupParameter('Name').AsString()
        self.Nummer = self.elem.Number
        self.Volumen_24h = self.Volumen_24h_berechnen()
        self.Volumen_lab_min,self.Volumen_lab_max = self.Volumen_lab_berechnen()
    
    @property
    def _Lab_Parameters(self):
        lab_params = {}
        for param in Liste_Labor:
            lab_params[param] = get_value(self.elem.LookupParameter(param))
        return lab_params

    @property
    def _24h_Parameters(self):
        _24h_params = {}
        for param in Liste_24h:
            _24h_params[param] = get_value(self.elem.LookupParameter(param))
        return _24h_params

    def Volumen_24h_berechnen(self):
        summe = 0
        for key in self._24h_Parameters.keys():
            wert = self._24h_Parameters[key]
            if wert:
                vol_einzel = int(key[key.find('_kon')+4:])
                try:
                    if wert.find(',') == -1:
                        Liste = [wert]
                    else:
                        Liste = wert.split(',')
                    for i in Liste:
                        if not i:
                            continue
                        while(i[0] == ' '):
                            i = i[1:]
                        try:
                            anzahl = int(i[:i.find('xZeile')])
                            summe = summe + anzahl * vol_einzel
                        except:
                            logger.error('Form passt nicht, '+ key)
                except Exception as e:
                    logger.error(30*'-')
                    logger.error(e)
                    logger.error('24h-Abluft')
                    logger.error(self.Name+', '+ self.Nummer+',' + self.elem_id.ToString())
                    logger.error(30*'-')
        return summe

    def Volumen_lab_berechnen(self):
        summe_min = 0
        summe_max = 0
        for key in self._Lab_Parameters.keys():
            wert = self._Lab_Parameters[key]
            if wert:
                vol_max_einzel = int(key[key.find('_max')+4:])
                vol_min_einzel = int(key[key.find('_min')+4:key.find('_max')])
                try:
                    if wert.find(',') == -1:
                        Liste = [wert]
                    else:
                        Liste = wert.split(',')
                    for i in Liste:
                        if not i:
                            continue
                        while(i[0] == ' '):
                            i = i[1:]
                        try:
                            anzahl = int(i[:i.find('xZeile')])
                            summe_min = summe_min + anzahl * vol_min_einzel
                            summe_max = summe_max + anzahl * vol_max_einzel
                        except:
                            logger.error('Form passt nicht, '+ key)
                except Exception as e:
                    logger.error(30*'-')
                    logger.error(e)
                    logger.error('LAB-ABL-max')
                    logger.error(self.Name+', '+ self.Nummer+',' + self.elem_id.ToString())
                    logger.error(30*'-')
        return summe_min,summe_max

    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                if self.elem.LookupParameter(param_name):
                    self.elem.LookupParameter(param_name).SetValueString(str(wert))

        wert_schreiben("IGF_RLT_AbluftminSummeLabor", self.Volumen_lab_min)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor", self.Volumen_lab_max)
        wert_schreiben("IGF_RLT_AbluftSumme24h", self.Volumen_24h)

    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.Nummer,
            self.Name,
            self.Volumen_24h,
            self.Volumen_lab_min,
            self.Volumen_lab_max,
        ]

table_data = []
mepraum_liste = []
with forms.ProgressBar(title="{value}/{max_value} Luftmengenberechnung", cancellable=True, step=1) as pb:
    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())

#  Sortieren nach Raumnummer
table_data.sort()

output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=["Nummer", "Name", "24-Abluft", "Labor-min", "Labor-max"]
)

# Werte zuückschreiben + Abfrage
if forms.alert("Berechnete Werte in Modell schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP Räume", cancellable=True, step=5) as pb2:
        t = DB.Transaction(doc,'Laboranschluss')
        t.Start()
        for n, mepraum in enumerate(mepraum_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mepraum_liste))
            mepraum.werte_schreiben()
        t.Commit()