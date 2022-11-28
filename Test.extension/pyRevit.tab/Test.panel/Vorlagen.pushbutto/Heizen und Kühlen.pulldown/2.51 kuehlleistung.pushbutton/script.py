# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from IGF_libKopie import get_value

__title__ = "2.51 Kühlleistung MEP Räume"
__doc__ = """Kühlleistungberechnung

input Parameter:
IGF_K_DeS_Leistung: Kühlleistung DeS
IGF_K_ULK_Leistung: Kühlleistung ULK
IGF_K_KA_Leistung: Kühlleistung Kälte
IGF_RLT_ZuluftKühlleistung: Kühlleistung Zuluft


IGF_S_KühllastLaborPWK: Kühllast für Laboreinrichtung über PKW
LIN_BA_CALCULATED_COOLING_LOAD: Kühllast Gebäude

output Parameter:
IGF_K_KühlleistungRaum: Summe von Zuluft- & DeS- & ULK- & Kältekühlleistung

IGF_RLT_ZuluftKühlleistung: Zuluftmenge*c*Dt
IGF_K_KühllastGesamt: Summe von Kühllast Gebäude und Kühllast Labor Raum
IGF_K_KühlleistungULK: Differenz von KühllastGesamt und ZuluftKühlleistung 

[Version: 1.1]
[2021.11.16]"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc


# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()


class MEPRaum:
    def __init__(self, element_id):

        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ['name', 'Name'],
            ['nummer', 'Nummer'],
            ['KL_DeS', 'IGF_K_DeS_Leistung'],
            ['KL_ULK', 'IGF_K_ULK_Leistung'],
            ['KL_KA', 'IGF_K_KA_Leistung'],
            ['KL_Zul', 'IGF_RLT_ZuluftKühlleistung'],
        ]

        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.get_element_attr(revit_name))

        # gesamte Kühlleistung berechnen
        self.KL_gesamt = self.KL_gesamt_berechnen()

    def get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return get_value(param)

    def Werte_Pruefen(self,para):
        if not para:
            para = 0
        return para


    def KL_gesamt_berechnen(self):
        Kuelhlleistung = 0
        self.KL_DeS = self.Werte_Pruefen(self.KL_DeS)
        self.KL_ULK = self.Werte_Pruefen(self.KL_ULK)
        self.KL_KA = self.Werte_Pruefen(self.KL_KA)
        self.KL_Zul = self.Werte_Pruefen(self.KL_Zul)

        Kuelhlleistung = self.KL_DeS + self.KL_ULK + self.KL_KA + self.KL_Zul

        return round(Kuelhlleistung, 2)


    def werte_schreiben(self):
        def wert_schreiben(param_name, wert):
            if not wert is None:
                self.element.LookupParameter(
                    param_name).SetValueString(str(wert))

        wert_schreiben("IGF_K_KühlleistungRaum", self.KL_gesamt)


mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} Kühlleistung berechnen',cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)
        mepraum_liste.append(mepraum)

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben',cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Kühlleistung MEP Räume')

        for n,mepraum in enumerate(mepraum_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mepraum_liste))
            mepraum.werte_schreiben()
        t.Commit()