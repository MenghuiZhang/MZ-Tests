# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms

__title__ = "1.00 MEP Räume aktualisieren"
__doc__ = """
MEP-Räume aktualisieren.
Der Basisversatz des Raumes wird zunächst auf 1 gesetzt und dann auf den ursprünglichen Wert gesetzt.

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

MEPRaumIds = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()

class MEPRaum:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.wert = self.elem.LookupParameter('Basisversatz').AsValueString()
    
    def changebasisversatz(self):
        if self.wert == '1':
            try:
                self.elem.LookupParameter('Basisversatz').SetValueString('0')
            except:
                pass
        else:
            try:
                self.elem.LookupParameter('Basisversatz').SetValueString('1')
            except:
                pass
    
    def changebasisversatz1(self):
        try:
            self.elem.LookupParameter('Basisversatz').SetValueString(self.wert)
        except:
            pass

mepraumliste = []

with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=10) as pb:
    for n, raum in enumerate(MEPRaumIds):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(MEPRaumIds))
        mepraum = MEPRaum(raum)
        mepraumliste.append(mepraum)

if forms.alert('MEP-Räume aktualisieren?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=10) as pb1:
        trans = DB.Transaction(doc,"Basisversatz")
        trans.Start()
        for n, raum in enumerate(mepraumliste):
            if pb1.cancelled:
                trans.RollBack()
                script.exit()
            pb1.update_progress(n+1, len(mepraumliste))
            raum.changebasisversatz()

        doc.Regenerate()
        trans.Commit()
        trans.Dispose()
        
    with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=10) as pb1:
        trans = DB.Transaction(doc,"MEP-Räume aktualisieren")
        trans.Start()
        for n, raum in enumerate(mepraumliste):
            if pb1.cancelled:
                trans.RollBack()
                script.exit()
            pb1.update_progress(n+1, len(mepraumliste))
            raum.changebasisversatz1()
        trans.Commit()
        trans.Dispose()