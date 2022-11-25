# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script,forms
import time
from System.Windows.Controls import *
from Autodesk.Revit.DB import *
from pyrevit.forms import WPFWindow
import System
from System.Windows import Application, Window
from System.Collections.ObjectModel import *
from pyIGF_logInfo import getlog

start = time.time()


__title__ = "1.40 schreibt Geschosszugehörigkeit in Trasse"
__doc__ = """IGF_Trassenzugehörigkeit (HLS-Bauteile, Kabeltrassen, Leerrohr, Luftkanäle, Rohre)"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
getlog(__title__)

uidoc = revit.uidoc
doc = revit.doc

HLS_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
    .WhereElementIsNotElementType()
HLS = HLS_collector.ToElementIds()

Rohre_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_PipeCurves)\
    .WhereElementIsNotElementType()
Rohre = Rohre_collector.ToElementIds()

Luft_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_DuctCurves)\
    .WhereElementIsNotElementType()
Luft = Luft_collector.ToElementIds()

Kabel_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_CableTray)\
    .WhereElementIsNotElementType()
Kabel = Kabel_collector.ToElementIds()

Leer_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Conduit)\
    .WhereElementIsNotElementType()
Leer = Leer_collector.ToElementIds()

class TrassenEbenen(object):
    @property
    def CAx(self):
        return self._CAx
    @CAx.setter
    def CAx(self, value):
        self._CAx = value
    @property
    def IGF(self):
        return self._IGF
    @IGF.setter
    def IGF(self, value):
        self._IGF = value


Liste = []
Liste1 = ObservableCollection[TrassenEbenen]()
def Ermitteln(coll):
    for el in coll:
        le = el.LookupParameter('CAx_Trassenbezugsebene').AsString()
        if not le in Liste and le:
            Liste.append(le)
            cl_ebene = TrassenEbenen()
            cl_ebene.CAx = le
            cl_ebene.IGF = le
            Liste1.Add(cl_ebene)

Ermitteln(HLS_collector)
Ermitteln(Rohre_collector)
Ermitteln(Luft_collector)
Ermitteln(Kabel_collector)
Ermitteln(Leer_collector)

class EbenenVereinheitlichen(WPFWindow):
    def __init__(self, xaml_file_name,liste):
        self.Liste = liste
        WPFWindow.__init__(self, xaml_file_name)
        self.dataGrid.ItemsSource = self.Liste

    def ok(self,sender,args):
        self.Close()
    def cancel(self,sender,args):
        self.Close()
        total = time.time() - start
        logger.info("total time: {} {}".format(total, 100 * "_"))
        script.exit()

EbVer = EbenenVereinheitlichen('window.xaml',Liste1)
EbVer.ShowDialog()
Ebene_dict = {}
for el in Liste1:
    Ebene_dict[el.CAx] = el.IGF


t = Transaction(doc,'Trassenzugehörigkeit')
t.Start()

def DatenSchreiben(collector,Kennzeichen,ids):
    title = '{value}/{max_value} Trassenzugehörigkeit-' + Kennzeichen
    step = int(len(ids)/200)
    if step == 0:
        step = 1
    with forms.ProgressBar(title=title, cancellable=True, step=step) as pb:
        n = 0

        for el in collector:
            if pb.cancelled:
                t.RollBack()
                script.exit()
            n += 1
            pb.update_progress(n, len(ids))
            Ebene = el.LookupParameter('CAx_Trassenbezugsebene').AsString()
            if not Ebene:
                continue
            if Ebene in Ebene_dict.keys():
                el.LookupParameter('IGF_Trassenzugehörigkeit').Set(Ebene_dict[Ebene])


DatenSchreiben(HLS_collector,'HLS_Bauteile',HLS)
DatenSchreiben(Rohre_collector,'Rohre',Rohre)
DatenSchreiben(Luft_collector,'Luftkanäle',Luft)
DatenSchreiben(Kabel_collector,'Kabeltrassen',Kabel)
DatenSchreiben(Leer_collector,'Leerrohr',Leer)
t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
