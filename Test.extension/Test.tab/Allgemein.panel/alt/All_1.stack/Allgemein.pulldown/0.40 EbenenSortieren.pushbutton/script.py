# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw,clr
clr.AddReference("PresentationFramework")
import time
from System.Windows.Controls import *
from Autodesk.Revit.DB import Transaction
from pyrevit.forms import WPFWindow, SelectFromList
import System
from System.Windows import Application, Window
from System.Collections.ObjectModel import *



start = time.time()


__title__ = "0.40 EbenenSortieren"
__doc__ = """EbenenSortieren"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)


# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

class Ebenen(object):
    @property
    def Ebene(self):
        return self._Ebene
    @Ebene.setter
    def Ebene(self, value):
        self._Ebene = value
    @property
    def Abk(self):
        return self._Abk
    @Abk.setter
    def Abk(self, value):
        self._Abk = value
    @property
    def Nr(self):
        return self._Nr
    @Nr.setter
    def Nr(self, value):
        self._Nr = value

Liste = []
Liste1 = ObservableCollection[Ebenen]()

for el in spaces_collector:
    le = el.Level.Name
    if not le in Liste:
        Liste.append(le)
        cl_ebene = Ebenen()
        cl_ebene.Ebene = le
        cl_ebene.Abk = None
        cl_ebene.Nr = None
        Liste1.Add(cl_ebene)


class beschriftungen(WPFWindow):
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

Bet = beschriftungen('window.xaml',Liste1)
Bet.ShowDialog()
Nr_dict = {}
Abk_dict = {}
for el in Liste1:
    Nr_dict[el.Ebene] = el.Nr
    Abk_dict[el.Ebene] = el.Abk

t = Transaction(doc,"Ebenenortieren")
t.Start()
title = "{value}/{max_value} MEP Räume"
step = int(len(spaces)/100)+1
with forms.ProgressBar(title=title, cancellable=True, step=step) as pb:
    n = 0
    for item in spaces_collector:
        n += 1
        if pb.cancelled:
            t.RollBack()
            total = time.time() - start
            logger.info("total time: {} {}".format(total, 100 * "_"))
            script.exit()
        pb.update_progress(n, len(spaces))
        name = item.Level.Name
        try:
            if Nr_dict[name]:
                item.LookupParameter("IGF_RLT_Verteilung_EbenenSortieren").Set(int(Nr_dict[name]))
            if Abk_dict[name]:
                item.LookupParameter("IGF_RLT_Verteilung_EbenenName").Set(Abk_dict[name])
        except Exception as e:
            logger.error(e)

t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
