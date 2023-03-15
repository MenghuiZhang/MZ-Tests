# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "0.41 _RUB_EbenenSortieren"
__doc__ = """_RUB_EbenenSortieren"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

EbeneNummer = {
    2285843: -6,
    2285844: -5,
    2285845: -4,
    2285846: -3,
    2285847: -2,
    2285848: -1,
    2283079: 1,
    2283080: 2,
    2283081: 3,
    2283082: 4,
    2283083: 5,
    2283084: 6,
    2283085: 7,
    2283086: 8,
    1862541: 9,
}
EbeneKurz = {
    2285843: 'E06',
    2285844: 'E05',
    2285845: 'E04',
    2285846: 'E03',
    2285847: 'E02',
    2285848: 'E01',
    2283079: 'E1',
    2283080: 'E2',
    2283081: 'E3',
    2283082: 'E4',
    2283083: 'E5',
    2283084: 'E6',
    2283085: 'E7',
    2283086: 'E8',
    1862541: 'E9',
}
# EbeneNummer = {
#     2285843: 0,
#     2285844: 0,
#     2285845: 0,
#     2285846: 0,
#     2285847: 0,
#     2285848: 0,
#     2283079: 0,
#     2283080: 0,
#     2283081: 0,
#     2283082: 0,
#     2283083: 0,
#     2283084: 0,
#     2283085: 0,
#     2283086: 0,
#     1862541: 0,
# }
# EbeneKurz = {
#     2285843: '',
#     2285844: '',
#     2285845: '',
#     2285846: '',
#     2285847: '',
#     2285848: '',
#     2283079: '',
#     2283080: '',
#     2283081: '',
#     2283082: '',
#     2283083: '',
#     2283084: '',
#     2283085: '',
#     2283086: '',
#     1862541: '',
# }
t = Transaction(doc,"Ebenen")
t.Start()
for item in spaces_collector:
    levelid = int(item.LevelId.ToString())
    nummer = EbeneNummer[levelid]
    name = EbeneKurz[levelid]
    item.LookupParameter("IGF_RLT_Verteilung_EbenenSortieren").Set(nummer)
    item.LookupParameter("IGF_RLT_Verteilung_EbenenName").Set(name)
t.Commit()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
