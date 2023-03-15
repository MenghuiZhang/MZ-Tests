# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
import rpw
import time
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

start = time.time()

__title__ = "0.54 Formteile/Zubehör_Materialien"
__doc__ = """Formteile/Zubehör_Materialien(Luftkanalzubehör, Luftdurchlässe, Luftkanalformteile,
HLS-Bauteile, Rohrzubehör, Rohrformteile)"""
__author__ = "Menghui Zhang"


from pyIGF_logInfo import getlog
getlog(__title__)


ex = Excel.ApplicationClass()
logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

duct_collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystemType))
ducts = duct_collector.ToElementIds()

pipe_collector = FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Plumbing.PipingSystemType))
pipes = pipe_collector.ToElementIds()

Luftkanalzubehoer_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType()
Luftkanalzubehoer = Luftkanalzubehoer_collector.ToElementIds()

Luftkanalformteile_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType()
Luftkanalformteile = Luftkanalformteile_collector.ToElementIds()

Luftdurchlaesse_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
Luftdurchlaesse = Luftdurchlaesse_collector.ToElementIds()

HLS_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
HLS = HLS_collector.ToElementIds()

Rohrzubehoer_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType()
Rohrzubehoer = Luftkanalzubehoer_collector.ToElementIds()

Rohrformteile_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType()
Rohrformteile = Rohrformteile_collector.ToElementIds()


Material_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Material))
Material_dict = {}
for ele in Material_collector:
    Material_dict[ele.Name] = ele.Id

pipeDaten = {}
for el in pipe_collector:
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    if Mate != '<Nach Kategorie>':
        pipeDaten[Mate] = Mate

LuftDaten = {}
for el in duct_collector:
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    if Mate != '<Nach Kategorie>':
        LuftDaten[Name] = Mate

#systenDaten = pipeDaten + LuftDaten

def ExcelLesen(filepath):
    book = ex.Workbooks.Open(filepath)
    exceldaten = {}
    for sheet in book.Worksheets:
        name = sheet.Name
        if name == 'Familien':
            rows = sheet.UsedRange.Rows.Count
            for row in range(2, rows + 1):
                rowlist = []
                if not sheet.Cells[row, 1].Value2 in ['Luftkanal Systeme', 'Rohr Systeme']:
                    Wert1 = sheet.Cells[row, 2].Value2
                    Wert2 = sheet.Cells[row, 10].Value2
                    Wert3 = sheet.Cells[row, 4].Value2
                    exceldaten[Wert1] = [Wert2,Wert3]
    book.Save()
    book.Close()
    return exceldaten


def DatenSchreiben(familien,dict,ids,kate):
    title = '{value}/{max_value} Element in Kategorien '+kate
    trans0 = Transaction(doc, 'Material in' + kate)
    trans0.Start()
    with forms.ProgressBar(title=title,cancellable=True, step=1) as pb:
        n_1 = 0
        for el in familien:
            if pb.cancelled:
                trans0.RollBack()
                script.exit()
            n_1 += 1
            pb.update_progress(n_1, len(ids))
            try:
                if el.LookupParameter('IGF_Material'):
                    FamilyName = el.Symbol.FamilyName
                    SystemType = el.LookupParameter('Systemtyp').AsValueString()
                    if FamilyName in dict.keys():
                        if dict[FamilyName][0] == 'System':
                            if SystemType in pipeDaten.keys():
                                el.LookupParameter('IGF_Material').Set(Material_dict[pipeDaten[SystemType]])
                            elif SystemType in LuftDaten.keys():
                                el.LookupParameter('IGF_Material').Set(Material_dict[LuftDaten[SystemType]])
                        elif dict[FamilyName][0] == 'Bauteilart':
                            if dict[FamilyName][1] in Material_dict.keys():
                                el.LookupParameter('IGF_Material').Set(Material_dict[dict[FamilyName][1]])
            except Exception as E:
                logger.error(E)


    trans0.Commit()



path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')
excel_Daten = ExcelLesen(path)

DatenSchreiben(Luftkanalzubehoer_collector,excel_Daten,Luftkanalzubehoer,'Luftkanalzubehör')
DatenSchreiben(Luftkanalformteile_collector,excel_Daten,Luftkanalformteile,'Luftkanalformteile')
DatenSchreiben(Luftdurchlaesse_collector,excel_Daten,Luftdurchlaesse,'Luftdurchlässe')
#DatenSchreiben(HLS_collector,excel_Daten,HLS,'HLS_Bauteile')
# DatenSchreiben(Rohrzubehoer_collector,excel_Daten,Rohrzubehoer)
# DatenSchreiben(Rohrformteile_collector,excel_Daten,Rohrformteile)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
