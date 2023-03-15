# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
import rpw
import time
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

start = time.time()

__title__ = "0.51 Materialien Export"
__doc__ = """Materialien von Luftkanal-, Rohtsysteme, Luftkanalformteile, Luftkanalzubehör, Luftdurchlass, HLS-Bauteile und Rohrformteile exportieren"""
__author__ = "Menghui Zhang"


from pyIGF_logInfo import getlog
getlog(__title__)


ex = Excel.ApplicationClass()
logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

duct_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystemType))
ducts = duct_collector.ToElementIds()
Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))

pipe_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Plumbing.PipingSystemType))
pipes = pipe_collector.ToElementIds()
pipeDaten = []
for el in pipe_collector:
    cate = el.Category.Name
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    if Mate == '<Nach Kategorie>':
        Mate = 'Nach Kategorie'
    pipeDaten.append([cate,Name,Mate])

LuftDaten = []
for el in duct_collector:
    cate = el.Category.Name
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    if Mate == '<Nach Kategorie>':
        Mate = 'Nach Kategorie'
    LuftDaten.append([cate,Name,Mate])


LuftFormteil = []
LuftZubehoer = []
Luftdurchlass = []
HLS = []
Rohrzubehoer = []
Rohrformteile = []
for el in Family:
    category = el.FamilyCategory.Name
    Tpy = el.GetFamilySymbolIds()
    symbol = None
    for i in Tpy:
        symbol = doc.GetElement(i)
        break
    Family = el.Name
    Farbe = None
    if not symbol:
        continue
    if category == 'Luftkanalformteile':
        if symbol.LookupParameter('CAx Material'):
            Farbe = symbol.LookupParameter('CAx Material').AsValueString()
        if symbol.LookupParameter('CAx Materialkz') and Farbe == None:
            Farbe = symbol.LookupParameter('CAx Materialkz').AsString()
        if symbol.LookupParameter('LIN_VE_DIM_MATERIAL') and Farbe == None:
            Farbe = symbol.LookupParameter('LIN_VE_DIM_MATERIAL').AsValueString()
        if Farbe == '<Nach Kategorie>':
            Farbe = 'Nach Kategorie'
        LuftFormteil.append([category,Family,Farbe])

    elif category == 'Luftkanalzubehör':
        if symbol.LookupParameter('CAx Material'):
            Farbe = symbol.LookupParameter('CAx Material').AsValueString()
        if symbol.LookupParameter('CAx Materialkz') and Farbe == None:
            Farbe = symbol.LookupParameter('CAx Materialkz').AsString()
        if symbol.LookupParameter('LIN_VE_DIM_MATERIAL') and Farbe == None:
            Farbe = symbol.LookupParameter('LIN_VE_DIM_MATERIAL').AsValueString()
        if Farbe == '<Nach Kategorie>':
            Farbe = 'Nach Kategorie'
        LuftZubehoer.append([category,Family,Farbe])

    elif category == 'Luftdurchlässe':
        if symbol.LookupParameter('CAx Material'):
            Farbe = symbol.LookupParameter('CAx Material').AsValueString()
        if symbol.LookupParameter('CAx Materialkz') and Farbe == None:
            Farbe = symbol.LookupParameter('CAx Materialkz').AsString()
        if symbol.LookupParameter('LIN_VE_DIM_MATERIAL') and Farbe == None:
            Farbe = symbol.LookupParameter('LIN_VE_DIM_MATERIAL').AsValueString()
        if Farbe == '<Nach Kategorie>':
            Farbe = 'Nach Kategorie'
        Luftdurchlass.append([category,Family,Farbe])

    elif category == 'Rohrformteile':
        if symbol.LookupParameter('CAx Material'):
            Farbe = symbol.LookupParameter('CAx Material').AsValueString()
        if symbol.LookupParameter('CAx Materialkz') and Farbe == None:
            Farbe = symbol.LookupParameter('CAx Materialkz').AsString()
        if symbol.LookupParameter('LIN_VE_DIM_MATERIAL') and Farbe == None:
            Farbe = symbol.LookupParameter('LIN_VE_DIM_MATERIAL').AsValueString()
        if symbol.LookupParameter('ltp_Material') and Farbe == None:
            Farbe = symbol.LookupParameter('ltp_Material').AsValueString()
        if symbol.LookupParameter('Color') and Farbe == None:
            Farbe = symbol.LookupParameter('Color').AsValueString()
        if Farbe == '<Nach Kategorie>':
            Farbe = 'Nach Kategorie'
        Rohrformteile.append([category,Family,Farbe])

    elif category == 'HLS-Bauteile':
        if symbol.LookupParameter('CAx Material'):
            Farbe = symbol.LookupParameter('CAx Material').AsValueString()
        if symbol.LookupParameter('CAx Materialkz') and Farbe == None:
            Farbe = symbol.LookupParameter('CAx Materialkz').AsString()
        if symbol.LookupParameter('LIN_VE_DIM_MATERIAL') and Farbe == None:
            Farbe = symbol.LookupParameter('LIN_VE_DIM_MATERIAL').AsValueString()
        if Farbe == '<Nach Kategorie>':
            Farbe = 'Nach Kategorie'
        HLS.append([category,Family,Farbe])

    # elif category == 'Rohrzubehör':
    #     if symbol.LookupParameter('CAx Material'):
    #         Farbe = symbol.LookupParameter('CAx Material').AsValueString()
    #     if symbol.LookupParameter('CAx Materialkz') and Farbe == None:
    #         Farbe = symbol.LookupParameter('CAx Materialkz').AsString()
    #     if symbol.LookupParameter('LIN_VE_DIM_MATERIAL') and Farbe == None:
    #         Farbe = symbol.LookupParameter('LIN_VE_DIM_MATERIAL').AsValueString()
    #     if Farbe == '<Nach Kategorie>':
    #         Farbe = 'Nach Kategorie'
    #     Rohrzubehoer.append([category,Family,Farbe])

def ExcelSchreiben(filepath,Daten):
    book = ex.Workbooks.Open(filepath)
    for sheet in book.Worksheets:
        name = sheet.Name
        Rows = sheet.UsedRange.Rows.Count
        dict = {}
        if name == 'Familien':
            if len(dict) == 0:
                for n in range(2,Rows+1):
                    Familien = sheet.UsedRange.Cells[n,2].Value2
                    dict[Familien] = n
            for el in Daten:
                if not el[1] in dict.keys():
                    _rows = sheet.UsedRange.Rows.Count
                    sheet.Cells[_rows+1,1] = el[0]
                    sheet.Cells[_rows+1,2] = el[1]
                    sheet.Cells[_rows+1,3] = el[2]
                    dict[el[1]] = _rows+1

    book.Save()
    book.Close()

path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')
ExcelSchreiben(path,LuftDaten)
ExcelSchreiben(path,pipeDaten)
ExcelSchreiben(path,LuftFormteil)
ExcelSchreiben(path,LuftZubehoer)
ExcelSchreiben(path,Luftdurchlass)
ExcelSchreiben(path,HLS)
ExcelSchreiben(path,Rohrformteile)


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
