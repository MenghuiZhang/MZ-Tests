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

__title__ = "0.53 System_Material"
__doc__ = """System_Material"""
__author__ = "Menghui Zhang"


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

Material_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Material))
Material_dict = {}
for ele in Material_collector:
    Material_dict[ele.Name] = ele.Id

LuftDaten = []
for el in duct_collector:
    cate = el.Category.Name
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    if Mate == '<Nach Kategorie>':
        Mate = 'Nach Kategorie'
    LuftDaten.append([cate,Name,Mate])
new = []
for el in LuftDaten:
    if el[2] == 'Nach Kategorie':
        new.append(el[1])
for el in pipeDaten:
    if el[2] == 'Nach Kategorie':
        new.append(el[1])
print(new)
LuftFormteil = []
LuftZubehoer = []
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



def ExcelLesen(filepath):
    book = ex.Workbooks.Open(filepath)
    LuftkanalSystemeDaten = []
    for sheet in book.Worksheets:
        name = sheet.Name
        if name == 'Familien':
            rows = sheet.UsedRange.Rows.Count
            for row in range(2, rows + 1):
                rowlist = []
                if sheet.Cells[row, 1].Value2 == 'Luftkanal Systeme':
                    for col in [2,3,4,5]:
                        Wert = sheet.Cells[row, col].Value2
                        rowlist.append(Wert)

                    LuftkanalSystemeDaten.append(rowlist)
    book.Save()
    book.Close()

    if any(LuftkanalSystemeDaten):
        output.print_table(
                table_data=LuftkanalSystemeDaten,
                title="IGF_Material aus Excel",
                columns=['System', 'Material (Werkstoff)', 'Material (Systemfarbe)', 'Material (BLB)']
            )
    return LuftkanalSystemeDaten


def dictErstellen(daten,Material):
    out_dict = {}
    for el in daten:
        if Material == 'Material (Werkstoff)':
            if el[1]:
                out_dict[el[0]] = el[1]
        elif Material == 'Material (Systemfarbe)':
            if el[2]:
                out_dict[el[0]] = el[2]
        elif Material == 'Material (BLB)':
            if el[3]:
                out_dict[el[0]] = el[3]

    return out_dict

def ExcelSchreiben(filepath,Daten):
    book = ex.Workbooks.Open(filepath)
    for sheet in book.Worksheets:
        name = sheet.Name
        Rows = sheet.UsedRange.Rows.Count
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
def ExcelSchreibenSys(filepath,Daten):
    book = ex.Workbooks.Open(filepath)
    for sheet in book.Worksheets:
        name = sheet.Name
        Rows = sheet.UsedRange.Rows.Count
        dict = {}
        if name == 'Familien':
            for n in range(2,Rows+1):
                Familien = sheet.UsedRange.Cells[n,2].Value2
                if Familien in Daten:
                    sheet.Cells[n,3] = 'Nach Kategorie'

    book.Save()
    book.Close()

def DatenSchreiben(familien,dict,ids):
    title = '{value}/{max_value} Material für Luftsysteme '
    trans0 = Transaction(doc, 'Luftsysteme_Material')
    trans0.Start()
    with forms.ProgressBar(title=title,cancellable=True, step=1) as pb:
        n_1 = 0
        for el in familien:
            if pb.cancelled:
                trans0.RollBack()
                script.exit()
            n_1 += 1
            pb.update_progress(n_1, len(ids))

            if el.LookupParameter('Material'):
                Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                if Name in dict.keys():
                    if dict[Name] in Material_dict.keys():
                        el.LookupParameter('Material').Set(Material_dict[dict[Name]])
    trans0.Commit()



# items = ['Material (Werkstoff)', 'Material (Systemfarbe)', 'Material (BLB)']
# Mate = forms.SelectFromList.show(items, button_name='Select Material')

# path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')
# excel_Daten = ExcelLesen(path)
# dict_Familien = dictErstellen(excel_Daten,Mate)
# DatenSchreiben(duct_collector,dict_Familien,ducts)

path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')
# ExcelSchreiben(path,LuftFormteil)
# ExcelSchreiben(path,LuftZubehoer)
ExcelSchreibenSys(path,new)

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
