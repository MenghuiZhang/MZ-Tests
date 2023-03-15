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


from pyIGF_logInfo import getlog
getlog(__title__)


ex = Excel.ApplicationClass()
logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

duct_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystemType))
ducts = duct_collector.ToElementIds()

pipe_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Plumbing.PipingSystemType))
pipes = pipe_collector.ToElementIds()


Material_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Material))
Material_dict = {}
for ele in Material_collector:
    Material_dict[ele.Name] = ele.Id


def ExcelLesen(filepath):
    book = ex.Workbooks.Open(filepath)
    SystemDaten = []
    for sheet in book.Worksheets:
        name = sheet.Name
        if name == 'Familien':
            rows = sheet.UsedRange.Rows.Count
            for row in range(2, rows + 1):
                rowlist = []
                if sheet.Cells[row, 1].Value2 in ['Luftkanal Systeme', 'Rohr Systeme']:
                    for col in [2,3,4,5]:
                        Wert = sheet.Cells[row, col].Value2
                        rowlist.append(Wert)
                    SystemDaten.append(rowlist)
    book.Save()
    book.Close()
    SystemDaten.sort()

    if any(SystemDaten):
        output.print_table(
                table_data=SystemDaten,
                title="IGF_Material aus Excel",
                columns=['System', 'Material (Werkstoff)', 'Material (Systemfarbe)', 'Material (BLB)']
            )
    return SystemDaten


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


def DatenSchreiben(familien,dict,ids):
    title = '{value}/{max_value} Material für Systeme '
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



items = ['Material (Werkstoff)', 'Material (Systemfarbe)', 'Material (BLB)']
Mate = forms.SelectFromList.show(items, button_name='Select Material')

path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')
excel_Daten = ExcelLesen(path)
dict_Familien = dictErstellen(excel_Daten,Mate)
DatenSchreiben(duct_collector,dict_Familien,ducts)
#DatenSchreiben(pipe_collector,dict_Familien,pipes)



total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
