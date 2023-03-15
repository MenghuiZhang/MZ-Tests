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

__title__ = "0.54 RohrSystem_Material"
__doc__ = """System_Material"""
__author__ = "Menghui Zhang"


from pyIGF_logInfo import getlog
getlog(__title__)


ex = Excel.ApplicationClass()
logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

pipe_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Plumbing.PipingSystemType))
pipes = pipe_collector.ToElementIds()
duct_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Mechanical.MechanicalSystemType))
ducts = duct_collector.ToElementIds()
Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
Material_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Material))
Material_dict = {}
for ele in Material_collector:
    Material_dict[ele.Name] = ele.Id

# for el in Family:
#     category = el.FamilyCategory.Name
#     Family = el.Name
#     if category in ['HLS-Bauteile','Rohrzubehör']:
#         Material = el.LookupParameter[]
pipeDaten = []
for el in pipe_collector:
    cate = el.Category.Name
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    pipeDaten.append([cate,Name,Mate])



LuftDaten = []
for el in duct_collector:
    cate = el.Category.Name
    Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    Mate = el.LookupParameter('Material').AsValueString()
    LuftDaten.append([cate,Name,Mate])

new_liste = pipeDaten + LuftDaten

def ExcelLesen(filepath):
    book = ex.Workbooks.Open(filepath)
    Daten = []
    for sheet in book.Worksheets:
        name = sheet.Name
        if name == 'Familien':
            rows = sheet.UsedRange.Rows.Count

            for row in range(2, rows + 1):
                rowlist = []
                if sheet.Cells[row, 1].Value2 == 'Lueftungssysteme Typen':
                    for col in [2,3,4,5]:
                        Wert = sheet.Cells[row, col].Value2
                        rowlist.append(Wert)

                    Daten.append(rowlist)
    book.Save()
    book.Close()

    if any(Daten):
        output.print_table(
                table_data=Daten,
                title="IGF_Material aus Excel",
                columns=['System', 'Material (Werkstoff)', 'Material (Systemfarbe)', 'Material (BLB)']
            )
    return Daten

def ExcelSchreiben(filepath,Daten):
    book = ex.Workbooks.Open(filepath)
    for sheet in book.Worksheets:
        name = sheet.Name
        if name == 'Familien':
            for n in range(len(Daten)):
                sheet.Cells[n + 2,1] = Daten[n][0]
                sheet.Cells[n + 2,2] = Daten[n][1]
                sheet.Cells[n + 2,3] = Daten[n][2]

    book.Save()
    book.Close()


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


# pipe_liste = []
# for item in pipe:
#     Name = item.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
#     category = item.Category.Name
#     pipe_liste.append([category,Name])

def DatenSchreiben(familien,dict,ids):
    title = '{value}/{max_value} Material für Luftsysteme '
    with forms.ProgressBar(title=title,cancellable=True, step=1) as pb:
        n_1 = 0
        for el in familien:
            if pb.cancelled:
                script.exit()
                n_1 += 1
                pb.update_progress(n_1, len(ids))

            if el.LookupParameter('Material'):
                Name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                if Name in dict.keys():
                    if dict[Name] in Material_dict.keys():
                        el.LookupParameter('Material').Set(Material_dict[dict[Name]])


# for item in duct:
#     Name = item.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
#     category = item.Category.Name
#     duct_liste.append([category,Name])

# elec_liste = []
# for item in elec:
#     Name = item.ElectricalSystemType.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
#     category = item.Category.Name
#     elec_liste.append([category,Name])
# output.print_table(
#     table_data=pipe_liste,
#     title="pipe_liste",
#     columns=['Kategorie', 'Name', ]
# # )
# output.print_table(
#     table_data=duct_liste,
#     title="duct_liste",
#     columns=['Kategorie', 'Name', ]
# )


#new_list = pipe_liste+duct_liste

# items = ['Material (Werkstoff)', 'Material (Systemfarbe)', 'Material (BLB)']
# Mate = forms.SelectFromList.show(items, button_name='Select Material')

# path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')


# excel_Daten = ExcelLesen(path)
# dict_Familien = dictErstellen(excel_Daten,Mate)
# trans0 = Transaction(doc, 'Material')
# trans0.Start()
# DatenSchreiben(duct_collector,dict_Familien,ducts)
# trans0.Commit()
path = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material_new.xlsx')
ExcelSchreiben(path,new_liste)

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
