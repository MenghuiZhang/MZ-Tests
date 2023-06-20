# coding: utf8
import os
from pyrevit import revit, DB
from pyrevit import script
import xlsxwriter
from AK_Liste_GUI import Excelerstellen
from IGF_forms import ExcelSuche
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
from System.Collections.Generic import List

__title__ = "BSK Vergleichen"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()


def exportstand0():
    elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
    cl = []
    for el in elems:
        if el.Symbol.FamilyName.find('BSK') != -1 or el.Symbol.FamilyName.find('Brand') != -1 or el.Name.find('BSK') != -1 or el.Name.find('Brand') != -1:
            cl.append(el)
    elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
    for el in elems:
        if el.Symbol.FamilyName.find('CAx Abluft Brandschutzventil RU - Wandeinbau') != -1 or el.Symbol.FamilyName.find('CAx Zuluft Brandschutzventil RU - Wandeinbau') != -1:
            cl.append(el)
    liste = [['Nr','X','Y','Z','elemid']]
    for el in cl:
        box = el.get_BoundingBox(None)
        locat = (box.Min + box.Max) /2
        nr = el.LookupParameter('IGF_RLT_BSK-Nummer').AsString()
        liste.append([nr,locat.X,locat.Y,locat.Z,el.Id.IntegerValue])
    path = r'C:\Users\zhang\Desktop\BSK_MFC_alt.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(5):
            worksheet.write(row,col, liste[row][col])
    workbook.close()



def exportstand1():
    elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
    cl = []
    for el in elems:
        if el.Symbol.FamilyName.find('BSK') != -1 or el.Symbol.FamilyName.find('Brand') != -1 or el.Name.find('BSK') != -1 or el.Name.find('Brand') != -1:
            cl.append(el.Id)
    elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
    for el in elems:
        if el.Symbol.FamilyName.find('CAx Abluft Brandschutzventil RU - Wandeinbau') != -1 or el.Symbol.FamilyName.find('CAx Zuluft Brandschutzventil RU - Wandeinbau') != -1:
            cl.append(el.Id)
    liste_elemid = List[DB.ElementId](cl)

    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(r'C:\Users\zhang\Desktop\BSK_MFC_alt.xlsx',FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)
    _dict = {}
    try:
        for sheet in book.Workbook.Worksheets:
            
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                liste = []
                nummer = sheet.Cells[row, 1].Value
                x = sheet.Cells[row, 2].Value
                y = sheet.Cells[row, 3].Value
                z = sheet.Cells[row, 4].Value
                elemid = sheet.Cells[row, 5].Value
                P = DB.XYZ(float(x),float(y),float(z))

                _dict[P] = [nummer,elemid]
                    
                
        book.Save()
        book.Dispose()
        fs.Close()
        fs.Dispose()
    except Exception as e:
        logger.error(e)
        book.Save()
        book.Dispose()
        fs.Close()
        fs.Dispose()
        script.exit()
    
    Liste = [['Alt','Neu','alt_id','ner_id']]
    liste1 = []
    for p in _dict.keys():
        nr = _dict[p][0]
        elemid = _dict[p][1]
        if not nr:
            nr = 'keine Nummer vergeben'

        cl = DB.FilteredElementCollector(doc,liste_elemid).WherePasses(DB.BoundingBoxContainsPointFilter(p)).ToElements()
        if cl.Count == 0:
            Liste.append([nr,'ausgenommen',elemid,''])
        else:
            for el in cl:
                nr1 = el.LookupParameter('IGF_RLT_BSK-Nummer').AsString()
                Liste.append([nr,nr1,elemid,el.Id.IntegerValue])
                liste1.append(el.Id.IntegerValue)
    
    for el in liste_elemid:
        if el.IntegerValue not in liste1:
            nr = doc.GetElement(el).LookupParameter('IGF_RLT_BSK-Nummer').AsString()
            Liste.append(['neu',nr,'',el.IntegerValue])
    
    path = r'C:\Users\zhang\Desktop\BSK_MFC_neu.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(Liste)):
        for col in range(4):
            worksheet.write(row,col, Liste[row][col])
    workbook.close()
    
# exportstand0()
   
exportstand1()
