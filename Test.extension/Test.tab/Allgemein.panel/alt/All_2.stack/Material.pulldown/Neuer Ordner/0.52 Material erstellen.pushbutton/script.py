# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Visual import AppearanceAssetEditScope
import struct
import rpw
import time
import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

start = time.time()

__title__ = "0.52 Material erstellen"
__doc__ = """Material erstellen"""
__author__ = "Menghui Zhang"

from pyIGF_logInfo import getlog
getlog(__title__)


exapp = Excel.ApplicationClass()
logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp

Material_collector = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Material))
Material_dict = {}

for ele in Material_collector:
    Material_dict[ele.Name] = ele

def ExcelLesen(inExcelPath):
    outDict = {}
    book = exapp.Workbooks.Open(inExcelPath)

    for sheet in book.Worksheets:
        groupName = sheet.Name
        rows = sheet.UsedRange.Rows.Count
        if groupName == 'Familien':
            with forms.ProgressBar(title='{value}/{max_value} Daten Lesen',cancellable=True, step=10) as pb:
                n_1 = 0
                for row in range(2, rows+1):
                    n_1 += 1
                    if pb.cancelled:
                        script.exit()
                        book.Save()
                        book.Close()
                    pb.update_progress(n_1, rows-1)
                    original = sheet.UsedRange.Cells[row, 3].Value2
                    Name = sheet.UsedRange.Cells[row, 4].Value2
                    Farbe = sheet.UsedRange.Cells[row,6].Value2
                    R = sheet.UsedRange.Cells[row,7].Value2
                    G = sheet.UsedRange.Cells[row,8].Value2
                    B = sheet.UsedRange.Cells[row,9].Value2
                    if Farbe:
                        outDict[Name] = [original,R,G,B]
                        logger.info('{}:{}'.format(Name,[original,R,G,B]))
    book.Save()
    book.Close()

    return outDict

def MaterialErstellen(inMaterial,newMaterial):
    tran = Transaction(doc,'Material erstellt')
    tran.Start()
    with forms.ProgressBar(title='{value}/{max_value} Material',cancellable=True, step=1) as pb:
        n_1 = 0
        editScope = AppearanceAssetEditScope(doc)
        for el in newMaterial.keys():
            n_1 += 1
            if pb.cancelled:
                tran.RollBack()
                script.exit()
            pb.update_progress(n_1, len(newMaterial))
            if not el in inMaterial.keys():
                logger.info('{} erstellt'.format(el))
                id = DB.Material.Create(doc,el)
                mate = doc.GetElement(id)
                R = int(newMaterial[el][1])
                G = int(newMaterial[el][2])
                B = int(newMaterial[el][3])
                farbe = DB.Color(System.Byte(R),System.Byte(G),System.Byte(B))
                mate.Color = farbe
                mate.Transparency = 0
                mate.SurfaceForegroundPatternColor = farbe
                mate.UseRenderAppearanceForShading = False
                aa = DB.AppearanceAssetElement.GetAppearanceAssetElementByName(doc, 'Generisch(1885)').Duplicate(el)
                TherId = None
                StruId = None
                if newMaterial[el][0] and newMaterial[el][0] in inMaterial.keys():
                    TherId = inMaterial[newMaterial[el][0]].ThermalAssetId
                    StruId = inMaterial[newMaterial[el][0]].StructuralAssetId

                editableAsset = editScope.Start(aa.Id)
                asset = aa.GetRenderingAsset()
                dict = {}
                for n in range(asset.Size):
                    item = asset[n]
                    try:
                        Name = item.Name
                        dict[Name] = n

                    except Exception as e:
                        pass
                try:
                    editableAsset[dict["generic_is_metal"]].Value = False
                    editableAsset[dict["generic_reflectivity_at_0deg"]].Value = 0
                    editableAsset[dict["generic_reflectivity_at_90deg"]].Value = 0
                    editableAsset[dict["generic_transparency_image_fade"]].Value = 1
                    editableAsset[dict["generic_diffuse_image_fade"]].Value = 0
                    editableAsset[dict["generic_transparency"]].Value = 0
                    editableAsset[dict["generic_glossiness"]].Value = 0
                    editableAsset[dict["generic_self_illum_filter_map"]].SetValueAsColor(farbe)
                    editableAsset[dict["generic_diffuse"]].SetValueAsColor(farbe)

                except Exception as e:
                    logger.error(e)

                editScope.Commit(True)
                mate.AppearanceAssetId = aa.Id
                if TherId:
                    mate.ThermalAssetId = TherId
                    mate.StructuralAssetId = StruId
                logger.info(30*'--')
    tran.Commit()


excelPath1  = rpw.ui.forms.TextInput('Excel: ', default ='R:\\Vorlagen\\_IGF\\_Material\\IGF_Material.xlsx')
Daten = ExcelLesen(excelPath1)
MaterialErstellen(Material_dict,Daten)

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
