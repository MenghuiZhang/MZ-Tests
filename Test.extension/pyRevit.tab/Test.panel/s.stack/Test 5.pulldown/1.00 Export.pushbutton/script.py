# coding: utf8
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms
import xlsxwriter

__title__ = "1.30 Export"
__doc__ = """schreibt Raumzustandsdaten von Architektur Modell in Zustandsparameter des Berechnungsmodells.
Bezug auf Volumen, Umfang, Licht Höhe, Fläche.

Kategorie: Räume

[2021.11.25]
Version: 1.1
"""

__author__ = "Menghui Zhang"

logger = script.get_logger()
uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass


spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()

if not spaces:
    logger.error("Keine MEP-Räume in aktueller Projekt gefunden")
    script.exit()

revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el

revitLinks_collector.Dispose()


rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()
rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()

raums = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElementIds()


class MEPRaum:
    def __init__(self,elemid = None):
        self.elemid = elemid
        if self.elemid != None:
            self.elem = doc.GetElement(self.elemid)
            self.Name = get_value(self.elem.LookupParameter('Name'))
            self.nummer = self.elem.Number

            self.Flaeche = get_value(self.elem.LookupParameter('Fläche'))
            self.Umfang = get_value(self.elem.LookupParameter('Umfang'))
            self.Hoehe = get_value(self.elem.LookupParameter('Lichte Höhe'))
            self.Volumen = get_value(self.elem.LookupParameter('Volumen'))
            self.level = self.elem.LookupParameter('Ebene').AsValueString()
        else:
            self.Name = ''
            self.nummer = ''

            self.Flaeche = 0.01
            self.Umfang = 0.01
            self.Hoehe = 0.01
            self.Volumen = 0.01
            self.level = None

class Raum:
    def __init__(self,elemid = None):
        self.elemid = elemid
        if self.elemid != None:
            self.elem = rvtdoc.GetElement(self.elemid)
            self.Name = get_value(self.elem.LookupParameter('Name'))
            self.nummer = self.elem.Number

            self.Flaeche = get_value(self.elem.LookupParameter('Fläche'))
            self.Umfang = get_value(self.elem.LookupParameter('Umfang'))
            self.Hoehe = get_value(self.elem.LookupParameter('Lichte Höhe'))
            self.Volumen = get_value(self.elem.LookupParameter('Volumen'))
            self.level = self.elem.LookupParameter('Ebene').AsValueString()
        else:
            self.Name = ''
            self.nummer = ''

            self.Flaeche = 0
            self.Umfang = 0
            self.Hoehe = 0
            self.Volumen = 0
            self.level = None


DICT_MEP = {}
DICT_RAUM = {}
DICT_Alle = {}

for el in spaces:
    space = MEPRaum(el)
    DICT_MEP[space.nummer] = space

for el in raums:
    raum = Raum(el)
    DICT_RAUM[raum.nummer] = raum

for el in DICT_MEP.keys():
    if el in DICT_RAUM.keys():
        DICT_Alle[el] = [DICT_MEP[el],DICT_RAUM[el]]
    else:
        DICT_Alle[el] = [DICT_MEP[el],Raum()]

for el in DICT_RAUM.keys():
    if el in DICT_Alle.keys():
        continue
    DICT_Alle[el] = [MEPRaum(),DICT_RAUM[el]]

path = r'C:\Users\Zhang\Desktop\Abgleich.xlsx'
workbook = xlsxwriter.Workbook(path)
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format()
cell_format.set_pattern(1)  # This is optional when using a solid fill.
cell_format.set_bg_color('gray')
worksheet.write(0, 0, 'Nummer')
worksheet.write(0, 1, 'Name_A')
worksheet.write(0, 2, 'Name_MEP')
worksheet.write(0, 3, 'Fläche_A (m²)')
worksheet.write(0, 4, 'Fläche_MEP (m²)')
worksheet.write(0, 5, 'Umfang_A (m)')
worksheet.write(0, 6, 'Umfang_MEP (m)')
worksheet.write(0, 7, 'Höhe_A (m)')
worksheet.write(0, 8, 'Höhe_MEP (m)')
worksheet.write(0, 9, 'Volumen_A (m³)')
worksheet.write(0, 10, 'Volumen_MEP (m³)')
worksheet.write(0, 11, 'Abgleich')
worksheet.write(0, 12, 'ElementId')

nummer_liste = DICT_Alle.keys()[:]
nummer_liste.sort()
row = 0
for nnn in range(1,len(nummer_liste)+1):
    nummer = nummer_liste[nnn-1]

    el = DICT_Alle[nummer]
    mep = el[0]
    raum = el[1]
    if nummer.find('.0') != -1:
        continue
    if nummer.find('070.') != -1:
        continue
    if (not raum.level) and (not mep.level):
        continue
    
    abgleich = ''
    if raum.Name == '':
        abgleich = 'nur in TGA-Modell'
    elif mep.Name == '':
        abgleich = 'nur in Architektur-Modell'
    else:
        if raum.Name != mep.Name:
            abgleich += 'Name passt nicht. '
        if mep.Flaeche < 0.001:
            mep.Flaeche = 0.1
        if abs(float(raum.Flaeche)/float(mep.Flaeche)-1) > 0.05:
            abgleich += 'Fläche passt nicht. '
        if mep.Umfang < 0.001:
            mep.Umfang = 0.1
        if abs(float(raum.Umfang)/float(mep.Umfang)*1000-1) > 0.05:
            abgleich += 'Umfang passt nicht. '
        # if int(round(raum.Hoehe)) != int(round(mep.Hoehe)):
        #     abgleich += 'Höhe passt nicht. '
        # if int(round(raum.Volumen)) != int(round(mep.Volumen)):
        #     abgleich += 'Volumen passt nicht. '
    if abgleich == '':
        continue
    row += 1
    worksheet.write(row, 11, abgleich)
    
    worksheet.write(row, 1, raum.Name)
    worksheet.write(row, 2, mep.Name,cell_format)
    worksheet.write(row, 3, round(raum.Flaeche,2))
    worksheet.write(row, 4, round(mep.Flaeche,2),cell_format)
    worksheet.write(row, 5, round(raum.Umfang,2))
    worksheet.write(row, 6, round(mep.Umfang/1000,2),cell_format)
    worksheet.write(row, 7, round(raum.Hoehe,2))
    worksheet.write(row, 8, round(mep.Hoehe/1000,2),cell_format)
    worksheet.write(row, 9, round(raum.Volumen,2))
    worksheet.write(row, 10,round(mep.Volumen,2),cell_format)    
    if mep.elemid != None:
        worksheet.write(row, 12, mep.elemid.ToString())
    else:
         worksheet.write(row, 12, '')
  
    worksheet.write(row, 0, nummer)
    
worksheet.autofilter(0, 0, row-1, 12)
workbook.close()