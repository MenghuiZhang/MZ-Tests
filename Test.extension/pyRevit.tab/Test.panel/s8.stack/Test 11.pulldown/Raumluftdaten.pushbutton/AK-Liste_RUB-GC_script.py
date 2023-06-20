# coding: utf8
import os
from pyrevit import revit, DB
from pyrevit import script
from IGF_forms import ExcelSuche
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
from IGF_lib import get_value
__title__ = "Raumluftdaten übernehmen GC"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

adresse = 'Excel Adresse'

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.ShowDialog()
# try:
#     config.set_value('adresse', ExcelWPF.Adresse.Text)
#     config.save_config()
# except:
#     logger.error('kein Excel gegeben')
#     script.exit()


dict_raumdaten = {}

class Berechnung:
    def __init__(self,nummer,name,einheit):
        self.nummer = nummer
        self.name = name
        self.einheit = einheit

dict_berechnung = {
    "Fläche": Berechnung(1, 'Fläche', 'm³/h pro m²'),
    "Luftwechsel": Berechnung(2, 'Luftwechsel', '-1/h'),
    "Personenbezogen": Berechnung(3, 'Fläche', 'm³/h pro P'),
    "Manuell": Berechnung(4, 'manuell', 'm³/h'),
    "NurZuMa": Berechnung(5, 'nurZUMa', 'm³/h'),
    "objektbezogen": Berechnung(6, 'nurABMa', 'm³/h'),
    "nurZU_Fläche": Berechnung(5.1, 'nurZU_Fläche', 'm³/h pro m²'),
    "nurZU_Luftwechsel": Berechnung(5.2, 'nurZU_Luftwechsel', '-1/h'),
    "nurZU_Person": Berechnung(5.3, 'nurZU_Person', 'm³/h pro P'),
    "nurAB_Fläche": Berechnung(6.1, 'nurAB_Fläche', 'm³/h pro m²'),
    "nurAB_Luftwechsel": Berechnung(6.2, 'nurAB_Luftwechsel', '-1/h'),
    "nurAB_Person": Berechnung(6.3, 'nurAB_Person', 'm³/h pro P'),
    "Pers_u_Fläche2_5qm": Berechnung(7, 'Pers_u_Fläche2_5qm', 'm³/h pro P + 2.5 * m³/h pro m²'),
    "keine": Berechnung(9, 'keine', ''),
}

class MEPRaumdaten:
    def __init__(self):
        self.nummer = None
        self.name = None
        self.raumtyp = None
        self.personen = None
        self.berechnungsname = None
        self.berechnungsnummer = None
        self.berechnungsfaktor = None
        self.nachtbetrieb = None
        self.reduzierungsfaktor = None
        self.zuluftmin = None
        self.abluftmin = None
        self.zuluftnacht = None
        self.abluftnacht = None
        self.anlagenzu = None
        self.anlagenab = None
        self.schahct = None
        self.flaeche = None
        self.volumen = None
    
    def wert_schreiben(self,mepraum):
        mepraum.LookupParameter('TGA_RLT_VolumenstromProNummer').Set(dict_berechnung[self.berechnungsname].nummer)
        mepraum.LookupParameter('TGA_RLT_VolumenstromProName').Set(dict_berechnung[self.berechnungsname].name)
        mepraum.LookupParameter('TGA_RLT_VolumenstromProEinheit').Set(dict_berechnung[self.berechnungsname].einheit)
        mepraum.LookupParameter('IGF_RLT_Nachtbetrieb').Set(int(self.nachtbetrieb))
        mepraum.LookupParameter('TGA_RLT_VolumenstromProFaktor').Set(float(self.berechnungsfaktor))
        if mepraum.LookupParameter('Personenzahl').IsReadOnly is False:
            mepraum.LookupParameter('Personenzahl').Set(float(self.personen))
        mepraum.LookupParameter('IGF_RLT_Raum-ReduziertFaktor').Set(float(self.reduzierungsfaktor))
        mepraum.LookupParameter('IGF_RLT_ZuluftminRaum').SetValueString(str(self.zuluftmin))
        mepraum.LookupParameter('IGF_RLT_AbluftminRaum').SetValueString(str(self.abluftmin))
        mepraum.LookupParameter('IGF_RLT_ZuluftNachtRaum').SetValueString(str(self.zuluftnacht))
        mepraum.LookupParameter('IGF_RLT_AbluftNachtRaum').SetValueString(str(self.abluftnacht))
        mepraum.LookupParameter('TGA_RLT_AnlagenNrZuluft').Set(str(self.anlagenzu))
        mepraum.LookupParameter('TGA_RLT_AnlagenNrAbluft').Set(str(self.anlagenab))
        mepraum.LookupParameter('TGA_RLT_SchachtZuluft').Set(str(self.schahct))
        mepraum.LookupParameter('TGA_RLT_SchachtAbluft').Set(str(self.schahct))


    def pruefen(self,mepraum):
        return abs(self.flaeche -  get_value( mepraum.LookupParameter('Fläche')))/ self.flaeche < 0.05  #or self.name == mepraum.LookupParameter('Name').AsString()
        #or abs(self.volumen -  get_value( mepraum.LookupParameter('Volumen'))) / self.volumen < 0.05
    def pruefen1(self,mepraum):
        return (abs(self.flaeche -  get_value( mepraum.LookupParameter('Fläche')))/ self.flaeche < 0.05 ) or self.name == mepraum.LookupParameter('Name').AsString()
        #or abs(self.volumen -  get_value( mepraum.LookupParameter('Volumen'))) / self.volumen < 0.05
    def exportliste(self,mepraum):
        return [
            self.nummer,
            self.name,
            mepraum.LookupParameter('Name').AsString(),
            self.raumtyp,
            mepraum.LookupParameter('Flächentyp').AsValueString(),
            self.flaeche,
            get_value( mepraum.LookupParameter('Fläche')),
            self.volumen,
            get_value( mepraum.LookupParameter('Volumen'))

        ]
    @staticmethod
    def exportliste1(mepraum):
        return [
            mepraum.Number,
            '',
            mepraum.LookupParameter('Name').AsString(),
            '',
            mepraum.LookupParameter('Flächentyp').AsValueString(),
            '',
            get_value( mepraum.LookupParameter('Fläche')),
            '',
            get_value( mepraum.LookupParameter('Volumen'))

        ]
        # return self.name == mepraum.LookupParameter('Name').AsString() and self.raumtyp == mepraum.LookupParameter('Flächentyp').AsValueString()


def Import_HK_MEP():
    path = ExcelWPF.Adresse.Text
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(path,FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)
    try:
        for sheet in book.Workbook.Worksheets:
            maxRowNum = sheet.Dimension.End.Row
            for row in range(19, maxRowNum + 1):
                mepraum = MEPRaumdaten()
                mepraum.nummer = sheet.Cells[row, 5].Value
                mepraum.name = sheet.Cells[row, 6].Value
                mepraum.raumtyp = sheet.Cells[row, 7].Value
                mepraum.personen = sheet.Cells[row, 8].Value
                mepraum.flaeche = sheet.Cells[row, 9].Value
                mepraum.volumen = sheet.Cells[row, 12].Value
                mepraum.berechnungsname = sheet.Cells[row, 13].Value
                mepraum.berechnungsfaktor = sheet.Cells[row, 14].Value
                mepraum.nachtbetrieb = sheet.Cells[row, 15].Value
                mepraum.reduzierungsfaktor = sheet.Cells[row, 16].Value
                mepraum.zuluftmin = sheet.Cells[row, 20].Value
                mepraum.abluftmin = sheet.Cells[row, 19].Value
                mepraum.zuluftnacht = sheet.Cells[row, 22].Value
                mepraum.abluftnacht = sheet.Cells[row, 21].Value
                mepraum.anlagenzu = sheet.Cells[row, 24].Value
                mepraum.anlagenab = sheet.Cells[row, 23].Value
                mepraum.schahct = sheet.Cells[row, 25].Value
                

                if mepraum.nummer not in dict_raumdaten.keys():
                    dict_raumdaten[mepraum.nummer] = [mepraum]
                else:
                    dict_raumdaten[mepraum.nummer].append(mepraum)

                # if mepraum.nummer+ mepraum.name not in dict_raumdaten.keys():
                #     dict_raumdaten[mepraum.nummer + mepraum.name] = mepraum
                # else:
                #     print(mepraum.nummer + ' - ' + mepraum.nummer)

                
                
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
Import_HK_MEP()

liste = []
spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
n = 0
t = DB.Transaction(doc,'daten übertragen')
t.Start()
for space in spaces:
    nummer = space.Number
    name = space.LookupParameter('Name').AsString()
    if nummer in dict_raumdaten.keys():
        anzahl = dict_raumdaten[nummer]
        result = False
        liste_neu = []
        for mep in dict_raumdaten[nummer]:
            # mep = dict_raumdaten[nummer]
            if anzahl ==1:

                if mep.pruefen(space):
                    mep.wert_schreiben(space)
                else:
                    n+=1
                    liste.append(mep.exportliste(space))
                    logger.error('überprüfen {}'.format(nummer+'-'+name))
            else:
                if mep.pruefen(space):
                    mep.wert_schreiben(space)
                else:
                    n+=1
                    liste_neu.append(mep.exportliste(space))
                    logger.warning('überprüfen {}'.format(nummer+'-'+name))
        if len(liste_neu) == 2:
            liste.extend(liste_neu)
    # if nummer+name in dict_raumdaten.keys():
    #     mep = dict_raumdaten[nummer+name]
    #     if mep.pruefen(space):
    #         mep.wert_schreiben(space)
    #     else:
    #         n+=1
    #         logger.error('überprüfen {}'.format(nummer+'-'+name))
    else:
        n+=1
        liste.append(MEPRaumdaten.exportliste1(space))
        print('Raumnummer {} nicht gefunden'.format(nummer+'-'+name))
t.Commit()

liste.sort()
liste.insert(0, ['Raumnummer','Raumname in Raumbuch','Raumname in Revit','Raumtyp in Raumbuch','Raumtyp in Revit','Fläche in Raumbuch','Fläche in Revit','Volumen in Raumbuch','Volumen in Revit'])
import xlsxwriter
path = r"C:\Users\zhang\Desktop\RUB GC Abgleich.xlsx"
workbook = xlsxwriter.Workbook(path)
worksheet = workbook.add_worksheet()

for row in range(len(liste)):
    for col in range(len(liste[0])):
        try:worksheet.write(row, col, liste[row][col])
        except:pass

workbook.close()
