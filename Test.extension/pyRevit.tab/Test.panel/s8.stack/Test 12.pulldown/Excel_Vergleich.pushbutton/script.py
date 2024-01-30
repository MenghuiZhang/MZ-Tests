# coding: utf8

from IGF_Klasse._labormedien import LaboranschlussType,LaborMedien_MEPRaum,LaborMedien_Heinekamp,LaborMedien_IGF
from IGF_Klasse._familie_typ import FamilieTypFactory
from rpw import revit,DB
import xlsxwriter
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
from IGF_Funktionen._Parameter import wert_schreibenbase
from System.IO import FileInfo

# from System.Windows.Input import Key


__title__ = "Raumluft GC importieren"
__doc__ = """

zwischen zwei Formteilen oder Zubehör ein Kanal-/Rohrstück in Länge X einfügen


[2023.03.30]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

# excel = r'C:\Users\zhang\Desktop\Raumliste GC.xlsx'
# doc = revit.doc

# Parameter_Dict = {}
# ExcelPackage.LicenseContext = LicenseContext.NonCommercial
# # fs = FileStream(excel,FileMode.Open,FileAccess.Read)
# fs = FileInfo(excel)
# book = ExcelPackage(fs)
# Parameter_Liste = []
# Bauteilnummer = ''
# try:
#     sheet = book.Workbook.Worksheets['Tabelle1']
#     # for sheet in book.Workbook.Worksheets:
#     maxColumnNum = sheet.Dimension.End.Column
#     maxRowNum = sheet.Dimension.End.Row
#     for row in range(2, maxRowNum + 1):
#         ebene = sheet.Cells[row, 1].Value
#         nummer = sheet.Cells[row, 2].Value
#         name = sheet.Cells[row, 3].Value
#         Personen = sheet.Cells[row, 4].Value
#         Flaeche = sheet.Cells[row, 5].Value
#         Hoehe = sheet.Cells[row, 6].Value
#         berechnung = sheet.Cells[row, 7].Value
#         faktor = sheet.Cells[row, 8].Value
#         reduzierung = sheet.Cells[row, 9].Value
#         abluft = sheet.Cells[row, 10].Value
#         zuluft = sheet.Cells[row, 11].Value    
#         abluftanlagen = sheet.Cells[row, 11].Value         
#         zuluftanlagen = sheet.Cells[row, 11].Value     
#         schacht = sheet.Cells[row, 11].Value    
#         Parameter_Dict.setdefault(ebene,{}).setdefault(nummer,{}).setdefault(name,[]) 
#         Parameter_Dict[ebene][nummer][name] = [Personen,Flaeche,Hoehe,berechnung,faktor,reduzierung,abluft,zuluft,abluftanlagen,zuluftanlagen,schacht]

#     book.Save()
#     book.Dispose()

# except Exception as e:
#     book.Save()
#     book.Dispose()

# berechnung_nach = {
#                         "Fläche":'1',
#                         "Luftwechsel":'2',
#                         "Personenbezogen":'3',
#                         "Manuell":'4',
#                         "NurZuMa":'5',
#                         "objektbezogen":'6',
#                         "nurZU_Fläche":'5.1',
#                         "nurZU_Luftwechsel":'5.2',
#                         "nurZU_Person":'5.3',
#                         "nurAB_Fläche":'6.1',
#                         "nurAB_Luftwechsel":'6.2',
#                         "nurAB_Person":'6.3',
#                         "Pers_u_Fläche2_5qm":'7',
#                         "keine":'9',
#                     }
    
# spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
# t = DB.Transaction(doc,'Raumluft GC')
# t.Start()
# for el in spaces:
#     ebene = 'Ebene '+el.Level.Name
#     nummer = el.Number
#     name = el.LookupParameter('Name').AsString()
#     result = False
#     if ebene in Parameter_Dict.keys():
#         if nummer in Parameter_Dict[ebene].keys():
#             if name in Parameter_Dict[ebene][nummer].keys():
#                 wert_schreibenbase(el.LookupParameter('Personenanzahl'),0)
#                 wert_schreibenbase(el.LookupParameter('Personenzahl'),Parameter_Dict[ebene][nummer][name][0])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_istReduziert'),1)
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_Nachtbetrieb'),0)
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_RaumFläche'),Parameter_Dict[ebene][nummer][name][1])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_RaumHöhe'),float(Parameter_Dict[ebene][nummer][name][2]) * 1000)
#                 wert_schreibenbase(el.LookupParameter('TGF_RLT_VolumenstromProName'),Parameter_Dict[ebene][nummer][name][3])
#                 wert_schreibenbase(el.LookupParameter('TGF_RLT_VolumenstromProNummer'),berechnung_nach[Parameter_Dict[ebene][nummer][name][3]])
#                 wert_schreibenbase(el.LookupParameter('TGF_RLT_VolumenstromProFaktor'),Parameter_Dict[ebene][nummer][name][4])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_Raum-ReduziertFaktor'),Parameter_Dict[ebene][nummer][name][5])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_AbluftminRaum'),Parameter_Dict[ebene][nummer][name][6])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_ZuluftminRaum'),Parameter_Dict[ebene][nummer][name][7])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_AbluftmaxRaum'),Parameter_Dict[ebene][nummer][name][6])
#                 wert_schreibenbase(el.LookupParameter('IGF_RLT_ZuluftmaxRaum'),Parameter_Dict[ebene][nummer][name][7])
#                 wert_schreibenbase(el.LookupParameter('TGA_RLT_AnlagenNrAbluft'),Parameter_Dict[ebene][nummer][name][8])
#                 wert_schreibenbase(el.LookupParameter('TGA_RLT_AnlagenNrZuluft'),Parameter_Dict[ebene][nummer][name][9])
#                 wert_schreibenbase(el.LookupParameter('TGA_RLT_SchachtZuluft'),Parameter_Dict[ebene][nummer][name][10])
#                 wert_schreibenbase(el.LookupParameter('TGA_RLT_SchachtAbluft'),Parameter_Dict[ebene][nummer][name][10])

#                 result = True
#             else:
#                 name0 = name
#                 if len(Parameter_Dict[ebene][nummer]) == 1:
#                     for name in Parameter_Dict[ebene][nummer].keys():
#                         print(name,name0)
#                         wert_schreibenbase(el.LookupParameter('Personenanzahl'),0)
#                         wert_schreibenbase(el.LookupParameter('Personenzahl'),Parameter_Dict[ebene][nummer][name][0])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_istReduziert'),1)
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_Nachtbetrieb'),0)
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_RaumFläche'),Parameter_Dict[ebene][nummer][name][1])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_RaumHöhe'),float(Parameter_Dict[ebene][nummer][name][2]) * 1000)
#                         wert_schreibenbase(el.LookupParameter('TGF_RLT_VolumenstromProName'),Parameter_Dict[ebene][nummer][name][3])
#                         wert_schreibenbase(el.LookupParameter('TGF_RLT_VolumenstromProNummer'),berechnung_nach[Parameter_Dict[ebene][nummer][name][3]])
#                         wert_schreibenbase(el.LookupParameter('TGF_RLT_VolumenstromProFaktor'),Parameter_Dict[ebene][nummer][name][4])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_Raum-ReduziertFaktor'),Parameter_Dict[ebene][nummer][name][5])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_AbluftminRaum'),Parameter_Dict[ebene][nummer][name][6])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_ZuluftminRaum'),Parameter_Dict[ebene][nummer][name][7])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_AbluftmaxRaum'),Parameter_Dict[ebene][nummer][name][6])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_ZuluftmaxRaum'),Parameter_Dict[ebene][nummer][name][7])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_AnlagenNr_RAB'),Parameter_Dict[ebene][nummer][name][8])
#                         wert_schreibenbase(el.LookupParameter('IGF_RLT_AnlagenNr_RZU'),Parameter_Dict[ebene][nummer][name][9])
#                         wert_schreibenbase(el.LookupParameter('TGA_RLT_SchachtZuluft'),Parameter_Dict[ebene][nummer][name][10])
#                         wert_schreibenbase(el.LookupParameter('TGA_RLT_SchachtAbluft'),Parameter_Dict[ebene][nummer][name][10])

#                 result = True

#     if not result:
#         print(el.Id.ToString(),ebene,nummer,name)

            
# t.Commit()



excel = r'C:\Users\zhang\Desktop\Text_Nummer.xlsx'
doc = revit.doc

Parameter_Dict = {}
ExcelPackage.LicenseContext = LicenseContext.NonCommercial
# fs = FileStream(excel,FileMode.Open,FileAccess.Read)
fs = FileInfo(excel)
book = ExcelPackage(fs)
_Dict0 = {}
_Dict1 = {}
Parameter_Liste = []
Bauteilnummer = ''
try:
    # sheet = book.Workbook.Worksheets['000 Raumnummer Planung']
    for sheet in book.Workbook.Worksheets:
        if sheet.Name == 'T1':
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                alznummer = sheet.Cells[row, 1].Value
                
                neunummer = sheet.Cells[row, 2].Value
                if not alznummer:continue
                name = sheet.Cells[row, 3].Value
                # neunummer = sheet.Cells[row, 6].Value    
                if  alznummer not in  _Dict0.keys():
                    _Dict0.setdefault(alznummer,{}).setdefault(neunummer,"")
                else:
                    print(alznummer)
                _Dict0[alznummer][neunummer] = name
        elif sheet.Name == 'T2':
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                alznummer = sheet.Cells[row, 1].Value
                
                neunummer = sheet.Cells[row, 2].Value
                if not alznummer:continue
                name = sheet.Cells[row, 3].Value
                # neunummer = sheet.Cells[row, 6].Value                
                if  alznummer not in  _Dict1.keys():
                    _Dict1.setdefault(alznummer,{}).setdefault(neunummer,"")
                else:
                    print(alznummer)
                _Dict1[alznummer][neunummer] = name

    book.Save()
    book.Dispose()

except Exception as e:
    print(e)
    book.Save()
    book.Dispose()


excel1 = r'C:\Users\zhang\Desktop\Text_Nummer1.xlsx'
Liste = []
for nummer0 in _Dict0.keys():
    for nummer1 in _Dict0[nummer0].keys():
        name = _Dict0[nummer0][nummer1]
        if nummer0 in _Dict1.keys():
            if nummer1 in _Dict1[nummer0].keys():
                name1 = _Dict1[nummer0][nummer1]
                if name.replace(' ','') == name1.replace(' ',''):
                    pass
                else:
                    Liste.append([nummer0,nummer1,name,nummer0,nummer1,name1,'1'])
            else:
                for nummer2 in  _Dict1[nummer0].keys():
                    name1 = _Dict1[nummer0][nummer2]
                Liste.append([nummer0,nummer1,name,nummer0,nummer2,name1,'2'])
        else:
            Liste.append([nummer0,nummer1,name,'','','','3'])
Liste.append(['','','','','',''])
for nummer0 in _Dict1.keys():
    for nummer1 in _Dict1[nummer0].keys():
        name = _Dict1[nummer0][nummer1]
        if nummer0 in _Dict0.keys():
            if nummer1 in _Dict0[nummer0].keys():
                name1 = _Dict0[nummer0][nummer1]
                if name == name1:
                    pass
                else:
                    pass
                    #Liste.append([nummer0,nummer1,name1,nummer0,nummer1,name,'1'])
            else:
                for nummer2 in _Dict0[nummer0].keys():
                    name1 = _Dict0[nummer0][nummer2]
                Liste.append([nummer0,nummer2,name1,nummer0,nummer1,name,'2'])
        else:
            Liste.append(['','','',nummer0,nummer1,name,'3'])

print(Liste)

# Create a new Excel file
workbook = xlsxwriter.Workbook(excel1)

# Add a worksheet
worksheet = workbook.add_worksheet()

# Define some data to be written

# Write the data to the worksheet
for row_num, row_data in enumerate(Liste):
    worksheet.write_row(row_num, 0, row_data)

# Close the workbook
workbook.close()
