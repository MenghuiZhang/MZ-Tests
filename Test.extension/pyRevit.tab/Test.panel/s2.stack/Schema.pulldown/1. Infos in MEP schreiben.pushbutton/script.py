# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import script
import os
from IGF_forms import ExcelSuche
from rpw import revit,DB
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
import xlsxwriter
from System.IO import FileInfo

__title__ = "1.2 abgleich test"
__doc__ = """

es gilt nur für die MEP-Räume in akuelle Ansiche des Schema-Modells.

Infos in MEP-Räume von TGA-Modell in Schema schreiben.

Vorgehensweise:

1. Ein MEP-Räume-Bauteilliste exportieren. sieht R:\pyRevit\10_Verknüpfung\02_Schema\Vorlage MEP-Räume
2. Exporteirte Excel anpassen (entsprechenden Parametername in 1. Zeile anpassen)
3. Entsprechende Ansicht (Bsp. Schema-Ansicht oder Bauteilliste) in Schema-Modell öffnen und das Skript durchlaufen.

[2022.08.02]
Version: 1.0
"""

__authors__ = "Menghui Zhang"


alt_Excel = r'C:\Users\zhang\Desktop\IGF_RLT - Luftmengen in MEP-Raum (IGF_3.1)_2023-11.12.xlsx'
neu_excel = r'C:\Users\zhang\Desktop\IGF_RLT - Luftmengen in MEP-Raum (IGF_3.1)_2023-12-14.xlsx'
abgleich = r'C:\Users\zhang\Desktop\IGF_RLT - Luftmengen in MEP-Raum (IGF_3.1)_2023-12-14_Abgleich.xlsx'


ExcelPackage.LicenseContext = LicenseContext.NonCommercial

Dict_alt = {}
fs = FileInfo(neu_excel)
book = ExcelPackage(fs)

try:
    for sheet in book.Workbook.Worksheets:
        maxRowNum = sheet.Dimension.End.Row
        for row in range(2, maxRowNum + 1):
            nummer = sheet.Cells[row, 2].Value
            if nummer == '' or  nummer == None:
                continue
            name = sheet.Cells[row, 3].Value
            flaeche = sheet.Cells[row, 4].Value
            volzu = sheet.Cells[row, 30].Value
            Dict_alt.setdefault(nummer,{}).setdefault(name,{}).setdefault(flaeche,[])
            Dict_alt[nummer][name][flaeche].append(volzu)

    book.Save()
    book.Dispose()

except Exception as e:
    book.Save()
    book.Dispose()
    print(e)

Dict_neu = {}
fs = FileInfo(abgleich)
book = ExcelPackage(fs)

try:
    for sheet in book.Workbook.Worksheets:
        maxRowNum = sheet.Dimension.End.Row
        for row in range(2, maxRowNum + 1):
            nummer = sheet.Cells[row, 1].Value
            if nummer == '' or  nummer == None:
                continue

            name = sheet.Cells[row, 3].Value
            flaeche = sheet.Cells[row, 5].Value
            volzu = Dict_alt[nummer][name][flaeche]
            if len(volzu) == 1:
                sheet.Cells[row, 17].Value = volzu[0]
            else:
                print(nummer,name)


    book.Save()
    book.Dispose()

except Exception as e:
    book.Save()
    book.Dispose()
    print(e)


# workbook = xlsxwriter.Workbook(abgleich)
# worksheet = workbook.add_worksheet()
# row_num = 0
# for nummer in Dict_neu.keys():
#     if nummer in Dict_alt.keys():
#         if len(Dict_neu[nummer]) == 1 and len(Dict_alt[nummer]) == 1:
#             worksheet.write(row_num, 0, nummer)
#             worksheet.write(row_num, 1, nummer)
#             worksheet.write(row_num, 2, Dict_neu[nummer][0][0])
#             worksheet.write(row_num, 3, Dict_alt[nummer][0][0])
#             worksheet.write(row_num, 4, Dict_neu[nummer][0][1])
#             worksheet.write(row_num, 5, Dict_alt[nummer][0][1])
#             worksheet.write(row_num, 6, Dict_neu[nummer][0][2])
#             worksheet.write(row_num, 7, Dict_alt[nummer][0][2])
#             worksheet.write(row_num, 8, Dict_neu[nummer][0][3])
#             worksheet.write(row_num, 9, Dict_alt[nummer][0][3])
#             worksheet.write(row_num, 10, Dict_neu[nummer][0][4])
#             worksheet.write(row_num, 11, Dict_alt[nummer][0][4])
#             worksheet.write(row_num, 12, Dict_neu[nummer][0][5])
#             worksheet.write(row_num, 13, Dict_alt[nummer][0][5])
#             worksheet.write(row_num, 14, Dict_neu[nummer][0][6])
#             worksheet.write(row_num, 15, Dict_alt[nummer][0][6])
#             row_num += 1
#         else:
#             for daten in Dict_neu[nummer]:
#                 worksheet.write(row_num, 0, nummer)
#                 worksheet.write(row_num, 2, daten[0])
#                 worksheet.write(row_num, 4, daten[1])
#                 worksheet.write(row_num, 6, daten[2])
#                 worksheet.write(row_num, 8, daten[3])
#                 worksheet.write(row_num, 10, daten[4])
#                 worksheet.write(row_num, 12, daten[5])
#                 worksheet.write(row_num, 14, daten[6])
#                 row_num += 1
            
#             for daten in Dict_alt[nummer]:
#                 worksheet.write(row_num, 0+1, nummer)
#                 worksheet.write(row_num, 2+1, daten[0])
#                 worksheet.write(row_num, 4+1, daten[1])
#                 worksheet.write(row_num, 6+1, daten[2])
#                 worksheet.write(row_num, 8+1, daten[3])
#                 worksheet.write(row_num, 10+1, daten[4])
#                 worksheet.write(row_num, 12+1, daten[5])
#                 worksheet.write(row_num, 14+1, daten[6])
#                 row_num += 1
#     else:
#         for daten in Dict_neu[nummer]:
#             worksheet.write(row_num, 0, nummer)
#             worksheet.write(row_num, 2, daten[0])
#             worksheet.write(row_num, 4, daten[1])
#             worksheet.write(row_num, 6, daten[2])
#             worksheet.write(row_num, 8, daten[3])
#             worksheet.write(row_num, 10, daten[4])
#             worksheet.write(row_num, 12, daten[5])
#             worksheet.write(row_num, 14, daten[6])
#             row_num += 1

# for nummer in Dict_alt.keys():
#     if nummer not in Dict_neu.keys():
#         for daten in Dict_alt[nummer]:
#             worksheet.write(row_num, 0+1, nummer)
#             worksheet.write(row_num, 2+1, daten[0])
#             worksheet.write(row_num, 4+1, daten[1])
#             worksheet.write(row_num, 6+1, daten[2])
#             worksheet.write(row_num, 8+1, daten[3])
#             worksheet.write(row_num, 10+1, daten[4])
#             worksheet.write(row_num, 12+1, daten[5])
#             worksheet.write(row_num, 14+1, daten[6])
#             row_num += 1


# workbook.close()
