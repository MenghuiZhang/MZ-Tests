import sys
import clr
import os
import xlsxwriter
from getpass import getuser
from time import strftime, localtime
doc = __revit__.ActiveUIDocument.Document
# sys.path.Add(os.path.dirname(__file__))
sys.path.Add(os.path.dirname(__file__)+'\\net40')
clr.AddReference('NPOI')
clr.AddReference('NPOI.OOXML')
import NPOI as np
from System.IO import FileStream,FileMode,FileAccess,FileShare

def get_cell(row,column):
    cell = row.GetCell(column)
    if cell:return cell
    else:
        cell = row.CreateCell(column)
        return cell

def getlog1(Programmname):
    path = r'R:\pyRevit\xx_Skripte\Historie.xlsx'
    fs = FileStream(path,FileMode.Open,FileAccess.Read)
    book1 = np.XSSF.UserModel.XSSFWorkbook(fs)
    sheet = book1.GetSheetAt(0)
    # exapp = ApplicationClass()
    # exapp.Visible = False
    
    if not os.path.isfile(path):
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet()
        ueberschrift = ['Zeit', 'Benutzername', 'Name', 'Number', 'ClientName', 'Cloud-Modell',  "Programm"]
        for col in range(len(ueberschrift)):
            worksheet.write(0, col, ueberschrift[col])
        workbook.close()

    benutzername = getuser()
    zeit = strftime("%d.%m.%Y-%H:%M", localtime())
    name = doc.ProjectInformation.Name
    number = doc.ProjectInformation.Number
    ClientName = doc.ProjectInformation.ClientName
    cloud = str(doc.IsModelInCloud)
    logdaten = [zeit, benutzername, name, number, ClientName,cloud, Programmname]
    rows = sheet.LastRowNum
    for col in range(len(logdaten)):
        cell = get_cell(rows+1,col)
        cell.SetCellValue(logdaten[col])
    fs = FileStream(path, FileMode.Create, FileAccess.Write)
    book1.Write(fs)
    book1.Close()
    fs.Close()
    # cell_0 = get_cell(e_param.row,0)
    #                         cell_1 = get_cell(e_param.row,1)
    #                         cell_0.SetCellValue('')
    #                         cell_1.SetCellValue('Projektparameter')
    # book = exapp.Workbooks.Open(path)
    # sheet = book.Worksheets[1]
    # rows = sheet.UsedRange.Rows.Count
    # try:
    #     rows += 1
    #     for col in range(1,len(logdaten)+1):
    #         sheet.Cells[rows,col] = logdaten[col-1]
        
    #     book.Save()
    #     book.Close()
    #     Marshal.FinalReleaseComObject(sheet)
    #     Marshal.FinalReleaseComObject(book)
    #     exapp.Quit()
    #     Marshal.FinalReleaseComObject(exapp)
    # except Exception as e:
    #     book.Save()
    #     book.Close()
    #     Marshal.FinalReleaseComObject(sheet)
    #     Marshal.FinalReleaseComObject(book)
    #     exapp.Quit()
    #     Marshal.FinalReleaseComObject(exapp)


# from excel import _NPOI
# excelPath = r'C:\Users\Zhang\Desktop\IGF_Parameter_RUB-GC_2021.xlsx'
# fs = _NPOI.FileStream(excelPath,_NPOI.FileMode.Open,_NPOI.FileAccess.Read)#,_NPOI.FileShare.ReadWrite
# # ps = _NPOI.np.POIFS.FileSystem.NPOIFSFileSystem(fs)
# book1 = _NPOI.np.XSSF.UserModel.XSSFWorkbook(fs)
# sheet = book1.GetSheetAt(0)
# row = sheet.GetRow(11)
# row.GetCell(7)
# cell = row.CreateCell(7)
# cell.SetCellValue('Test')
# fs = _NPOI.FileStream(excelPath, _NPOI.FileMode.Create, _NPOI.FileAccess.Write)#,_NPOI.FileShare.ReadWrite
# # fout.Flush()
# book1.Write(fs)
# book1.Close()
# fs.Close()
# # book1 = None
# # fout.Close()

## Beispiel Excel Lesen

# MODE: FileMode.OpenOrCreate,FileMode.Open,FileMode.Append,FileMode.CreateNew,FileMode.Truncate,FileMode.Create
# ACCESS: FileAccess.ReadWrite,FileAccess.Read,FileAccess.Write

# fs = FileStream(FilePath,FileMode.OpenOrCreate,FileAccess.ReadWrite)

# bis excel 2007 (.xls)
# workbook = np.HSSF.UserModel.HSSFWorkbook(fs)

# ab excel 2007 (.xls)
# workbook = np.XSSF.UserModel.XSSFWorkbook(fs)

# w.NumberOfSheets

# sheet = w.GetSheetAt(index)
# sheet = w.GetSheet(sheetname)

# Anzahl von Rows: sheet.LastRowNum
# Anzahl von Columns des Rows: sheet.GetRow(r).LastRowNum

# if sheet.GetRow(r):
#   cell = sheet.GetRow(r).GetCell(0)
#   Text = cell.StringCellValue

# for n in range(w.NumberOfSheets):
#     sheet = w.GetSheetAt(n)
#     print(sheet.SheetName,sheet.LastRowNum)
#     rs = sheet.LastRowNum
#     for r in range(rs+1):
#         a += 1
#         if sheet.GetRow(r):text = sheet.GetRow(r).GetCell(0)
