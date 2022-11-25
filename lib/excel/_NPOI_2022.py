import sys
import clr
import os
# sys.path.Add(os.path.dirname(__file__))
sys.path.Add(os.path.dirname(__file__)+'\\net45')
clr.AddReference('NPOI')
clr.AddReference('NPOI.OOXML')
import NPOI as np
from System.IO import FileStream,FileMode,FileAccess,FileShare

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
