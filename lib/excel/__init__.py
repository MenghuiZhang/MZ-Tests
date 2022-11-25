# coding: utf8
import xlsxwriter
import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal

def get_cell_Daten(sheet,r,c):
    """für NPOI"""
    row = sheet.GetRow(r)
    if row:
        cell = row.GetCell(c)
        if cell:
            if cell.CellType.value__ == 0:
                try:return cell.NumericCellValue
                except:return ''
            elif cell.CellType.value__ == 1:
                try:return cell.StringCellValue
                except:
                    try:return cell.RichStringCellValue
                    except:return ''

            elif cell.CellType.value__ == 2:
                try:return cell.StringCellValue
                except:
                    try:return cell.NumericCellValue
                    except:return ''
            elif cell.CellType.value__ == 3:
                return ''
            elif cell.CellType.value__ == 4:
                try:return cell.BooleanCellValue
                except:return ''

            elif cell.CellType.value__ == 5:
                try:return cell.ErrorCellValue
                except:return ''
            return ''
        else:
            return ''
    else:
        return ''

def get_cell(row,column):
    """für NPOI"""
    cell = row.GetCell(column)
    if cell:return cell
    else:
        cell = row.CreateCell(column)
        return cell


class WExcel():
    def __init__(self,dt,ename = None,sname = None,onesheet = True):
        self.data = dt
        self.ename = System.GUID.NewGuid()+'.xaml' if not ename else ename
        self.sname = sname 
        self.onesheet = onesheet
    
    def wexcel(self):
        b = self.cexcel()
        if self.onesheet:
            if self.sname:
                self.sname = self.sname if len(self.sname) < 32 else self.sname[:31]
                self.wonesheet(self.data,b,self.sname)
            else:self.wonesheet(self.data,b)
        else:
            try:self.sname = list(set(self.sname))
            except:pass
            for n,d in enumerate(self.data):
                if self.sname:
                    try:sname = self.sname[n]
                    except:sname = None
                else:sname = None
                self.wonesheet(d,b,sname)
        b.close()  

    @staticmethod
    def write_one_sheet(d,p,sheetname = None):
        e = xlsxwriter.Workbook(p)
        if sheetname:worksheet = e.add_worksheet(sheetname)
        else:worksheet = e.add_worksheet()
        if len(d) == 0:
            return
        else:
            if len(d[0]) == 0:
                return
        for r in range(len(d)):
            for c in range(len(d[0])):
                wert = d[r][c]
                if wert.GetType().ToString() == 'System.String':
                    worksheet.write(r, c, wert)
                else:
                    try:worksheet.write_number(r, c, wert)
                    except:worksheet.write(r, c, wert)
        worksheet.autofilter(0, 0, int(len(d))-1, int(len(d[0])-1))
        e.close()

    @staticmethod
    def write_sheetswithname(da,p):
        e = xlsxwriter.Workbook(p)
        for sheetname in da.keys():
            d = da[sheetname]
            if sheetname:worksheet = e.add_worksheet(sheetname)
            else:worksheet = e.add_worksheet()
            if len(d) == 0:
                continue
            else:
                if len(d[0]) == 0:
                    continue
            for r in range(len(d)):
                for c in range(len(d[0])):
                    wert = d[r][c]
                    if wert.GetType().ToString() == 'System.String':
                        worksheet.write(r, c, wert)
                    else:
                        try:worksheet.write_number(r, c, wert)
                        except:worksheet.write(r, c, wert)
            worksheet.autofilter(0, 0, int(len(d))-1, int(len(d[0])-1))
        e.close()
    
    @staticmethod
    def write_sheetswithoutname(da,p):
        e = xlsxwriter.Workbook(p)
        for d in da:
            worksheet = e.add_worksheet()
            if len(d) == 0:
                continue
            else:
                if len(d[0]) == 0:
                    continue
            for r in range(len(d)):
                for c in range(len(d[0])):
                    wert = d[r][c]
                    if wert.GetType().ToString() == 'System.String':
                        worksheet.write(r, c, wert)
                    else:
                        try:worksheet.write_number(r, c, wert)
                        except:worksheet.write(r, c, wert)
            worksheet.autofilter(0, 0, int(len(d))-1, int(len(d[0])-1))
        e.close()


    def cexcel(self):
        return xlsxwriter.Workbook(self.ename)
    
    def wonesheet(self,d,e,sheetname = None):
        if sheetname:worksheet = e.add_worksheet(sheetname)
        else:worksheet = e.add_worksheet()
        if len(d) == 0:
            return
        else:
            if len(d[0]) == 0:
                return
        for r in range(len(d)):
            for c in range(len(d[0])):
                wert = d[r][c]
                if wert.GetType().ToString() == 'System.String':
                    worksheet.write(r, c, wert)
                else:
                    try:worksheet.write_number(r, c, wert)
                    except:worksheet.write(r, c, wert)
        worksheet.autofilter(0, 0, int(len(d))-1, int(len(d[0])-1))

class LExcel():
    def __init__(self,p):
        self.path = p
    
    @property
    def existexcel(self):
        return System.IO.File.Exists(self.p)
    
    def readexcel(self):
        if not self.existexcel:
            return {}
        _dict = {}
        ep = ex.ApplicationClass()
        b = ep.Workbooks.Open(self.path)
        for sheet in b.Worksheets:
            sliste = []
            for r in range(1,sheet.UsedRange.Rows.Count+1):
                rliste = [sheet.Cells[r, c].Value2 for c in range(1,sheet.UsedRange.Columns.Count+1)]
                sliste.append(rliste)
            _dict[sheet.Name] = sliste
        b.Close()
        Marshal.FinalReleaseComObject(b)
        ex.Quit()
        Marshal.FinalReleaseComObject(ex)

        return _dict