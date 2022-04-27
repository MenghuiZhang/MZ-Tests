# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import System

exapp = Excel.ApplicationClass()
exapp.Visible = False


"""
Funktion:
Excel_Lesen(path = None, sheet_Name = None, Rows = None, Cols = None)
"""

def Excel_Lesen(path = None, sheet_Name = None, Rows = None, Cols = None):
    out_dict = {}
    book = None
    try:
        book = exapp.Workbooks.Open(path)
    except Exception as e:
        print(e)
        return False

    def daten_ermitteln(sheetName = None, rows = None, cols = None):
        out_list = []
        sheet = None
        if sheetName:
            try:
                sheet = book.Worksheets[sheetName]
            except Exception as ex:
                print(ex)
                return False
            if rows:
                for row in rows:
                    temp_list = []
                    if cols:
                        for col in cols:
                            wert = sheet.Range.Cells[row, col].Value2
                            temp_list.append(wert)
                    else:
                        cols = sheet.UsedRange.Columns.Count
                        for col in cols:
                            wert = sheet.Range.Cells[row, col].Value2
                            temp_list.append(wert)
                    out_list.append(temp_list)
                return out_list
            else:
                rows = sheet.UsedRange.Rows.Count
                temp_list = []
                for row in rows:
                    temp_list = []
                    if cols:
                        for col in cols:
                            wert = sheet.Range.Cells[row, col].Value2
                            temp_list.append(wert)
                    else:
                        cols = sheet.UsedRange.Columns.Count
                        for col in cols:
                            wert = sheet.Range.Cells[row, col].Value2
                            temp_list.append(wert)
                out_list.append(temp_list)
            return out_list

    if sheet_Name:
        return daten_ermitteln(sheetName = sheet_Name, rows = Rows, cols = Cols)
    else:
        daten = None
        for s in book.Worksheets:
            sheetname = s.Name
            daten = daten_ermitteln(sheetName = sheetname, rows = Rows, cols = Cols)
            out_dict[sheetname] = daten
        return out_dict
    book.Save()
    book.Close()

def Liste2Dict(Liste):
    outdict = {}
    for ele in Liste:
        outdict[ele[0]] = ele[1:]
    return outdict

def Excelueberschreiben(Path = None, Dict = None, rows = None, cols = None):
    book = exapp.Workbooks.Open(Path)
    for sheet in book.Worksheets:
        sheetName = sheet.Name
        daten = None
        if sheetName in Dict.keys():
            daten = Dict[sheetName]
        else:
            continue

        daten_dict = Liste2Dict(daten)

        for row in rows:
            abgleich = sheet.Cells[row,cols[0]].Value2
            if abgleich in daten_dict.keys():
                n = 0
                for col in cols[1:]:
                    sheet.Cells[row,col] = daten_dict[abgleich][n]
                    n += 1
            else:
                pass

    book.Save()
    book.Close()

def Excelschreiben(Path = None, Dict = None, cols = None):
    book = exapp.Workbooks.Open(Path)
    for sheet in book.Worksheets:
        sheetName = sheet.Name
        rows = sheet.UsedRange.Rows.Count
        daten = None
        if sheetName in Dict.keys():
            daten = Dict[sheetName]
        else:
            continue

        daten_dict = Liste2Dict(daten)
        temp = {}
        for el in rows:
            temp[sheet.Cells[el,cols[0]].Value2] = el

        for da_input in daten_dict.keys():
            if da_input in temp.keys():
                n = 0
                for col in cols[1:]:
                    sheet.Cells[temp[da_input],col] = daten_dict[da_input][n]
                    n += 1
            else:
                n = 0
                rows = rows + 1
                for col in cols[1:]:
                    sheet.Cells[rows+1,col] = daten_dict[da_input][n]
                    n += 1


    sheet_list = []
    for sh in book.Worksheets:
        sheetName = sh.Name
        sheet_list.append(sh)
    for el in Dict.keys():
        if not el in sheet_list:
            sheet = book.Worksheets.Add()
            sheet.Name = el
            daten = Dict[el]
            j = 0
            for n in len(daten):
                for col in cols:
                    sheet.Cells[n+1,col] = daten[n][j]
                    j += 1

    book.Save()
    book.Close()
