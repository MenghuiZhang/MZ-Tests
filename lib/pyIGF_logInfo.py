# coding: utf8
import xlsxwriter
import os
from time import strftime, localtime
from getpass import getuser
from rpw import revit
import sqlite3

import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass
from System.Runtime.InteropServices import Marshal

doc = revit.doc

def getlog(Programmname):#,docname):
    exapp = ApplicationClass()
    exapp.Visible = False
    path = r'R:\Vorlagen\_IGF\_pyRevit\Historie.xlsx'
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
    book = exapp.Workbooks.Open(path)
    sheet = book.Worksheets[1]
    rows = sheet.UsedRange.Rows.Count
    try:
        rows += 1
        for col in range(1,len(logdaten)+1):
            sheet.Cells[rows,col] = logdaten[col-1]
        
        book.Save()
        book.Close()
        Marshal.FinalReleaseComObject(sheet)
        Marshal.FinalReleaseComObject(book)
        exapp.Quit()
        Marshal.FinalReleaseComObject(exapp)
    except Exception as e:
        book.Save()
        book.Close()
        Marshal.FinalReleaseComObject(sheet)
        Marshal.FinalReleaseComObject(book)
        exapp.Quit()
        Marshal.FinalReleaseComObject(exapp)
    # path = r'R:\Vorlagen\_IGF\_pyRevit\Historie.db'
    # benutzername = getuser()
    # zeit = strftime("%d.%m.%Y-%H:%M", localtime())
    # name = docname.ProjectInformation.Name
    # number = docname.ProjectInformation.Number
    # ClientName = docname.ProjectInformation.ClientName
    # cloud = str(docname.IsModelInCloud)
    # logdaten = [(zeit, benutzername, name, number, ClientName,cloud, Programmname)]

    # con = sqlite3.connect(path)
    # cur = con.cursor()
    # try:
    #     cur.executemany('INSERT INTO Historie VALUES (?,?,?,?,?,?,?)', logdaten)
    # except:
    #     cur.execute('''CREATE TABLE Historie(
    #         Zeit TEXT,
    #         Benutzer TEXT,
    #         Name TEXT,
    #         Nummer TEXT,
    #         ClientName TEXT,
    #         Cloud_Modell TEXT,
    #         Programm TEXT)''')
    #     cur.executemany('INSERT INTO Historie VALUES (?,?,?,?,?,?,?)', logdaten)
    # con.commit()
    # con.close()
    
