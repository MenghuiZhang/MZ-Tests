import xlsxwriter
import os
from time import strftime, localtime
from getpass import getuser
import Autodesk.Revit.DB as DB

import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass
from System.Runtime.InteropServices import Marshal

doc = __revit__.ActiveUIDocument.Document

def getlog(Programmname):
    exapp = ApplicationClass()
    exapp.Visible = False
    path = r'R:\pyRevit\xx_Skripte\Historie.xlsx'
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


def getloglocal(Programmname):
    benutzername = getuser()
    zeit = strftime("%d.%m.%Y-%H:%M", localtime())
    name = doc.ProjectInformation.Name
    number = doc.ProjectInformation.Number
    ClientName = doc.ProjectInformation.ClientName
    cloud = doc.IsModelInCloud
    if cloud:
        cloud = 'CloudProjekt'
    else:
        cloud = 'NichtCloudProjekt'
    logdaten = str(zeit) +': '+ benutzername +', '+ name +', ' + number +', ' + ClientName + ', ' + cloud +', '+ Programmname
    Desktop = os.path.expanduser("~\AppData\Roaming\pyRevit")
    bericht = open(Desktop+"\\logdaten.txt","a")
    bericht.write("\n[{}]".format(logdaten))
    bericht.close()