# coding: utf8
import xlsxwriter
import os
from time import strftime, localtime
from getpass import getuser
import Autodesk.Revit.DB as DB

doc = __revit__.ActiveUIDocument.Document
revitversion = __revit__.Application.VersionNumber

if revitversion == '2020':
    import excel._NPOI_2020 as _NPOI

else:
    import excel._NPOI_2022 as _NPOI

def get_cell(sheet,row,column):
    Row = sheet.CreateRow(row)
    cell = Row.CreateCell(column)
    return cell

def getlog(Programmname):
    path = r'R:\pyRevit\xx_SkripteLogs\Historie.xlsx'
    if not os.path.isfile(path):
        try:
            workbook = xlsxwriter.Workbook(path)
            worksheet = workbook.add_worksheet()
            ueberschrift = ['Zeit', 'Benutzername', 'Name', 'Number', 'ClientName', 'Cloud-Modell',  "Programm"]
            for col in range(len(ueberschrift)):
                worksheet.write(0, col, ueberschrift[col])
            workbook.close()
        except:
            try:getloglocal(Programmname)
            except:pass
    
    fs = _NPOI.FileStream(path,_NPOI.FileMode.Open,_NPOI.FileAccess.Read)
    book1 = _NPOI.np.XSSF.UserModel.XSSFWorkbook(fs)
    sheet = book1.GetSheetAt(0)

    benutzername = getuser()
    zeit = strftime("%d.%m.%Y-%H:%M", localtime())
    name = doc.ProjectInformation.Name
    number = doc.ProjectInformation.Number
    ClientName = doc.ProjectInformation.ClientName
    cloud = str(doc.IsModelInCloud)
    logdaten = [zeit, benutzername, name, number, ClientName,cloud, Programmname]
    rows = sheet.LastRowNum
    row = sheet.CreateRow(rows+1)
    for col in range(len(logdaten)):
        cell = row.CreateCell(col)
        cell.SetCellValue(logdaten[col])
    fs = _NPOI.FileStream(path, _NPOI.FileMode.Create, _NPOI.FileAccess.Write)
    book1.Write(fs)
    book1.Close()
    fs.Close()


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