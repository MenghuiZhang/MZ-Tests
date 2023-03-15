# coding: utf8
from Autodesk.Revit.UI import TaskDialog
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
import os
from Luftbilanz_Export import VSRHerstellerDaten

Liste_Datenbank = {}
Liste_Datenbank1 = {}
DICT_DatenBank = {}
datenbankadresse = r'R:\pyRevit\03_Lüftung\02_Luftungsbilanz\IGF_RLT_Volumenstromregler_Datenbank.xlsx'
Liste_Fabrikat = []

def get_Herstellerdaten(datenbankadresse):
    Liste_Fabrikat = []
    Liste_Datenbank = {}
    Liste_Datenbank1 = {}
    DICT_DatenBank = {}
    if os.path.exists(datenbankadresse):
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        fs = FileStream(datenbankadresse,FileMode.Open,FileAccess.Read)
        book = ExcelPackage(fs)
        
        for sheet in book.Workbook.Worksheets:
            try:
                fabrikat = sheet.Name
                Liste_Datenbank[fabrikat] = {}

                Liste_Datenbank1[fabrikat] = {}
                Liste_Fabrikat.append(fabrikat)
                maxRowNum = sheet.Dimension.End.Row
                for row in range(3,maxRowNum+1):
                    luftart = 0 # 0:Zu/Abluft,1:Zuluft,2:Abluft,3:PPS
                    vsrart = 'VVR'
                    art = sheet.Cells[row, 1].Value
                    A = sheet.Cells[row, 2].Value
                    B = sheet.Cells[row, 3].Value
                    D = sheet.Cells[row, 4].Value
                    Bezeichnung = sheet.Cells[row, 5].Value
                    Typ = sheet.Cells[row, 7].Value
                    Spannung = sheet.Cells[row, 8].Value
                    Nennstrom = sheet.Cells[row, 9].Value
                    Leistung = sheet.Cells[row, 10].Value
                    vmin = sheet.Cells[row, 11].Value
                    if not vmin:vmin = 0
                    vmax = sheet.Cells[row, 12].Value
                    if not vmax:vmax = 0
                    vnenn = sheet.Cells[row, 13].Value
                    if not vnenn:vnenn = 0
                    Bemerkung = sheet.Cells[row, 16].Value
                    
                    if any([A,B,D]):
                        if Bemerkung and Bemerkung.upper().find('PPS') != -1:
                            luftart = 3
                        else:
                            if art:
                                if art.upper().find('ZU') != -1:
                                    luftart = 1
                                elif art.upper().find('AB') != -1:
                                    luftart = 2
                        if Bemerkung:
                            if Bemerkung.upper().find('VARI') == -1:
                                vsrart = 'KVR'
                        if A and B:
                            mini = min([int(A),int(B)])
                            maxi = max([int(A),int(B)])
                            daten = VSRHerstellerDaten(fabrikat,Typ,mini,maxi,None,Spannung,Nennstrom,Leistung,vmin,vmax,vnenn,Bemerkung)
                            if luftart not in Liste_Datenbank1[fabrikat].keys():
                                Liste_Datenbank1[fabrikat][luftart] = {}
                            if vsrart not in Liste_Datenbank1[fabrikat][luftart].keys():
                                Liste_Datenbank1[fabrikat][luftart][vsrart] = {}
                            if maxi not in Liste_Datenbank1[fabrikat][luftart][vsrart].keys():
                                Liste_Datenbank1[fabrikat][luftart][vsrart][maxi] = {}
                            if mini not in Liste_Datenbank1[fabrikat][luftart][vsrart][maxi].keys():
                                Liste_Datenbank1[fabrikat][luftart][vsrart][maxi][mini] = []
                            Liste_Datenbank1[fabrikat][luftart][vsrart][maxi][mini].append(daten)
                            DICT_DatenBank[daten.typ] = daten

                        else:
                            daten = VSRHerstellerDaten(fabrikat,Typ,None,None,int(D),Spannung,Nennstrom,Leistung,vmin,vmax,vnenn,Bemerkung)
                            if luftart not in Liste_Datenbank[fabrikat].keys():
                                Liste_Datenbank[fabrikat][luftart] = {}
                            if vsrart not in Liste_Datenbank[fabrikat][luftart].keys():
                                Liste_Datenbank[fabrikat][luftart][vsrart] = {}
                            if int(D) not in Liste_Datenbank[fabrikat][luftart][vsrart].keys():
                                Liste_Datenbank[fabrikat][luftart][vsrart][int(D)] = []
                            Liste_Datenbank[fabrikat][luftart][vsrart][int(D)].append(daten)
                            DICT_DatenBank[daten.typ] = daten

            except Exception as e:
                print(e)
                TaskDialog.Show('Fehler','Fehler beim Lesen der Excel-Datei')        
                break   
            
   
        book.Dispose()
        fs.Dispose()
        del fs
        del book
    return Liste_Fabrikat,DICT_DatenBank,Liste_Datenbank,Liste_Datenbank1

Liste_Fabrikat,DICT_DatenBank,Liste_Datenbank,Liste_Datenbank1 = get_Herstellerdaten(datenbankadresse)