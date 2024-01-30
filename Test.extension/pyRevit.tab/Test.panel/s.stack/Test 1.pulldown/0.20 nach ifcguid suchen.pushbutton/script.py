# coding: utf8
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
from IGF_Funktionen._Parameter import wert_schreibenbase
from System.IO import FileInfo
from rpw import revit,DB

doc = revit.doc

uidoc = revit.uidoc

def Funktion_Regel_Info_Schreiben():
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    path = r"G:\PROJEKTE_us\2020-003-MFC-BBIM\Dokumente\_IGF\Technische_Planung\BERECHNUNGEN\Lueftung\BER 48 - 53 - Raumbilanzen UG - 4. OG MFC\Raumluftbilanz_MFC_UG.xlsx"
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileInfo(path)
    book = ExcelPackage(fs)
    sheet = book.Workbook.Worksheets['Sheet1']
    Dict1 = {el.LookupParameter("SBI_Bauteilnummerierung").AsString():el for el in cl}
    maxRowNum = sheet.Dimension.End.Row
    t = DB.Transaction(doc,"1")
    t.Start()
    for row in range(1,maxRowNum+1):
        nummer = sheet.Cells[row, 1].Value
        if not nummer:continue
        if nummer.find("KVR_") != -1 or nummer.find("VVR_") != -1:
            if nummer in Dict1.keys():
                Vmin =  sheet.Cells[row, 16].Value
                Vmax =  sheet.Cells[row, 17].Value
                typ = sheet.Cells[row, 6].Value
                try:
                    Dict1[nummer].LookupParameter("IGF_RLT_VSR_Regelbereich_Min").SetValueString(str(Vmin))
                except:pass
                try:
                    Dict1[nummer].LookupParameter("IGF_RLT_VSR_Regelbereich_Max (bei 6m³/h)").SetValueString(str(Vmax))
                except:pass
                try:
                    Dict1[nummer].LookupParameter("IGF_RLT_VSR_Herstellertyp").Set(typ)
                except:pass
    t.Commit()
    book.Save()
    book.Dispose()

Funktion_Regel_Info_Schreiben()



# excel = r'C:\Users\zhang\Desktop\LAKAS.xlsx'

# from System.Collections.Generic import List 

# # multi = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_DuctTerminal,DB.BuiltInCategory.OST_DuctAccessory]))
# # cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WherePasses(multi).WhereElementIsNotElementType().ToElementIds()

# mep = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
# _dict = {el.Number:el.LookupParameter("Name").AsString() for el in mep}
# BSK0 = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()
# BSK1 = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
# Liste = []
# for el in BSK0:
#     if el.Symbol.FamilyName.find("Brandschutzklappe") != -1:
#         Liste.append(el)

# for el in BSK1:
#     if el.Symbol.FamilyName.find("Brandschutz") != -1 or el.Name.find("KU-K30") != -1:
#         Liste.append(el)    

# class BSK:
#     def __init__(self,elem):
#         self.elem = elem
#         self.Id = int(elem.Id.ToString())
#         self.nummer1 = self.elem.LookupParameter("IGF_RLT_BSK-Nummer").AsString() 
#         self.Size = self.elem.LookupParameter("Größe").AsString() 
#         self.Bedienseite = self.elem.LookupParameter("IGF_X_Einbauort").AsString() 
#         self.Wirkunsort = self.elem.LookupParameter("IGF_X_Wirkungsort").AsString() 
#         self.elemid = self.elem.Id.ToString()

#         self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
#         self.wand = self.elem.LookupParameter('IGF_HLS_Brandschutz').AsString()
#         self.Location = self.elem.Location.Point
#         self.familytyp = self.elem.Symbol.FamilyName + ': '+self.elem.Name
#         self.Ebene = self.elem.LookupParameter("Ebene").AsValueString()


# path = r"C:\Users\zhang\Desktop\Brandschutzklappe_MFC.xlsx"
# # path1 = r"C:\Users\zhang\Desktop\MFC BSK UG _01.11.2024.xlsx"

# 

# fs1 = FileInfo(path1)
# book1 = ExcelPackage(fs1)

# sheet = book.Workbook.Worksheets['BSK']

# for row in range(2, 721):
#     nummer = sheet.Cells[row, 4].Value
#     if nummer:
#     # if nummer:
#         Text = ""
#         if sheet.Cells[row,10].Value == "Passt Nicht":
#             Text += "Größe geändert, "
        
#         if sheet.Cells[row,20].Value == "Passt Nicht":
#             Text += "BSK Klasse geändert, "
#         try:
#             if sheet.Cells[row,23].Value > 30:
#                 Text += "Verschiebung: " + str(sheet.Cells[row,23].Value) +  " cm"
#         except:
#             pass
#         if not Text:
#             Text = "Passt"
#         sheet.Cells[row,26].Value = Text
#     else:
#         sheet.Cells[row,26].Value = "entfallen"
# for row in range(721, 772):sheet.Cells[row,26].Value = "neu"


# 2.UG: -15.944881889763785
# 1.UG: -2.8871391076115436
# EG: 12.434383202099736
# 1.OG: 26.049868766404199
# 2.OG: 39.665354330708659
# 3.OG: 53.280839895013109

# Dict = {}
# Dict["EG"] = DB.ElementId(17024597)
# Dict["1.OG"] = DB.ElementId(17024600)
# Dict["2.OG"] = DB.ElementId(17024603)
# Dict["3.OG"] = DB.ElementId(17024605)
# Dict["4.OG"] = DB.ElementId(17024627)
# Dict["1.UG"] = DB.ElementId(21885480)
# Dict["2.UG"] = DB.ElementId(17024632)

# t = DB.Transaction(doc,"1")
# t.Start()
# for row in range(2, 772):
#     nummer = sheet.Cells[row,4].Value
# #     if nummer:
#         elem = doc.GetElement(DB.ElementId(int(nummer)))
#         _id = sheet.Cells[row,6].Value
#         elem.LookupParameter("IGF_RLT_BSK-Nummer").Set(_id)
    # if ebene == "Passt nicht":
    #     nummer = sheet.Cells[row,4].Value
    #     elem = doc.GetElement(DB.ElementId(int(nummer)))
    #     ebene1 = sheet.Cells[row,27].Value
    #     try:
    #         elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM).Set(Dict[ebene1])
    #         elem.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(elem.Location.Point.Z-doc.GetElement(Dict[ebene1]).LookupParameter("Ansicht").AsDouble())
    #     except Exception as e:
    #         print(int(nummer),(elem.Location.Point.Z-doc.GetElement(Dict[ebene1]).LookupParameter("Ansicht").AsDouble())*304.8)
# t.Commit()
    # if nummer:
    #     ebene = sheet.Cells[row,27].Value
    #     ebene1 = sheet.Cells[row,6].Value
    #     if ebene == "1.UG" and ebene1.find("_090_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"
    #     if ebene == "2.UG" and ebene1.find("_080_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"
    #     if ebene == "EG" and ebene1.find("_100_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"
    #     if ebene == "1.OG" and ebene1.find("_110_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"
    #     if ebene == "2.OG" and ebene1.find("_120_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"
    #     if ebene == "3.OG" and ebene1.find("_130_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"
    #     if ebene == "4.OG" and ebene1.find("_140_") == -1:
    #         sheet.Cells[row,30].Value = "Passt nicht"


        # if ebene != ebene1:
        #     sheet.Cells[row,30].Value = "Passt nicht"
        # else:
        #     sheet.Cells[row,30].Value = "Passt"
        

        # elem = doc.GetElement(DB.ElementId(int(nummer)))
        # height = elem.Location.Point.Z
        # ebene = ""
        # if height <= -15.944881889763785:
        #     ebene = "2.UG"
        # elif height <= -1.18110236220472:
        #     ebene = "1.UG"
        # elif height <= 12.434383202099736:
        #     ebene = "EG"
        # elif height <= 26.049868766404199:
        #     ebene = "1.OG"
        # elif height <= 39.665354330708659:
        #     ebene = "2.OG"
        # elif height <= 53.280839895013109:
        #     ebene = "3.OG"
        # else:
        #     ebene = "4.OG"
        # sheet.Cells[row,27].Value = ebene

        # try:
        #     if sheet.Cells[row,6].Value.strip() == "MFC_" + sheet.Cells[row,5].Value.strip():
        #         sheet.Cells[row,7].Value = "Passt"
        #     else:
        #         sheet.Cells[row,7].Value = "Passt Nicht"
        # except:
        #     sheet.Cells[row,7].Value = "Passt Nicht"

        # if sheet.Cells[row,8].Value == sheet.Cells[row,9].Value:
        #     sheet.Cells[row,10].Value = "Passt"
        # else:
        #     sheet.Cells[row,10].Value = "Passt Nicht"

        # if sheet.Cells[row,15].Value == sheet.Cells[row,16].Value:
        #     sheet.Cells[row,17].Value = "Passt"
        # else:
        #     sheet.Cells[row,17].Value = "Passt Nicht"

        # if sheet.Cells[row,18].Value == sheet.Cells[row,19].Value:
        #     sheet.Cells[row,20].Value = "Passt"
        # else:
        #     sheet.Cells[row,20].Value = "Passt Nicht"


        # if sheet.Cells[row,21].Value == sheet.Cells[row,22].Value:
        #     sheet.Cells[row,23].Value = 0
        # else:
        #     loc = list(eval(sheet.Cells[row,21].Value))
        #     loc1 = list(eval(sheet.Cells[row,22].Value))
        #     xyz = DB.XYZ(loc[0],loc[1],loc[2])
        #     xyz1 = DB.XYZ(loc1[0],loc1[1],loc1[2])
        #     distance = int(xyz.DistanceTo(xyz1) * 304.8)/10.0
        #     sheet.Cells[row,23].Value = distance
            




# sheet1 = book1.Workbook.Worksheets['Sheet1']



# DIct = {}
# for n in range(2,200):
#     nummer0 = sheet1.Cells[n,2].Value
#     nummer1 = sheet1.Cells[n,3].Value
#     if nummer0:
#         DIct[nummer0] = nummer1



# for n in range(2,721):
#     nummer0 = sheet.Cells[n,3].Value
#     nummer1 = sheet.Cells[n,4].Value
#     _id = sheet.Cells[n,6].Value

#     if nummer0 in DIct.keys():
#         if nummer1 != DIct[nummer0]:
#             print(_id,nummer1,DIct[nummer0])

# sheet1 = book.Workbook.Worksheets['BSK_12.01.2024']
# liste = []
# for n in range(2,721):
#     nummer0 = sheet.Cells[n,4].Value
#     liste.append(nummer0)
# for n in  range(721,817):
#     nummer0 = sheet.Cells[n,4].Value
#     if nummer0 in liste:
#         sheet.Cells[n,1].Value = "Zur Prüfen"
#     else:
#         sheet.Cells[n,1].Value = ""
#     if nummer0:
#         for n1 in range(2,731):
#             nummer1 = sheet1.Cells[n1,2].Value
#             if nummer0 == nummer1:
#                 sheet.Cells[n,2].Value = sheet1.Cells[n1,1].Value
#                 sheet.Cells[n,4].Value = sheet1.Cells[n1,2].Value
#                 sheet.Cells[n,7].Value = sheet1.Cells[n1,3].Value
#                 sheet.Cells[n,10].Value = sheet1.Cells[n1,4].Value
#                 sheet.Cells[n,13].Value = sheet1.Cells[n1,5].Value
#                 sheet.Cells[n,16].Value = sheet1.Cells[n1,6].Value
#                 sheet.Cells[n,19].Value = sheet1.Cells[n1,7].Value
#                 sheet.Cells[n,22].Value = sheet1.Cells[n1,8].Value
#                 sheet.Cells[n,25].Value = sheet1.Cells[n1,9].Value
#                 sheet.Cells[n,28].Value = sheet1.Cells[n1,10].Value
#                 sheet.Cells[n,1].Value = "Zur Prüfen"

# Dict = {}
# Dict[726] = 322
# Dict[764] = 169
# Dict[765] = 168
# Dict[769] = 380
# Dict[771] = 575
# # Dict[789] = 604
# Dict[807] = 167
# Dict[792] = 662
# Dict[791] = 661
# # Dict[790] = 602
# for n in Dict.keys():
#     n1 = Dict[n]
   

#         # for n1 in range(2,731):
#         #     nummer1 = sheet1.Cells[n1,2].Value
#         #     if nummer0 == nummer1:
#     sheet.Cells[n1,2].Value = sheet.Cells[n,2].Value
#     sheet.Cells[n1,4].Value = sheet.Cells[n,4].Value
#     sheet.Cells[n1,7].Value = sheet.Cells[n,7].Value
#     sheet.Cells[n1,10].Value = sheet.Cells[n,10].Value
#     sheet.Cells[n1,13].Value = sheet.Cells[n,13].Value
#     sheet.Cells[n1,16].Value = sheet.Cells[n,16].Value
#     sheet.Cells[n1,19].Value = sheet.Cells[n,19].Value
#     sheet.Cells[n1,22].Value = sheet.Cells[n,22].Value
#     sheet.Cells[n1,25].Value = sheet.Cells[n,25].Value
#     sheet.Cells[n1,28].Value = sheet.Cells[n,28].Value
#     sheet.Cells[n,1].Value = "Zur Prüfen"

# for n in range(721,817):
#     location = sheet.Cells[n,25].Value
#     location_list = list(eval(location))
#     xyz = DB.XYZ(location_list[0],location_list[1],location_list[2])
#     einbau = sheet.Cells[n,13].Value
#     wiru = sheet.Cells[n,16].Value
#     systemtyp = sheet.Cells[n,19].Value

#     for n1 in range(2,721):
#         location1 = sheet.Cells[n1,24].Value
#         location_list1 = list(eval(location1))

#         xyz1 = DB.XYZ(location_list1[0],location_list1[1],location_list1[2])
#         if xyz.DistanceTo(xyz1) < 30/304.8:
#             einbau1 = sheet.Cells[n1,12].Value
#             wiru1 = sheet.Cells[n1,15].Value
#             systemtyp1 = sheet.Cells[n1,18].Value
#             print(30*"-")
#             print(n,n1)
#             print(einbau,wiru,systemtyp)
#             print(einbau1,wiru1,systemtyp1)
            

    # nummer0 = sheet.Cells[n,4].Value
    # if nummer0 in liste:
    #     sheet.Cells[n,1].Value = "Zur Prüfen"
    # else:
    #     sheet.Cells[n,1].Value = ""


# Dict = {}
# # liste = []
# for n in range(2,721):
#     liste.append(sheet.Cells[n,3].Value)
#     # Dict[sheet1.Cells[n,2].Value] = n
# n1 = 721



# for n in range(721,810):
#     _id = sheet.Cells[n,4].Value
#     # einbau = sheet.Cells[n,13].Value
#     # wiru = sheet.Cells[n,16].Value
#     # systemtyp = sheet.Cells[n,19].Value
#     for n1 in range(2,721):
#         _id1 = sheet.Cells[n1,4].Value
#         if _id ==_id1:
#         # nummer = sheet.Cells[n1,4].Value
#         # if not nummer:
#         #     einbau1 = sheet.Cells[n1,12].Value
#         #     wiru1 = sheet.Cells[n1,15].Value
#         #     systemtyp1 = sheet.Cells[n1,18].Value
#         #     if (einbau == einbau1 and wiru == wiru1 and systemtyp == systemtyp1) or (einbau == wiru1 and wiru == einbau1 and systemtyp == systemtyp1):
#                 sheet.Cells[n1,2].Value = sheet.Cells[n,2].Value
#                 sheet.Cells[n1,4].Value = sheet.Cells[n,4].Value
#                 sheet.Cells[n1,7].Value = sheet.Cells[n,7].Value
#                 sheet.Cells[n1,10].Value = sheet.Cells[n,10].Value
#                 sheet.Cells[n1,13].Value = sheet.Cells[n,13].Value
#                 sheet.Cells[n1,16].Value = sheet.Cells[n,16].Value
#                 sheet.Cells[n1,19].Value = sheet.Cells[n,19].Value
#                 sheet.Cells[n1,22].Value = sheet.Cells[n,22].Value
#                 sheet.Cells[n1,25].Value = sheet.Cells[n,25].Value
#                 sheet.Cells[n1,28].Value = sheet.Cells[n,28].Value
#                 sheet.Cells[n,1].Value = "Zur Prüfen"
#                 break

#         n1 += 1
# # n = 2

# for el in Liste:
    # bsk = BSK(el)
    # sheet.Cells[n,1].Value = bsk.Ebene
    # sheet.Cells[n,3].Value = bsk.Id
    # sheet.Cells[n,6].Value = bsk.nummer1
    # sheet.Cells[n,9].Value = bsk.Size
    # try:sheet.Cells[n,12].Value = bsk.Bedienseite + " [" +_dict[bsk.Bedienseite] + "]"
    # except:pass
    # try:sheet.Cells[n,15].Value = bsk.Wirkunsort + " [" +_dict[bsk.Wirkunsort] + "]"
    # except:pass
    # sheet.Cells[n,18].Value = bsk.System
    # sheet.Cells[n,21].Value = bsk.wand
    # sheet.Cells[n,24].Value = bsk.Location
    # sheet.Cells[n,27].Value = bsk.familytyp
    

    # sheet1.Cells[n,1].Value = bsk.Ebene
    # sheet1.Cells[n,2].Value = bsk.Id
    # sheet1.Cells[n,3].Value = bsk.nummer1
    # sheet1.Cells[n,4].Value = bsk.Size
    # try:sheet1.Cells[n,5].Value = bsk.Bedienseite + " [" +_dict[bsk.Bedienseite] + "]"
    # except:pass
    # try:sheet1.Cells[n,6].Value = bsk.Wirkunsort + " [" +_dict[bsk.Wirkunsort] + "]"
    # except:pass
    # sheet1.Cells[n,7].Value = bsk.System
    # sheet1.Cells[n,8].Value = bsk.wand
    # sheet1.Cells[n,9].Value = bsk.Location
    # sheet1.Cells[n,10].Value = bsk.familytyp

    # n += 1
    
# book.Save()
# book.Dispose()


# # Parameter_Dict = {}
# # 
# # # fs = FileStream(excel,FileMode.Open,FileAccess.Read)

# # Parameter_Liste = []
# # Bauteilnummer = ''
# # try:
# sheet = book.Workbook.Worksheets['Sheet1']
# # for sheet in book.Workbook.Worksheets:
# maxColumnNum = sheet.Dimension.End.Column
# maxRowNum = sheet.Dimension.End.Row


# # for row in range(2, 29):
# #     nummer = sheet.Cells[row, 2].Value

# #     if nummer:
# #         for row1 in range(29, 220):
# #             nummer1 = sheet.Cells[row1, 2].Value 
# #             if nummer == nummer1:
# #                 sheet.Cells[row1,2].Value = sheet.Cells[row, 2].Value
# #                 sheet.Cells[row1,5].Value = sheet.Cells[row, 5].Value
# #                 sheet.Cells[row1,8].Value = sheet.Cells[row, 8].Value
# #                 sheet.Cells[row1,11].Value = sheet.Cells[row, 11].Value
# #                 sheet.Cells[row1,14].Value = sheet.Cells[row, 14].Value
# #                 sheet.Cells[row1,17].Value = sheet.Cells[row, 17].Value
# #                 sheet.Cells[row1,20].Value = sheet.Cells[row, 20].Value
# #                 sheet.Cells[row1,23].Value = sheet.Cells[row, 23].Value
# #                 sheet.Cells[row1,25].Value = sheet.Cells[row, 25].Value
# #                 sheet.Cells[row1,24].Value = "Wenig als 100 cm"
# #                 sheet.Cells[row, 2].Value = ""
# #                 sheet.Cells[row, 3].Value = ""
# #                 sheet.Cells[row, 5].Value = ""
# #                 sheet.Cells[row, 8].Value = ""
# #                 sheet.Cells[row, 11].Value = ""
# #                 sheet.Cells[row, 14].Value = ""
# #                 sheet.Cells[row, 17].Value = ""
# #                 sheet.Cells[row, 20].Value = ""
# #                 sheet.Cells[row, 23].Value = ""
# #                 sheet.Cells[row, 25].Value = ""

# #             Text = ""
# #             if sheet.Cells[row,12].Value == "Passt Nicht":
# #                 Text += "Größe geändert, "
            
# #             if sheet.Cells[row,15].Value == "Passt Nicht" or sheet.Cells[row,18].Value == "Passt Nicht":
# #                 if sheet.Cells[row,14].Value == sheet.Cells[row,16].Value and sheet.Cells[row,17].Value == sheet.Cells[row,13].Value:
# #                     pass
# #                 else:
# #                     Text += "Einbauort/Wirkungsort geändert, "

# #             if sheet.Cells[row,21].Value == "Passt Nicht":
# #                 Text += "Systemtyp geändert, "
# #             if sheet.Cells[row,24].Value == "Passt Nicht":
# #                 Text += "BSK Klasse geändert, "
            
# #             if sheet.Cells[row,27].Value != "Passt":
# #                 Text += "Verschiebung: " + sheet.Cells[row,27].Value
            
# #             sheet.Cells[row,29].Value = Text
# t = DB.Transaction(doc,'Nummer')
# t.Start()
# _090_Liste = []
# _080_Liste = []
# n = 229
# for row in range(2, 215):
#     nummer = sheet.Cells[row, 3].Value
    
#     if nummer:
#         # if cl.Contains(DB.ElementId(int(nummer))):
#             elem = doc.GetElement(DB.ElementId(int(nummer)))
#             # ebene = elem.LookupParameter("Ebene").AsValueString()
#             # sheet.Cells[row, 1].Value = ebene
#             id_ =  sheet.Cells[row, 6].Value
#             # if id_.find("090_") != -1:
#             #     liste1 = id_.split("_")
#             #     if liste1[-1] not in _090_Liste:
#             #         _090_Liste.append(liste1[-1])
#             #     else:print(id_,"090")
#             # elif id_.find("080_") != -1:
#             #     liste1 = id_.split("_")
#             #     if liste1[-1] not in _080_Liste:
#             #         _080_Liste.append(liste1[-1])
#             #     else:print(id_,"080")

#             elem.LookupParameter("IGF_RLT_BSK-Nummer").Set(id_)
#             elem.LookupParameter("SBI_Bauteilnummerierung").Set(id_)
#             # systemtyp = elem.LookupParameter('Systemtyp').AsValueString()
#             # if systemtyp.find('Zuluft') != -1:
#             #     if systemtyp.find('Labor') != -1:
#             #         anlage = 'A10'
#             #     elif systemtyp.find('GMP') != -1:
#             #         anlage = 'A30'
#             #     else:
#             #         anlage = 'A00'
#             #         # print(elem.Id)
#             # elif systemtyp.find('Abluft') != -1:
#             #     if systemtyp.find('Labor') != -1:
#             #         anlage = 'A20'
#             #     elif systemtyp.find('GMP') != -1:
#             #         anlage = 'A40'
#             #     elif systemtyp.find('24h') != -1:
#             #         anlage = 'A51'
#             #     elif systemtyp.find('nachbehandelt') != -1:
#             #         anlage = 'A61'

#             #     else:
#             #         anlage = 'A00'
#             #         # print(elem.Id)
#             # else:
#             #     anlage = 'A00'
#                 # print(elem.Id)
#             # try:
#             #     anlage = doc.GetElement(elem.LookupParameter('Systemtyp').AsElementId()).LookupParameter("IGF_X_AnlagenNr").AsValueString()
#             # except:
#             #     anlage = "xx"
#             # if id_:
                
#             #     if id_.find("090") != -1 and id_.find(anlage) != -1:

#             #         elem.LookupParameter("IGF_RLT_BSK-Nummer").Set(id_)
#             #         elem.LookupParameter("SBI_Bauteilnummerierung").Set(id_)
#             #     else:
#             #         print(elem.Id.ToString(),id_,anlage,systemtyp)
#             # else:
#             #     id_ = "MFC_090_"+anlage+"_"+str(n)
#             #     elem.LookupParameter("IGF_RLT_BSK-Nummer").Set(id_)
#             #     elem.LookupParameter("SBI_Bauteilnummerierung").Set(id_)
#             #     sheet.Cells[row, 6].Value = id_
#             #     n += 1

#         # else:
#         #     print(int(nummer))
    
                
# t.Commit()
# for row in range(2, 215):
#     nummer = sheet.Cells[row, 2].Value
#     nummer3 = sheet.Cells[row, 1].Value
#     if nummer and nummer3:
#     # if nummer:
#         Text = ""
#         if sheet.Cells[row,12].Value == "Passt Nicht":
#             Text += "Größe geändert, "
        
#         if sheet.Cells[row,15].Value == "Passt Nicht" or sheet.Cells[row,18].Value == "Passt Nicht":
#             if sheet.Cells[row,14].Value == sheet.Cells[row,16].Value and sheet.Cells[row,17].Value == sheet.Cells[row,13].Value:
#                 pass
#             else:
#                 Text += "Einbauort/Wirkungsort geändert, "

#         if sheet.Cells[row,21].Value == "Passt Nicht":
#             Text += "Systemtyp geändert, "
#         if sheet.Cells[row,24].Value == "Passt Nicht":
#             Text += "BSK Klasse geändert, "
        
#         if sheet.Cells[row,27].Value != "Passt":
#             Text += "Verschiebung: " + sheet.Cells[row,27].Value
        
#         sheet.Cells[row,29].Value = Text


#         try:
#             if sheet.Cells[row,8].Value == "MFC_" + sheet.Cells[row,7].Value:
#                 sheet.Cells[row,9].Value = "Passt"
#             else:
#                 sheet.Cells[row,9].Value = "Passt Nicht"
#         except:
#             sheet.Cells[row,9].Value = "Passt Nicht"

#         if sheet.Cells[row,11].Value == sheet.Cells[row,10].Value:
#             sheet.Cells[row,12].Value = "Passt"
#         else:
#             sheet.Cells[row,12].Value = "Passt Nicht"

#         if sheet.Cells[row,14].Value == sheet.Cells[row,13].Value:
#             sheet.Cells[row,15].Value = "Passt"
#         else:
#             sheet.Cells[row,15].Value = "Passt Nicht"

#         if sheet.Cells[row,17].Value == sheet.Cells[row,16].Value:
#             sheet.Cells[row,18].Value = "Passt"
#         else:
#             sheet.Cells[row,18].Value = "Passt Nicht"

#         if sheet.Cells[row,20].Value == sheet.Cells[row,19].Value:
#             sheet.Cells[row,21].Value = "Passt"
#         else:
#             sheet.Cells[row,21].Value = "Passt Nicht"

#         if sheet.Cells[row,23].Value == sheet.Cells[row,22].Value:
#             sheet.Cells[row,24].Value = "Passt"
#         else:
#             sheet.Cells[row,24].Value = "Passt Nicht"

#         if sheet.Cells[row,26].Value == sheet.Cells[row,25].Value:
#             sheet.Cells[row,27].Value = "Passt"
#         else:
#             loc = list(eval(sheet.Cells[row,25].Value))
#             loc1 = list(eval(sheet.Cells[row,26].Value))
#             xyz = DB.XYZ(loc[0],loc[1],loc[2])
#             xyz1 = DB.XYZ(loc1[0],loc1[1],loc1[2])
#             if xyz.DistanceTo(xyz1) < 100/304.8:
#                 sheet.Cells[row,27].Value  = "wenig als 10 cm"
#             elif xyz.DistanceTo(xyz1) < 200/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 20 cm"
#             elif xyz.DistanceTo(xyz1) < 300/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 30 cm"
#             elif xyz.DistanceTo(xyz1) < 400/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 40 cm"
#             elif xyz.DistanceTo(xyz1) < 500/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 50 cm"
#             elif xyz.DistanceTo(xyz1) < 600/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 60 cm"
#             elif xyz.DistanceTo(xyz1) < 700/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 70 cm"
#             elif xyz.DistanceTo(xyz1) < 800/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 80 cm"
#             elif xyz.DistanceTo(xyz1) < 900/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 90 cm"
#             elif xyz.DistanceTo(xyz1) < 1000/304.8:
#                 sheet.Cells[row,27].Value = "wenig als 100 cm"
#             else:
#                 sheet.Cells[row,27].Value = "mehr als 100 cm"

    
#     # if nummer:
#     #     loc1 = list(eval(sheet.Cells[row,23].Value))
#     #     for row1 in range(181,460):
#     #         nummer1 = sheet.Cells[row1, 1].Value
#     #         nummer2 = sheet.Cells[row1, 2].Value
#     #         if not nummer2:
#     #             loc = list(eval(sheet.Cells[row1,22].Value))
                
#     #             #print(loc,loc1)
#     #             #print(sheet.Cells[row1,7].Value,sheet.Cells[row, 8].Value,sheet.Cells[row1,7].Value == sheet.Cells[row, 8].Value)
#     #             if  sheet.Cells[row1,10].Value == sheet.Cells[row, 14].Value and \
#     #                 sheet.Cells[row1,13].Value == sheet.Cells[row, 11].Value and \
#     #                 sheet.Cells[row1,16].Value == sheet.Cells[row, 17].Value:
#     #             # if    DB.XYZ(loc[0],loc[1],loc[2]).DistanceTo(DB.XYZ(loc1[0],loc1[1],loc1[2])) < 1000 / 304.8:
                    
#     #                 sheet.Cells[row1,2].Value = sheet.Cells[row, 2].Value
#     #                 sheet.Cells[row1,5].Value = sheet.Cells[row, 5].Value
#     #                 sheet.Cells[row1,8].Value = sheet.Cells[row, 8].Value
#     #                 sheet.Cells[row1,11].Value = sheet.Cells[row, 11].Value
#     #                 sheet.Cells[row1,14].Value = sheet.Cells[row, 14].Value
#     #                 sheet.Cells[row1,17].Value = sheet.Cells[row, 17].Value
#     #                 sheet.Cells[row1,20].Value = sheet.Cells[row, 20].Value
#     #                 sheet.Cells[row1,23].Value = sheet.Cells[row, 23].Value
#     #                 sheet.Cells[row1,25].Value = sheet.Cells[row, 25].Value
#     #                 sheet.Cells[row1,24].Value = "Wenig als 100 cm"
#     #                 sheet.Cells[row, 2].Value = ""
#     #                 sheet.Cells[row, 3].Value = ""
#     #                 sheet.Cells[row, 5].Value = ""
#     #                 sheet.Cells[row, 8].Value = ""
#     #                 sheet.Cells[row, 11].Value = ""
#     #                 sheet.Cells[row, 14].Value = ""
#     #                 sheet.Cells[row, 17].Value = ""
#     #                 sheet.Cells[row, 20].Value = ""
#     #                 sheet.Cells[row, 23].Value = ""
#     #                 sheet.Cells[row, 25].Value = ""

   
# book.Save()
# book.Dispose()




# #     book.Save()
# #     book.Dispose()

# # except Exception as e:
# #     book.Save()
# #     book.Dispose()

# # liste = []
# # cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
# # for el in cl:
# #     liste.append(BSK(el))
# # n = maxRowNum
# # for bsk in liste:
# #     if int(bsk.elemid) in Parameter_Dict.keys():
# #         row = Parameter_Dict[int(bsk.elemid)]
# #         sheet.Cells[row,2].Value = bsk.elemid
# #         sheet.Cells[row,3].Value = "Passt"

# #         sheet.Cells[row,5].Value = bsk.nummer0
# #         if sheet.Cells[row,4].Value == sheet.Cells[row,5].Value:
# #             sheet.Cells[row,6].Value = "Passt"
# #         else:
# #             sheet.Cells[row,6].Value = "Passt Nicht"

# #         sheet.Cells[row,8].Value = bsk.nummer1
# #         try:
# #             if sheet.Cells[row,8].Value == "MFC_" + sheet.Cells[row,7].Value:
# #                 sheet.Cells[row,9].Value = "Passt"
# #             else:
# #                 sheet.Cells[row,9].Value = "Passt Nicht"
# #         except:
# #             sheet.Cells[row,9].Value = "Passt Nicht"

# #         sheet.Cells[row,11].Value = bsk.Size
# #         if sheet.Cells[row,11].Value == sheet.Cells[row,10].Value:
# #             sheet.Cells[row,12].Value = "Passt"
# #         else:
# #             sheet.Cells[row,12].Value = "Passt Nicht"

# #         sheet.Cells[row,14].Value = bsk.Bedienseite
# #         if sheet.Cells[row,14].Value == sheet.Cells[row,13].Value:
# #             sheet.Cells[row,15].Value = "Passt"
# #         else:
# #             sheet.Cells[row,15].Value = "Passt Nicht"

# #         sheet.Cells[row,17].Value = bsk.Wirkunsort
# #         if sheet.Cells[row,17].Value == sheet.Cells[row,16].Value:
# #             sheet.Cells[row,18].Value = "Passt"
# #         else:
# #             sheet.Cells[row,18].Value = "Passt Nicht"

# #         sheet.Cells[row,20].Value = bsk.System
# #         if sheet.Cells[row,20].Value == sheet.Cells[row,19].Value:
# #             sheet.Cells[row,21].Value = "Passt"
# #         else:
# #             sheet.Cells[row,21].Value = "Passt Nicht"

# #         sheet.Cells[row,23].Value = bsk.wand
# #         if sheet.Cells[row,23].Value == sheet.Cells[row,22].Value:
# #             sheet.Cells[row,24].Value = "Passt"
# #         else:
# #             sheet.Cells[row,24].Value = "Passt Nicht"

# #         sheet.Cells[row,26].Value = bsk.Location
# #         if sheet.Cells[row,26].Value == sheet.Cells[row,25].Value:
# #             sheet.Cells[row,27].Value = "Passt"
# #         else:
# #             loc = list(eval(sheet.Cells[row,25].Value))
# #             xyz = DB.XYZ(loc[0],loc[1],loc[2])
# #             if xyz.DistanceTo(bsk.elem.Location.Point) < 100/304.8:
# #                 sheet.Cells[row,27].Value  = "wenig als 10 cm"
# #             elif xyz.DistanceTo(bsk.elem.Location.Point) < 500/304.8:
# #                 sheet.Cells[row,27].Value = "wenig als 50 cm"
# #             else:
# #                 sheet.Cells[row,27].Value = "Passt nicht"
# #         sheet.Cells[row,28].Value = bsk.familytyp

# #     else:
# #         n += 1
# #         row = n
# #         sheet.Cells[row,2].Value = bsk.elemid
# #         sheet.Cells[row,3].Value = "Neu"

# #         sheet.Cells[row,5].Value = bsk.nummer0

# #         sheet.Cells[row,8].Value = bsk.nummer1

# #         sheet.Cells[row,11].Value = bsk.Size

# #         sheet.Cells[row,14].Value = bsk.Bedienseite

# #         sheet.Cells[row,17].Value = bsk.Wirkunsort

# #         sheet.Cells[row,20].Value = bsk.System

# #         sheet.Cells[row,23].Value = bsk.wand

# #         sheet.Cells[row,26].Value = bsk.Location
# #         sheet.Cells[row,28].Value = bsk.familytyp

# # book.Save()
# # book.Dispose()
# from IGF_Funktionen._Parameter import get_value
# cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
# meps = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
# dict_ = {el.Number:el.LookupParameter("Name").AsString() for el in meps}
# Dict = {}
# Dict["EG"] = DB.ElementId(17024597)
# Dict["1.OG"] = DB.ElementId(17024600)
# Dict["2.OG"] = DB.ElementId(17024603)
# Dict["3.OG"] = DB.ElementId(17024605)
# Dict["4.OG"] = DB.ElementId(17024627)
# Dict["1.UG"] = DB.ElementId(21885480)
# Dict["2.UG"] = DB.ElementId(17024632)

# class Trassen:
#     def __init__(self,elem):
#         self.elem = elem
#         self.Raum1 = ""
#         self.Raum2 = self.elem.LookupParameter("IGF_X_Wirkungsort").AsString()
#         # self.bsk = self.elem.LookupParameter("IGF_RLT_BSK-Nummer").AsString()
#         self.Raum3 = ""
#         self.Liste = []
#         # self.Connectors = self.elem.ConnectorManager.Connectors
#         self.get_Raum()
    
#     def get_Raum(self):
#         mep = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)]
#         if not mep:
#             boundingbox = self.elem.get_BoundingBox(None)
#             if boundingbox:
#                 location = (boundingbox.Max + boundingbox.Min) / 2
#                 mep = doc.GetSpaceAtPoint(location)
#         if mep:
#             self.Raum1 = mep.Number + " [" + mep.LookupParameter("Name").AsString() + "]"
#         # if self.bsk:
#         #     einbau =  self.elem.LookupParameter("IGF_X_Einbauort").AsString()
#         #     if einbau in dict_.keys():
#         #         self.Raum1 = einbau + " [" + dict_[einbau] + "]"
#         if self.Raum2:
#             if self.Raum2 in dict_.keys():
#                 self.Raum3 = self.Raum2 + " [" + dict_[self.Raum2] + "]"
#         # for conn in self.Connectors:
#         #     basisz = conn.CoordinateSystem.BasisZ
#         #     mep = doc.GetSpaceAtPoint(conn.Origin)
#         #     if not mep:
#         #         mep = doc.GetSpaceAtPoint(conn.Origin + 100 / 304.8 * basisz )
#         #     if not mep:
#         #         mep = doc.GetSpaceAtPoint(conn.Origin + 300 / 304.8 * basisz )
#         #     if mep:
#         #         text = mep.Number + " [" + mep.LookupParameter("Name").AsString() + "]"
#         #         if text not in self.Liste:self.Liste.append(text)
        
#         # for el in self.Liste:
#         #     if not self.Raum1:self.Raum1 = el
#         #     else:self.Raum2 = el
    
#     def wert_schreiben(self):
#         self.elem.LookupParameter("IGF_X_Einbauort_0").Set(self.Raum1)
#         self.elem.LookupParameter("IGF_X_Wirkungsort_0").Set(self.Raum3)
#         # # self.elem.LookupParameter("CAx_Trassenbezugsebene").Set(self.elem.ReferenceLevel.Name)
#         self.elem.LookupParameter("IGF_X_ElementId").Set(self.elem.Id.ToString())

#         # height = self.elem.Location.Point.Z
#         # ebene = ""
#         # if height <= -15.944881889763785:
#         #     ebene = "2.UG"
#         # elif height <= -1.18110236220472:
#         #     ebene = "1.UG"
#         # elif height <= 12.434383202099736:
#         #     ebene = "EG"
#         # elif height <= 26.049868766404199:
#         #     ebene = "1.OG"
#         # elif height <= 39.665354330708659:
#         #     ebene = "2.OG"
#         # elif height <= 53.280839895013109:
#         #     ebene = "3.OG"
#         # else:
#         #     ebene = "4.OG"
#         # try:
#         #     self.elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM).Set(Dict[ebene])
            
#         #     self.elem.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(self.elem.Location.Point.Z-doc.GetElement(Dict[ebene]).LookupParameter("Ansicht").AsDouble())

#         # except Exception as e:print(e) 
            


# t =DB.Transaction(doc,"1")
# t.Start()
# for el in cl:
#     Trassen(el).wert_schreiben()
# t.Commit()  

        