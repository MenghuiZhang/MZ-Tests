# coding: utf8
import os
from pyrevit import revit, DB
from pyrevit import script
import xlsxwriter
from AK_Liste_GUI import Excelerstellen
from IGF_forms import ExcelSuche
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext

__title__ = "Export Deckensegel_Schema"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
logger = script.get_logger()

projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number
config = script.get_config('Deckensegel' + projectinfo)

adresse = ''

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = ""
except:
    pass

# excel = Excelerstellen(exceladresse = adresse)
# try:
#     excel.ShowDialog()
# except Exception as e:
#     logger.error(e)
#     excel.Close()
#     script.exit()

# try:
#     config.adresse = excel.Adresse.Text
#     script.save_config()
# except:
#     logger.error('kein Excel gegeben')
#     script.exit()


projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number

config = script.get_config('Schema-Deckensegel -' + projectinfo)
adresse = 'Excel Adresse'

try:
    adresse = config.adresse
    if not os.path.exists(config.adresse):
        config.adresse = ''
        adresse = "Excel Adresse"
except:
    pass

ExcelWPF = ExcelSuche(exceladresse = adresse)
ExcelWPF.Title = 'MEP-Räume Bauteilliste'
ExcelWPF.ShowDialog()
try:
    config.adresse = ExcelWPF.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel ausgewählt!')
    script.exit()

dict_regelventile_excel = {}

dict_6_wege_ventile_excel = {}

ExcelPackage.LicenseContext = LicenseContext.NonCommercial
fs = FileStream(ExcelWPF.Adresse.Text,FileMode.Open,FileAccess.Read)
book = ExcelPackage(fs)

try:
    for sheet in book.Workbook.Worksheets:
        if sheet.Name == 'IGF_L':
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                liste = []
                nummer = sheet.Cells[row, 3].Value
                subnummer = sheet.Cells[row, 4].Value


                dict_regelventile_excel[nummer] = subnummer
        # else:
        #     maxRowNum = sheet.Dimension.End.Row
        #     for row in range(2, maxRowNum + 1):
        #         liste = []
        #         nummer = sheet.Cells[row, 1].Value
        #         subnummer = sheet.Cells[row, 3].Value
        #         dict_6_wege_ventile_excel[subnummer] = nummer
           
                
            
    book.Save()
    book.Dispose()
    fs.Close()
    fs.Dispose()
except Exception as e:
    logger.error(e)
    book.Save()
    book.Dispose()
    fs.Close()
    fs.Dispose()
    script.exit()

print(dict_regelventile_excel)
import sys
sys.exit()

class Endverbraucher:
    def __init__(self, elem, IstVorlauf = True, Regelventil = None, Wege_6 = None, BauteilId = None):
        self.elem = elem
        self.IstVorlauf = IstVorlauf
        self.elemid = self.elem.Id
        self.doc = self.elem.Document
        self.list = []
        self.Regelventil_Text = Regelventil
        self.Wege_6_Text = Wege_6
        self.regelventil = None
        self.sechswege = None
        self.BauteilId = BauteilId
        if self.Regelventil_Text:
            self.get_Ventile(self.elem)
        
        # if self.Regelventil_Text and not self.regelventil:
        #     print('{} hat kein regelventil'.format(self.elemid.ToString()))
        
        # if self.Wege_6_Text and not self.sechswege:
        #     print('{} hat kein 6-Wege-Ventil'.format(self.elemid.ToString()))
    
    @property
    def BauteilNummer(self):
        if self.BauteilId:
            para = self.elem.LookupParameter(self.BauteilId)
            if para:
                return para.AsString()
            return ''
        return ''
    
    @property
    def Phase(self):
        return self.doc.GetElement(self.elem.CreatedPhaseId)

    @property
    def Raum(self):
        if self.Phase:
            return self.elem.Space[self.Phase]
    
    @property
    def Raumnummer(self):
        if self.Raum:
            return self.Raum.Number
        else:
            return ''
    
    @property
    def Raumname(self):
        if self.Raum:
            return self.Raum.LookupParameter('Name').AsString()
        else:
            return ''
    
    @property
    def Ebene(self):
        if self.Raum:
            return self.Raum.LookupParameter('Ebene').AsValueString()
        else:
            return ''
    
    def get_Ventile(self, elem):
        '''Ermittlung der Ventil Wirkungsorte'''
        if not self.Wege_6_Text and self.regelventil:
            return 
        if self.sechswege:
            return
        if len(self.list) > 300:
            return
        elemid = elem.Id.ToString()
        self.list.append(elemid)
        cate = elem.Category.Name
        if not cate in ['Rohr Systeme', 'Rohrdämmung']:
            try:   conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:return

            if conns.Size > 8:
                return
            
            for conn in conns:
                if elemid == self.elemid.ToString():
                    try:
                        if self.IstVorlauf:
                            if conn.PipeSystemType.ToString() == 'ReturnHydronic':
                                continue
                        else:
                            if conn.PipeSystemType.ToString() == 'SupplyHydronic':
                                continue
                    except:
                        continue

                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner

                    if not owner.Id.ToString() in self.list:
                        if owner.Category.Name == 'HLS-Bauteile':
                            return
                        if owner.Category.Name == 'Rohrzubehör':
                            faminame = owner.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
                            if self.Regelventil_Text and faminame.find(self.Regelventil_Text) != -1:
                                self.regelventil = owner.Id.ToString()
                                if not self.Wege_6_Text:
                                    return
                            elif self.Wege_6_Text and faminame.find(self.Wege_6_Text) != -1:
                                self.sechswege = owner.Id.ToString()
                                return
                        self.get_Ventile(owner)

class Regelkomponent:
    def __init__(self,elem,Liste,BauteilId = None):
        self.elem = elem
        self.liste = Liste
        self.BauteilId = BauteilId
        self.Subitem_aus_TGA = ''
        self._bauteilnummer_aus_TGA = ''
    
    @property
    def Pruefen(self):
        return self.Subitem_aus_TGA == self.SubItemText
    
    # @property
    # def BauteilNummer(self):
    #     if self.BauteilId and self.elem:
    #         para = self.elem.LookupParameter(self.BauteilId)
    #         if para:
    #             return para.AsString()
    #         return ''
    #     return ''

    # @property
    # def bauteilnummer_aus_TGA(self):
    #     if self.BauteilNummer and self.Subitem_aus_TGA:
    #         return self.BauteilNummer
    #     else:
    #         return self._bauteilnummer_aus_TGA
    
    @property
    def SubItem(self):
        Liste = []
        for el in self.liste:
            Liste.append(el.BauteilNummer)
        return sorted(Liste)
    @property
    def SubItemText(self):
        text = ''
        if len(self.SubItem) == 0:
            return text
        for el in self.SubItem:
            text += el + ', '
        return text[:-2]
    
    # @property
    # def Ausgabe(self):
    #     try:return [self.BauteilNummer,self.bauteilnummer_aus_TGA, self.elem.Id.ToString(),self.SubItemText,self.Subitem_aus_TGA,self.Pruefen]
    #     except:return [self.BauteilNummer,self.bauteilnummer_aus_TGA, '',self.SubItemText,self.Subitem_aus_TGA,self.Pruefen]

    def werte_schreiben(self):
        self.elem.LookupParameter('IGF_X_Bauteilnummerierung').Set(self._bauteilnummer_aus_TGA)
ParamId = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)


# HK_Segel =   DB.FilteredElementCollector(doc)\
#             .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'Deckensegel',False)))\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'kühlen',False)))\
#             .WhereElementIsNotElementType()\
#             .ToElements()
# H_Segel =    DB.FilteredElementCollector(doc)\
#             .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'Deckensegel',False)))\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'heizen',False)))\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateNotContainsRule(ParamId,'kühlen',False)))\
#             .WhereElementIsNotElementType()\
#             .ToElements()

# HK_Liste = [Endverbraucher(elem,True,'Regelventil','6_Wege','IGF_X_Bauteilnummerierung') for elem in HK_Segel]
# H_Liste = [Endverbraucher(elem,True,'Regelventil',None,'IGF_X_Bauteilnummerierung') for elem in H_Segel]
        
Segel =   DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
            .WhereElementIsNotElementType()\
            .ContainedInDesignOption(DB.ElementId(-1))\
            .ToElements()

HK_Segel = []
H_Segel = []
for el in Segel:
    sysname = el.LookupParameter('Systemname').AsString()
    if sysname.find('KLD') != -1:
        HK_Segel.append(el)
    elif sysname.find('HZD') != -1:
        H_Segel.append(el)
    else:
        print(el.Id)

# class Segel:
#     def __init__(self,elem):
#         self.elem = elem
#         self.beschriftung = self.get_beschriftung()
    
#     def get_beschriftung(self):
#         elems = self.elem.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_MechanicalEquipmentTags))
#         for elid in elems:
#             elem = doc.GetElement(elid)
#             if elem.OwnerViewId.IntegerValue == 34663323:
#                 return elem
    
#     def change_beschriftungHK(self):
#         if self.beschriftung:
#             try:
#                 self.beschriftung.ChangeTypeId(DB.ElementId(36268918))
#             except:
#                 pass
#     def change_beschriftungH(self):
#         if self.beschriftung:
#             try:
#                 self.beschriftung.ChangeTypeId(DB.ElementId(34564463))
#             except:
#                 pass

# t = DB.Transaction(doc,'1')
# t.Start()
# for el in H_Segel:
#     Segel(el).change_beschriftungH()
# for el in HK_Segel:
#     Segel(el).change_beschriftungHK()
# t.Commit()
# H_Segel =    DB.FilteredElementCollector(doc)\
#             .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel_h',False)))\
#             .WhereElementIsNotElementType()\
#             .ToElements()

HK_Liste = [Endverbraucher(elem,True,'ABQM','DET_W_V3W_R3','IGF_X_Bauteilnummerierung') for elem in HK_Segel]
H_Liste = [Endverbraucher(elem,True,'ABQM',None,'IGF_X_Bauteilnummerierung') for elem in H_Segel]

dict_regelventile = {}

dict_6_wege_ventile = {}

for el in HK_Liste:
    if el.regelventil:
        if el.regelventil not in dict_regelventile.keys():
            dict_regelventile[el.regelventil] = []
        dict_regelventile[el.regelventil].append(el)
    if el.sechswege:
        if el.sechswege not in dict_6_wege_ventile.keys():
            dict_6_wege_ventile[el.sechswege] = []
        dict_6_wege_ventile[el.sechswege].append(el)

for el in H_Liste:
    if el.regelventil:
        if el.regelventil not in dict_regelventile.keys():
            dict_regelventile[el.regelventil] = []
        dict_regelventile[el.regelventil].append(el)


Liste_Regelventil = []
Liste_6Wegeventil = []

Ids_Regelventil = []
Ids_6Wegeventil = []
for elid in dict_regelventile.keys():
    elem = doc.GetElement(DB.ElementId(int(elid)))
    regel = Regelkomponent(elem,dict_regelventile[elid],'IGF_X_Bauteilnummerierung')
    if regel.SubItemText in dict_regelventile_excel.keys():
        regel._bauteilnummer_aus_TGA = dict_regelventile_excel[regel.SubItemText]
    else:
        print(elid)

    Liste_Regelventil.append(regel)


for elid in dict_6_wege_ventile.keys():
    elem = doc.GetElement(DB.ElementId(int(elid)))
    regel = Regelkomponent(elem,dict_6_wege_ventile[elid],'IGF_X_Bauteilnummerierung')
    if regel.SubItemText in dict_6_wege_ventile_excel.keys():
        regel._bauteilnummer_aus_TGA = dict_6_wege_ventile_excel[regel.SubItemText]
    Liste_6Wegeventil.append(regel)

z = DB.Transaction(doc)
z.Start('1')
for el in Liste_Regelventil:
    el.werte_schreiben()
for el in Liste_6Wegeventil:
    el.werte_schreiben()
z.Commit()

# for nummer in dict_regelventile_excel.keys():
#     if nummer not in Ids_Regelventil:

#         regel = Regelkomponent(None,[],'')
#         regel.bauteilnummer_aus_TGA = nummer
#         regel.Subitem_aus_TGA = dict_regelventile_excel[nummer]
#         Liste_Regelventil.append(regel.Ausgabe)

# for nummer in dict_6_wege_ventile_excel.keys():
#     if nummer not in Ids_6Wegeventil:
#         regel = Regelkomponent(None,[],'')
#         regel.bauteilnummer_aus_TGA = nummer
#         regel.Subitem_aus_TGA = dict_6_wege_ventile_excel[nummer]
#         Liste_6Wegeventil.append(regel.Ausgabe)


# Liste_Regelventil.sort()
# Liste_6Wegeventil.sort()
# Liste_Regelventil.insert(0,['Bauteilnummer','TGA_Nummer','RevitId','DS_Nummer','DS_TGA','Prüfen'])
# Liste_6Wegeventil.insert(0,['Bauteilnummer','TGA_Nummer','RevitId','DS_Nummer','DS_TGA','Prüfen'])


# # Excel Export
# path = excel.Adresse.Text
# workbook = xlsxwriter.Workbook(path)
# worksheet = workbook.add_worksheet('Regelventil')
# worksheet2 = workbook.add_worksheet('6-Wege-Ventil')

# for row in range(len(Liste_Regelventil)):
#     try:worksheet.write(row, 0, Liste_Regelventil[row][0])
#     except:pass
#     try:worksheet.write(row, 1, Liste_Regelventil[row][1])
#     except:pass
#     try:worksheet.write(row, 2, Liste_Regelventil[row][2])
#     except:pass
#     try:worksheet.write(row, 3, Liste_Regelventil[row][3])
#     except:pass
#     try:worksheet.write(row, 4, Liste_Regelventil[row][4])
#     except:pass
#     try:worksheet.write(row, 5, Liste_Regelventil[row][5])
#     except:pass

# for row in range(len(Liste_6Wegeventil)):
#     try:worksheet2.write(row, 0, Liste_6Wegeventil[row][0])
#     except:pass
#     try:worksheet2.write(row, 1, Liste_6Wegeventil[row][1])
#     except:pass
#     try:worksheet2.write(row, 2, Liste_6Wegeventil[row][2])
#     except:pass
#     try:worksheet2.write(row, 3, Liste_6Wegeventil[row][3])
#     except:pass
#     try:worksheet2.write(row, 4, Liste_6Wegeventil[row][4])
#     except:pass
#     try:worksheet2.write(row, 5, Liste_6Wegeventil[row][5])
#     except:pass

# worksheet.freeze_panes(1, 0)
# worksheet2.freeze_panes(1, 0)
# workbook.close()