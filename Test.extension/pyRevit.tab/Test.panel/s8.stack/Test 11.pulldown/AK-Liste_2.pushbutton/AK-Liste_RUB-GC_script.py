# coding: utf8
import os
from pyrevit import revit, DB
from pyrevit import script
import xlsxwriter
from AK_Liste_GUI import Excelerstellen
from IGF_forms import ExcelSuche
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext

__title__ = "Deckensegel_Schema_Abgleich"
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

excel = Excelerstellen(exceladresse = adresse)
try:
    excel.ShowDialog()
except Exception as e:
    logger.error(e)
    excel.Close()
    script.exit()

try:
    config.adresse = excel.Adresse.Text
    script.save_config()
except:
    logger.error('kein Excel gegeben')
    script.exit()


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
        
        maxRowNum = sheet.Dimension.End.Row
        for row in range(2, maxRowNum + 1):
            liste = []
            nummer = sheet.Cells[row, 1].Value
            subnummer = sheet.Cells[row, 2].Value
            anzahl = sheet.Cells[row, 3].Value
            if nummer in dict_regelventile_excel.keys():
                dict_regelventile_excel[nummer][0] += ', ' + subnummer
                dict_regelventile_excel[nummer][1] += anzahl
            else:
                dict_regelventile_excel[nummer] = [subnummer,anzahl]
                
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
        # if self.Regelventil_Text:
        #     self.get_Ventile(self.elem)
        
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
        self.Anzahl_TGA = 0
        self._bauteilnummer_aus_TGA = ''
    
    @property
    def Pruefen(self):
        return self.Subitem_aus_TGA == self.SubItemText
    
    @property
    def BauteilNummer(self):
        if self.BauteilId and self.elem:
            para = self.elem.LookupParameter(self.BauteilId)
            if para:
                return para.AsString()
            return ''
        return ''

    @property
    def bauteilnummer_aus_TGA(self):
        if self.BauteilNummer and self.Subitem_aus_TGA:
            return self.BauteilNummer
        else:
            return self._bauteilnummer_aus_TGA
    
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
    
    @property
    def Ausgabe(self):
        try:return [self.BauteilNummer,self.bauteilnummer_aus_TGA, self.elem.Id.ToString(),self.SubItemText,self.Subitem_aus_TGA,self.Pruefen]
        except:return [self.BauteilNummer,self.bauteilnummer_aus_TGA, '',self.SubItemText,self.Subitem_aus_TGA,self.Pruefen]

class MEPRaum:
    def __init__(self,elem,liste):
        self.elem = elem
        self.SubItem = liste
        self.SubItem.sort()
        self.Subitem_aus_TGA = ''
        self.nummer = self.elem.Number
    
    @property
    def Pruefen(self):
        if self.Subitem_aus_TGA == self.SubItemText:
            return 0
        else:
            return 1
    

    @property
    def SubItem(self):
        Liste = []
        for el in self.liste:
            Liste.append(el.BauteilNummer)
        return sorted(Liste)
    @property
    def Anzahl(self):
        return len(self.SubItem)

    @property
    def SubItemText(self):
        text = ''
        if len(self.SubItem) == 0:
            return text
        for el in self.SubItem:
            text += el + ', '
        return text[:-2]
    
    def wert_Schreiben(self):
        self.elem.LookupParameter('IGF_DSP_Schema').Set(self.SubItemText)
        self.elem.LookupParameter('IGF_DSP_TGA').Set(self.Subitem_aus_TGA)
        self.elem.LookupParameter('IGF_MEP_Überprüfen').Set(self.Pruefen)
        self.elem.LookupParameter('IGF_DSP_Anzahl_Schema').Set(self.Anzahl)
        try:self.elem.LookupParameter('IGF_DSP_Anzahl_TGA').Set(int(self.Anzahl_TGA))
        except:self.elem.LookupParameter('IGF_DSP_Anzahl_TGA').Set(0)

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
        
HK_Segel =   DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel_k',False)))\
            .WhereElementIsNotElementType()\
            .ContainedInDesignOption(DB.ElementId(-1))\
            .ToElements()
H_Segel =    DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel_h',False)))\
            .WhereElementIsNotElementType()\
            .ContainedInDesignOption(DB.ElementId(-1))\
            .ToElements()

HK_Liste = [Endverbraucher(elem,True,'ABQM','DET_W_V3W_R3','IGF_X_Bauteilnummerierung') for elem in HK_Segel]
H_Liste = [Endverbraucher(elem,True,'ABQM',None,'IGF_X_Bauteilnummerierung') for elem in H_Segel]

def get_text(liste):
    text = ''
    liste.sort()
    if len(liste) == 0:
        return text
    for el in liste:
        text += el + ', '
    return text[:-2]

def get_Liste_DES(Liste):
    _dict = {}
    liste = []
    for el in Liste:
        if el.Raum:
            if el.Raum.Id.IntegerValue not in _dict.keys():
                _dict[el.Raum.Id.IntegerValue] = []
            _dict[el.Raum.Id.IntegerValue].append(el.BauteilNummer)
        else:
            print(el.Id)
    
    return _dict


def funktion0():
    _dict0 = get_Liste_DES(HK_Liste)
    _dict1 = get_Liste_DES(H_Liste)
    Liste = []
    if len(_dict0) > 0:
        for el in _dict0.keys():
            elem = doc.GetElement(DB.ElementId(el))
            mepraum = MEPRaum(elem, _dict0[el])
            if mepraum.nummer in dict_regelventile_excel.keys():
                mepraum.Subitem_aus_TGA = dict_regelventile_excel[mepraum.nummer][0]
                mepraum.Anzahl_TGA = dict_regelventile_excel[mepraum.nummer][1]
            Liste.append(mepraum)
    
    liste0 = []
    for el in Liste:liste0.append(el.nummer)
    for el in dict_regelventile_excel.keys():
        if el not in liste0:
            print(el)
    t = DB.Transaction(doc,'Test')
    t.Start()
    for el in Liste:
        el.wert_Schreiben()
    t.Commit()
    t.Dispose()


funktion0()





# dict_regelventile = {}

# dict_6_wege_ventile = {}

# for el in HK_Liste:
#     if el.regelventil:
#         if el.regelventil not in dict_regelventile.keys():
#             dict_regelventile[el.regelventil] = []
#         dict_regelventile[el.regelventil].append(el)
#     if el.sechswege:
#         if el.sechswege not in dict_6_wege_ventile.keys():
#             dict_6_wege_ventile[el.sechswege] = []
#         dict_6_wege_ventile[el.sechswege].append(el)

# for el in H_Liste:
#     if el.regelventil:
#         if el.regelventil not in dict_regelventile.keys():
#             dict_regelventile[el.regelventil] = []
#         dict_regelventile[el.regelventil].append(el)


# Liste_Regelventil = []
# Liste_6Wegeventil = []

# Ids_Regelventil = []
# Ids_6Wegeventil = []
# for elid in dict_regelventile.keys():
#     elem = doc.GetElement(DB.ElementId(int(elid)))
#     regel = Regelkomponent(elem,dict_regelventile[elid],'IGF_X_Bauteilnummerierung')
#     if regel.BauteilNummer in dict_regelventile_excel.keys():
#         regel.Subitem_aus_TGA = dict_regelventile_excel[regel.BauteilNummer]

#     Liste_Regelventil.append(regel.Ausgabe)
#     Ids_Regelventil.append(regel.BauteilNummer)

# for elid in dict_6_wege_ventile.keys():
#     elem = doc.GetElement(DB.ElementId(int(elid)))
#     regel = Regelkomponent(elem,dict_6_wege_ventile[elid],'IGF_X_Bauteilnummerierung')
#     if regel.BauteilNummer in dict_6_wege_ventile_excel.keys():
#         regel.Subitem_aus_TGA = dict_6_wege_ventile_excel[regel.BauteilNummer]
#     Liste_6Wegeventil.append(regel.Ausgabe)
#     Ids_6Wegeventil.append(regel.BauteilNummer)

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