# coding: utf8
import os
from pyrevit import revit, DB
from pyrevit import script
import xlsxwriter
from AK_Liste_GUI import Excelerstellen
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
__title__ = "Export Deckensegel"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

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
        
        if self.Regelventil_Text and not self.regelventil:
            print('{} hat kein regelventil'.format(self.elemid.ToString()))
        
        if self.Wege_6_Text and not self.sechswege:
            print('{} hat kein 6-Wege-Ventil'.format(self.elemid.ToString()))
    
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
                        elif self.IstVorlauf == False:
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

class Deckensegel:
    def __init__(self, elem, BauteilId = 'SBI_Bauteilnummerierung'):
        self.elem = elem
        self.elemid = self.elem.Id
        self.doc = self.elem.Document
        self.list = []
        self.regelventil = None
        self.sechswege = None
        self.BauteilId = BauteilId
        self.get_Ventile(self.elem)
        if not (self.regelventil and self.sechswege):
            print(self.elemid)
        
    
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
        if len(self.list) > 600:
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
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner

                    if not owner.Id.ToString() in self.list:
                        if owner.Category.Name == 'HLS-Bauteile':
                            return
                        if owner.Category.Name == 'Rohrzubehör':
                            faminame = owner.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
                            if faminame.find('ABQM') != -1:
                                self.regelventil = owner.Id.ToString()
                            elif faminame.find('Wege') != -1:
                                if faminame.find('K-RL') != -1:
                                    self.sechswege = owner.Id.ToString()
                                return
                        self.get_Ventile(owner)

class Deckensegel_Schema:
    def __init__(self, elem, BauteilId = 'IGF_X_Bauteil_ID_Text'):
        self.elem = elem
        self.elemid = self.elem.Id
        self.doc = self.elem.Document
        self.list = []
        self.regelventil = None
        self.sechswege = None
        self.BauteilId = BauteilId
        self.get_Ventile(self.elem)
        if not (self.regelventil and self.sechswege):
            print(self.elemid)
        
    
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
        if len(self.list) > 600:
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
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner

                    if not owner.Id.ToString() in self.list:
                        if owner.Category.Name == 'HLS-Bauteile':
                            return
                        if owner.Category.Name == 'Rohrzubehör':
                            faminame = owner.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
                            if faminame.find('ABQM') != -1:
                                self.regelventil = owner.Id.ToString()
                            elif faminame.find('W_V3W') != -1:
                                if owner.LookupParameter('M_HOR').AsInteger() == 1:
                                    self.sechswege = owner.Id.ToString()
                                return
                        self.get_Ventile(owner)

def export_DIA():
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Endverbraucher(elem,False,'ABQM',None,'SBI_Bauteilnummerierung') for elem in cl]
    liste = [['ID','ID2']]
    for el in HK_Liste:
        liste.append([el.BauteilNummer,doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('SBI_Bauteilnummerierung').AsString()])
    
    path = r'C:\Users\zhang\Desktop\DeckenInduktionskühler.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(2):
            worksheet.write(row, col, liste[row][col])
    workbook.close()

def export_FFU():
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Endverbraucher(elem,False,'ABQM',None,'SBI_Bauteilnummerierung') for elem in cl]
    liste = [['ID','ID2']]
    for el in HK_Liste:
        liste.append([el.BauteilNummer,doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('SBI_Bauteilnummerierung').AsString()])
    
    path = r'C:\Users\zhang\Desktop\FFU.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(2):
            worksheet.write(row, col, liste[row][col])
    workbook.close()

def export_Deckensegel():
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Deckensegel(elem) for elem in cl]
    liste = [['ID','ID2','ID3']]
    for el in HK_Liste:
        if not (el.regelventil and el.sechswege):
            print(el.elemid,0)
            print(el.regelventil,1)
            print(el.sechswege,2)
            continue
        liste.append([el.BauteilNummer,doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('SBI_Bauteilnummerierung').AsString(),doc.GetElement(DB.ElementId(int(el.sechswege))).LookupParameter('SBI_Bauteilnummerierung').AsString()])
    
    path = r'C:\Users\zhang\Desktop\Deckensegel.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(3):
            worksheet.write(row, col, liste[row][col])
    workbook.close()

def export_KD():
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Deckensegel(elem) for elem in cl]
    liste = [['ID','ID2','ID3']]
    for el in HK_Liste:
        liste.append([el.BauteilNummer,doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('SBI_Bauteilnummerierung').AsString(),doc.GetElement(DB.ElementId(int(el.sechswege))).LookupParameter('SBI_Bauteilnummerierung').AsString()])
    
    path = r'C:\Users\zhang\Desktop\Kühldecken.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(3):
            worksheet.write(row, col, liste[row][col])
    workbook.close()

def import_DIA():
    _dict = {}
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(r'C:\Users\zhang\Desktop\DeckenInduktionskühler.xlsx',FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)

    try:
        for sheet in book.Workbook.Worksheets:
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                liste = []
                nummer = sheet.Cells[row, 1].Value
                subnummer = sheet.Cells[row, 2].Value
                if nummer in _dict.keys():
                    print(nummer)
                else:
                    _dict[nummer] = subnummer
                    
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
    
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Endverbraucher(elem,False,'ABQM',None,'IGF_X_Bauteil_ID_Text') for elem in cl]
    t = DB.Transaction(doc,'test')
    t.Start()
    for el in HK_Liste:
        if el.BauteilNummer in _dict.keys():
            try:doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('ZS_MARK').Set(_dict[el.BauteilNummer])
            except Exception as e:print(e)
    t.Commit()
    
def import_FFU():
    _dict = {}
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(r'C:\Users\zhang\Desktop\FFU.xlsx',FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)

    try:
        for sheet in book.Workbook.Worksheets:
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                liste = []
                nummer = sheet.Cells[row, 1].Value
                subnummer = sheet.Cells[row, 2].Value
                if nummer in _dict.keys():
                    print(nummer)
                else:
                    _dict[nummer] = subnummer
                    
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
    
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Endverbraucher(elem,False,'ABQM',None,'IGF_X_Bauteil_ID_Text') for elem in cl]
    t = DB.Transaction(doc,'test')
    t.Start()
    for el in HK_Liste:
        if el.BauteilNummer in _dict.keys():
            try:doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('ZS_MARK').Set(_dict[el.BauteilNummer])
            except Exception as e:print(e)
    t.Commit()

def import_Deckensegel():
    _dict = {}
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(r'C:\Users\zhang\Desktop\Deckensegel.xlsx',FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)

    try:
        for sheet in book.Workbook.Worksheets:
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                liste = []
                nummer = sheet.Cells[row, 1].Value
                subnummer = sheet.Cells[row, 2].Value
                if nummer in _dict.keys():
                    print(nummer)
                else:
                    _dict[nummer] = subnummer
                    
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
    
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Deckensegel(elem) for elem in cl]
    t = DB.Transaction(doc,'test')
    t.Start()
    for el in HK_Liste:
        
        if not (el.regelventil and el.sechswege):
            print(el.elemid,0)
            print(el.regelventil,1)
            print(el.sechswege,2)
            continue
        if el.BauteilNummer in _dict.keys():
    
            # try:doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('ZS_MARK').Set(_dict[el.BauteilNummer])
            # except Exception as e:print(e)
            try:doc.GetElement(DB.ElementId(int(el.sechswege))).LookupParameter('SBI_Bauteilnummerierung').Set(_dict[el.BauteilNummer].replace('RVD','6WV'))
            except Exception as e:print(e)
        else:
            print(el.BauteilNummer)
        
            
    t.Commit()

def import_KD():
    _dict = {}
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(r'C:\Users\zhang\Desktop\Kühldecken.xlsx',FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)

    try:
        for sheet in book.Workbook.Worksheets:
            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                liste = []
                nummer = sheet.Cells[row, 1].Value
                subnummer = sheet.Cells[row, 2].Value
                if nummer in _dict.keys():
                    print(nummer)
                else:
                    _dict[nummer] = subnummer
                    
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
    
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WhereElementIsNotElementType()
    HK_Liste = [Deckensegel(elem) for elem in cl]
    t = DB.Transaction(doc,'test')
    t.Start()
    for el in HK_Liste:
        
        if not (el.regelventil and el.sechswege):
            print(el.elemid,0)
            print(el.regelventil,1)
            print(el.sechswege,2)
            continue
        if el.BauteilNummer in _dict.keys():
    
            # try:doc.GetElement(DB.ElementId(int(el.regelventil))).LookupParameter('ZS_MARK').Set(_dict[el.BauteilNummer])
            # except Exception as e:print(e)
            try:doc.GetElement(DB.ElementId(int(el.sechswege))).LookupParameter('SBI_Bauteilnummerierung').Set(_dict[el.BauteilNummer].replace('RVD','6WV'))
            except Exception as e:print(e)
        else:
            print(el.BauteilNummer)
        
            
    t.Commit()
# import_Deckensegel()
import_KD()
# import_FFU()  
# import_DIA() 
# export_FFU()   
# export_Deckensegel()
# export_KD()
# projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number
# config = script.get_config('Deckensegel' + projectinfo)

# def changetype(el,liste):
    
#     tags = el.GetDependentElements(DB.ElementClassFilter(DB.IndependenttTag))
#     for _id in tags:
#         tag = doc.GetElement(_id)
#         if tag.OwnerViewId.IntegerValue == 35336197:
#             if el.Space[doc.GetElement(el.CreatedPhaseId)].Id.IntegerValue in liste:tag.ChangeTypeId(DB.ElementId(36831953))
#             else:tag.ChangeTypeId(DB.ElementId(36822207))

#     # conns = el.MEPModel.ConnectorManager.Connectors
#     # for conn in conns:
#     #     refs = conn.AllRefs
#     #     for ref in refs:
#     #         if ref.Owner.Category.Name == 'Rohrformteile':
#     #             conns1 = ref.Owner.MEPModel.ConnectorManager.Connectors
                
#     #             try:
#     #                 for conn1 in conns1:
#     #                     ref1s = conn1.AllRefs
#     #                     for ref1 in ref1s:
#     #                         try:
#     #                             d = ref1.Owner.LookupParameter('Durchmesser').AsValueString()
#     #                             el.LookupParameter('MC_D').SetValueString(d)
#     #                             # symbols = el.Symbol.Family.GetFamilySymbolIds()
#     #                             # for s in symbols:
#     #                             #     name = doc.GetElement(s).get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
#     #                             #     if name.find('R30'+d)!=-1:
#     #                             #     # if name.find(d) != -1 and name.find(d+'0') == -1:
#     #                             #         el.ChangeTypeId(s)
#     #                             return
#     #                         except:pass
#     #             except Exception as e:
#     #                 print(e)
#     #                 pass
# # _Systemfilter = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(DB.ElementId(DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM),'Rücklauf',False))
# # _Systemfilter1 = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(DB.ElementId(DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM),'Vorlauf',False))
# _FamilyFilter = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM),'SBI_DET_Segel_K',False))
# # _MCFilter = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM),'MC_Connection',False))

# cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WherePasses(_FamilyFilter).ToElements()
# liste = []
# for el in cl:
#     if el.Symbol.FamilyName == 'SBI_DET_Segel_K_ohne Anschluss':
#         liste.append(el.Space[doc.GetElement(el.CreatedPhaseId)].Id.IntegerValue)
# liste = set(liste)
# t = DB.Transaction(doc,'1')
# t.Start()
# for el in cl:
#     if el.Symbol.FamilyName == 'SBI_DET_Segel_K':changetype(el)
# t.Commit()
# adresse = ''

# try:
#     adresse = config.adresse
#     if not os.path.exists(config.adresse):
#         config.adresse = ''
#         adresse = ""
# except:
#     pass

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


# class Endverbraucher:
#     def __init__(self, elem, IstVorlauf = True, Regelventil = None, Wege_6 = None, BauteilId = None):
#         self.elem = elem
#         self.IstVorlauf = IstVorlauf
#         self.elemid = self.elem.Id
#         self.doc = self.elem.Document
#         self.list = []
#         self.Regelventil_Text = Regelventil
#         self.Wege_6_Text = Wege_6
#         self.regelventil = None
#         self.sechswege = None
#         self.BauteilId = BauteilId
#         if self.Regelventil_Text:
#             self.get_Ventile(self.elem)
        
#         if self.Regelventil_Text and not self.regelventil:
#             print('{} hat kein regelventil'.format(self.elemid.ToString()))
        
#         if self.Wege_6_Text and not self.sechswege:
#             print('{} hat kein 6-Wege-Ventil'.format(self.elemid.ToString()))
    
#     @property
#     def BauteilNummer(self):
#         if self.BauteilId:
#             para = self.elem.LookupParameter(self.BauteilId)
#             if para:
#                 return para.AsString()
#             return ''
#         return ''
    
#     @property
#     def Phase(self):
#         return self.doc.GetElement(self.elem.CreatedPhaseId)

#     @property
#     def Raum(self):
#         if self.Phase:
#             return self.elem.Space[self.Phase]
    
#     @property
#     def Raumnummer(self):
#         if self.Raum:
#             return self.Raum.Number
#         else:
#             return ''
    
#     @property
#     def Raumname(self):
#         if self.Raum:
#             return self.Raum.LookupParameter('Name').AsString()
#         else:
#             return ''
    
#     @property
#     def Ebene(self):
#         if self.Raum:
#             return self.Raum.LookupParameter('Ebene').AsValueString()
#         else:
#             return ''
    
#     def get_Ventile(self, elem):
#         '''Ermittlung der Ventil Wirkungsorte'''
#         if not self.Wege_6_Text and self.regelventil:
#             return 
#         if self.sechswege:
#             return
#         if len(self.list) > 300:
#             return
#         elemid = elem.Id.ToString()
#         self.list.append(elemid)
#         cate = elem.Category.Name
#         if not cate in ['Rohr Systeme', 'Rohrdämmung']:
#             try:   conns = elem.MEPModel.ConnectorManager.Connectors
#             except:
#                 try:conns = elem.ConnectorManager.Connectors
#                 except:return

#             if conns.Size > 8:
#                 return
            
#             for conn in conns:
#                 if elemid == self.elemid.ToString():
#                     try:
#                         if self.IstVorlauf:
#                             if conn.PipeSystemType.ToString() == 'ReturnHydronic':
#                                 continue
#                         else:
#                             if conn.PipeSystemType.ToString() == 'SupplyHydronic':
#                                 continue
#                     except:
#                         continue

#                 refs = conn.AllRefs
#                 for ref in refs:
#                     owner = ref.Owner

#                     if not owner.Id.ToString() in self.list:
#                         if owner.Category.Name == 'HLS-Bauteile':
#                             return
#                         if owner.Category.Name == 'Rohrzubehör':
#                             faminame = owner.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
#                             if self.Regelventil_Text and faminame.find(self.Regelventil_Text) != -1:
#                                 self.regelventil = owner.Id.ToString()
#                                 if not self.Wege_6_Text:
#                                     return
#                             elif self.Wege_6_Text and faminame.find(self.Wege_6_Text) != -1:
#                                 self.sechswege = owner.Id.ToString()
#                                 return
#                         self.get_Ventile(owner)

# class Regelkomponent:
#     def __init__(self,elem,Liste,BauteilId = None):
#         self.elem = elem
#         self.liste = Liste
#         self.BauteilId = BauteilId

#     @property
#     def BauteilNummer(self):
#         if self.BauteilId:
#             para = self.elem.LookupParameter(self.BauteilId)
#             if para:
#                 return para.AsString()
#             return ''
#         return ''
    
#     @property
#     def SubItem(self):
#         Liste = []
#         for el in self.liste:
#             Liste.append(el.BauteilNummer)
#         return sorted(Liste)
#     @property
#     def SubItemText(self):
#         text = ''
#         if len(self.SubItem) == 0:
#             return text
#         for el in self.SubItem:
#             text += el + ', '
#         return text[:-2]
    
#     @property
#     def Ausgabe(self):
#         return [self.BauteilNummer,self.elem.Id.ToString(),self.SubItemText]


# ParamId = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)


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

# def get_text(liste):
#     text = ''
#     liste.sort()
#     if len(liste) == 0:
#         return text
#     for el in liste:
#         text += el + ', '
#     return text[:-2]

# def get_Liste_DES(Liste):
#     _dict = {}
#     liste = []
#     for el in Liste:
#         if el.Raum:
#             if el.Raum.Id.IntegerValue not in _dict.keys():
#                 _dict[el.Raum.Id.IntegerValue] = []
#             _dict[el.Raum.Id.IntegerValue].append(el.BauteilNummer)
    
#     for el in _dict.keys():
#         raumnummer = doc.GetElement(DB.ElementId(el)).Number
#         text = get_text(_dict[el])
#         liste.append([raumnummer,text,len(_dict[el])])
#     return liste


# def export(liste):
#     path = excel.Adresse.Text
#     workbook = xlsxwriter.Workbook(path)
#     worksheet = workbook.add_worksheet('DeS')

#     for row in range(len(liste)):
#         for col in range(len(liste[0])):
#             worksheet.write(row, col, liste[row][col])
#     workbook.close()

# def exportdeckensegel():
#     liste = []
#     liste.extend(get_Liste_DES(HK_Liste))
#     liste.extend(get_Liste_DES(H_Liste))
#     liste.sort()
#     liste.insert(0, ['raumnummer','des','anzahl'])
#     export(liste)

def export_HK_MEP():
    cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id)
    liste = [['ID','MEP']]
    for el in cl:
        mep = el.Space[doc.GetElement(el.CreatedPhaseId)]
        if mep:
            liste.append([el.LookupParameter('IGF_X_Bauteil_ID_Text').AsString(),mep.Number])
    
    path = r'C:\Users\zhang\Desktop\HK_MEP.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(2):
            worksheet.write(row, col, liste[row][col])
    workbook.close()

# for el in cl:
#     mep = el.Space[doc.GetElement(el.CreatedPhaseId)]
#     if not mep:
#         print(el.Id)
# export_HK_MEP()

def Import_HK_MEP():
    path = r'C:\Users\zhang\Desktop\IGF_X_Information ans Schema_Heizkörper Kopie 1_20.04.xlsx'
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(path,FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)
    _dict = {}
    try:
        for sheet in book.Workbook.Worksheets:

            maxRowNum = sheet.Dimension.End.Row
            for row in range(2, maxRowNum + 1):
                nummer = sheet.Cells[row, 1].Value
                subnummer = sheet.Cells[row, 2].Value

                _dict[nummer] = subnummer          
                
                
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

    elemids = uidoc.Selection.GetElementIds()
    t = DB.Transaction(doc,'Test')
    t.Start()
    for elemid in elemids:
        elem = doc.GetElement(elemid)
        Id = elem.LookupParameter('IGF_X_Bauteil_ID_Text').AsString()
        if Id in _dict.keys():
            mep = elem.Space[doc.GetElement(elem.CreatedPhaseId)]
            elem.LookupParameter('IGF_X_Einbauort').Set(_dict[Id])
            if mep:
                mep.Number = _dict[Id]
            else:
                print(elem.Id)
        else:
            print(elem.Id)
    t.Commit()
# Import_HK_MEP()
def export_HK_MEP_Schema():
    elemids = uidoc.Selection.GetElementIds()
    liste = [['RVT_Id','ID','MEP','Schema','Prüfen']]
    for elemid in elemids:
        el = doc.GetElement(elemid)
        mep = el.Space[doc.GetElement(el.CreatedPhaseId)]
        if mep:
            mepnummer = mep.Number
            Id_nummer = el.LookupParameter('IGF_X_Bauteil_ID_Text').AsString()
            einbauort = el.LookupParameter('IGF_X_Einbauort').AsString()
            liste.append([el.Id.ToString(),Id_nummer,mepnummer,einbauort,einbauort == mepnummer])
    
    path = r'C:\Users\zhang\Desktop\HK_MEP_Schema.xlsx'
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for row in range(len(liste)):
        for col in range(5):
            worksheet.write(row, col, liste[row][col])
    workbook.close()

# export_HK_MEP_Schema()
# exportdeckensegel()

# Ids_DeS = []
# for el in HK_Liste:
#     if el.BauteilNummer not in Ids_DeS:
#         Ids_DeS.append(el.BauteilNummer)
#     else:
#         print(el.BauteilNummer,el.elemid.IntegerValue)
#     if not el.BauteilNummer:
#         print(el.elemid)
# for el in H_Liste:
#     if el.BauteilNummer not in Ids_DeS:
#         Ids_DeS.append(el.BauteilNummer)
#     else:
#         print(el.BauteilNummer,el.elemid.IntegerValue)
#     if not el.BauteilNummer:
#         print(el.elemid)
# print('-------------------------')
        
# HK_Segel =   DB.FilteredElementCollector(doc)\
#             .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel_k',False)))\
#             .WhereElementIsNotElementType()\
#             .ToElements()
# H_Segel =    DB.FilteredElementCollector(doc)\
#             .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
#             .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel_h',False)))\
#             .WhereElementIsNotElementType()\
#             .ToElements()

# HK_Liste = [Endverbraucher(elem,True,'ABQM','DET_W_V3W_R3','IGF_X_Bauteilnummerierung') for elem in HK_Segel]
# H_Liste = [Endverbraucher(elem,True,'ABQM',None,'IGF_X_Bauteilnummerierung') for elem in H_Segel]

# liste = []
# liste1 = []
# for el in HK_Segel:
#     elem = Endverbraucher(el,True,'Regelventil','6_Wege','IGF_X_Bauteilnummerierung')
#     if elem.BauteilNummer not in liste:
#         liste.append(elem.BauteilNummer)
#     else:
#         liste1.append([elem.BauteilNummer,elem.elemid.ToString()])

# for el in H_Segel:
#     elem = Endverbraucher(el,True,'Regelventil',None,'IGF_X_Bauteilnummerierung')
#     if elem.BauteilNummer not in liste:
#         liste.append(elem.BauteilNummer)
#     else:
#         liste1.append([elem.BauteilNummer,elem.elemid.ToString()])
# liste1.sort()
# for el in liste1:
#     print(el)

# print(len(HK_Segel))
# print(len(H_Segel))

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
#     Liste_Regelventil.append(regel.Ausgabe)
#     if not regel.BauteilNummer:
#         print(regel.elem.Id.IntegerValue)
#     if regel.BauteilNummer not in Ids_Regelventil:
#         Ids_Regelventil.append(regel.BauteilNummer)
#     else:
#         print(regel.BauteilNummer,regel.elem.Id.IntegerValue)
# print('---')
# for elid in dict_6_wege_ventile.keys():
#     elem = doc.GetElement(DB.ElementId(int(elid)))
#     regel = Regelkomponent(elem,dict_6_wege_ventile[elid],'IGF_X_Bauteilnummerierung')
#     Liste_6Wegeventil.append(regel.Ausgabe)
#     if not regel.BauteilNummer:
#         print(regel.elem.Id.IntegerValue)
#     if regel.BauteilNummer not in Ids_6Wegeventil:
#         Ids_6Wegeventil.append(regel.BauteilNummer)
#     else:
#         print(regel.BauteilNummer,regel.elem.Id.IntegerValue)

# Liste_Regelventil.sort()
# Liste_6Wegeventil.sort()
# Liste_Regelventil.insert(0,['Bauteilnummer','RevitId','DS_Nummer'])
# Liste_6Wegeventil.insert(0,['Bauteilnummer','RevitId','DS_Nummer'])


# # Excel Export
# path = excel.Adresse.Text
# workbook = xlsxwriter.Workbook(path)
# worksheet = workbook.add_worksheet('Regelventil')
# worksheet2 = workbook.add_worksheet('6-Wege-Ventil')

# for row in range(len(Liste_Regelventil)):
#     worksheet.write(row, 0, Liste_Regelventil[row][0])
#     worksheet.write(row, 1, Liste_Regelventil[row][1])
#     worksheet.write(row, 2, Liste_Regelventil[row][2])

# for row in range(len(Liste_6Wegeventil)):
#     worksheet2.write(row, 0, Liste_6Wegeventil[row][0])
#     worksheet2.write(row, 1, Liste_6Wegeventil[row][1])
#     worksheet2.write(row, 2, Liste_6Wegeventil[row][2])

# worksheet.freeze_panes(1, 0)
# worksheet2.freeze_panes(1, 0)
# workbook.close()