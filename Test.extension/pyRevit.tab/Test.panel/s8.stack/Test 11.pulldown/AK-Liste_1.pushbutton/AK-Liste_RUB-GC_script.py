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
        if len(self.list) > 1000:
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
                    if ref.ConnectorType.ToString() != 'End':continue

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
    
    @property
    def BauteilNummer(self):
        if self.BauteilId and self.elem:
            para = self.elem.LookupParameter(self.BauteilId)
            if para:
                return para.AsString()
            return ''
        return ''

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
    

    @property
    def Ausgabe(self):
        return [self.BauteilNummer,self.elem.Id.ToString(),self.SubItemText]
    
    # @property
    # def Ausgabe(self):
    #     try:return [self.BauteilNummer,self.bauteilnummer_aus_TGA, self.elem.Id.ToString(),self.SubItemText,self.Subitem_aus_TGA,self.Pruefen]
    #     except:return [self.BauteilNummer,self.bauteilnummer_aus_TGA, '',self.SubItemText,self.Subitem_aus_TGA,self.Pruefen]

    def werte_schreiben(self):
        self.elem.LookupParameter('IGF_X_Bauteil_ID_Text').Set(self._bauteilnummer_aus_TGA)
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

# HK_Liste = [Endverbraucher(elem,True,'Regelventil','6_Wege','IGF_X_Bauteil_ID_Text') for elem in HK_Segel]
# H_Liste = [Endverbraucher(elem,True,'Regelventil',None,'IGF_X_Bauteil_ID_Text') for elem in H_Segel]
        
Segel =   DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'Deckensegel',False)))\
            .WherePasses(DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(ParamId,'segel',False)))\
            .WhereElementIsNotElementType()\
            .ContainedInDesignOption(DB.ElementId(-1))\
            .ToElements()

HK_Segel = []
H_Segel = []
for el in Segel:
    sysname = el.LookupParameter('Systemname').AsString()
    if sysname.find('K VL') != -1:
        HK_Segel.append(el)
    elif sysname.find('H VL') != -1:
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

HK_Liste = [Endverbraucher(elem,True,'Regelventil','Wege','IGF_X_Bauteil_ID_Text') for elem in HK_Segel]
H_Liste = [Endverbraucher(elem,True,'Regelventil',None,'IGF_X_Bauteil_ID_Text') for elem in H_Segel]
['5367686', '9982233', '9980647', '9980649', '13479897', '13479906', '13479900', '13480681', '13480316', '13480554', '13478743', '13480551', '13479891', '9980645', '9980659', '9982230', '13478740', '13480548', '13479921', '9980679', '9980681', '9982227', '13480313', '13480684', '13479912', '13479918', '13479909', '9980663', '9980661', '9982224']
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
    regel = Regelkomponent(elem,dict_regelventile[elid],'IGF_X_Bauteil_ID_Text')
    Liste_Regelventil.append(regel.Ausgabe)
    if not regel.BauteilNummer:
        print(regel.elem.Id.IntegerValue)
    if regel.BauteilNummer not in Ids_Regelventil:
        Ids_Regelventil.append(regel.BauteilNummer)
    else:
        print(regel.BauteilNummer,regel.elem.Id.IntegerValue)
print('---')
for elid in dict_6_wege_ventile.keys():
    elem = doc.GetElement(DB.ElementId(int(elid)))
    regel = Regelkomponent(elem,dict_6_wege_ventile[elid],'IGF_X_Bauteil_ID_Text')
    Liste_6Wegeventil.append(regel.Ausgabe)
    if not regel.BauteilNummer:
        print(regel.elem.Id.IntegerValue)
    if regel.BauteilNummer not in Ids_6Wegeventil:
        Ids_6Wegeventil.append(regel.BauteilNummer)
    else:
        print(regel.BauteilNummer,regel.elem.Id.IntegerValue)

Liste_Regelventil.sort()
Liste_6Wegeventil.sort()
Liste_Regelventil.insert(0,['Bauteilnummer','RevitId','DS_Nummer'])
Liste_6Wegeventil.insert(0,['Bauteilnummer','RevitId','DS_Nummer'])


# Excel Export
path = r"C:\Users\zhang\Desktop\1234567.xlsx"
workbook = xlsxwriter.Workbook(path)
worksheet = workbook.add_worksheet('Regelventil')
worksheet2 = workbook.add_worksheet('6-Wege-Ventil')

for row in range(len(Liste_Regelventil)):
    worksheet.write(row, 0, Liste_Regelventil[row][0])
    worksheet.write(row, 1, Liste_Regelventil[row][1])
    worksheet.write(row, 2, Liste_Regelventil[row][2])

for row in range(len(Liste_6Wegeventil)):
    worksheet2.write(row, 0, Liste_6Wegeventil[row][0])
    worksheet2.write(row, 1, Liste_6Wegeventil[row][1])
    worksheet2.write(row, 2, Liste_6Wegeventil[row][2])

worksheet.freeze_panes(1, 0)
worksheet2.freeze_panes(1, 0)
workbook.close()