# coding: utf8
import os
from pyrevit import revit, UI, DB
from pyrevit import script
import xlsxwriter
from operator import itemgetter
from IGF_Elementfilter import ElementFamilyTypeSeparatelyContains_Filter
from System import Guid
from IGF_lib import get_value



__title__ = "6-Wege-Ventil an MC"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc
error = []
output= script.get_output()
logger = script.get_logger()

_Systemfilter = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(DB.ElementId(DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM),'Rücklauf',False))
_Systemfilter1 = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateEqualsRule(DB.ElementId(DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM),'Vorlauf',False))
_FamilyFilter = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM),'SBI_DET_W_V3W_R3',False))
_MCFilter = DB.ElementParameterFilter(DB.ParameterFilterRuleFactory.CreateContainsRule(DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM),'MC_Connection',False))

wege_ventil = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WherePasses(_FamilyFilter).WherePasses(_Systemfilter).ToElements()
print(len(wege_ventil))

p = DB.XYZ(1,5,10)

t = DB.Transaction(doc,'test')
t.Start()
for el in wege_ventil:

    outlin = DB.Outline(el.get_BoundingBox(None).Min-p,el.get_BoundingBox(None).Max+p)
    fil = DB.BoundingBoxIntersectsFilter(outlin,2)
    coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().WherePasses(_FamilyFilter).WherePasses(_Systemfilter1).ToElements()
    if coll.Count == 1:
        num = el.LookupParameter('ZS_MARK').AsString()
        # if coll[0].Id.IntegerValue in [39311981,39327120,39333619,36580989]:
         
        coll[0].LookupParameter('ZS_MARK').Set(num.replace('6WV','6WV'))
    else:
        print(el.get_BoundingBox(None).Min-p,el.get_BoundingBox(None).Max+p)
        
        print(el.Id)
        for el in coll:
            print(el.Id)
            print(el.get_BoundingBox(None).Min,el.get_BoundingBox(None).Max)
        break


t.Commit()



# hksegel_coll = ElementFamilyTypeSeparatelyContains_Filter('Deckensegel', 'kühlen',False, DB.BuiltInCategory.OST_MechanicalEquipment)
# MCNode_coll = ElementFamilyTypeSeparatelyContains_Filter('MC_ConnectionNode_Hydronic', 'MC_Connection',False, DB.BuiltInCategory.OST_MechanicalEquipment)
# dict_mc = {}
# for el in MCNode_coll:
#     space = el.Space[doc.GetElement(el.CreatedPhaseId)]
#     if space:
#         if space.Number not in dict_mc.keys():
#             dict_mc[space.Number] = []
#         dict_mc[space.Number].append(el)

# class HKSegel:
#     def __init__(self, elemid):
#         self.elemid = elemid
#         self.elem = doc.GetElement(self.elemid)
#         self.name = self.elem.Symbol.FamilyName + ': ' + self.elem.Name
#         self.typ = ''
#         try:
#             self.raumnummer = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
#         except:
#             logger.error('Kein MEP Raum {}'.format(self.elemid.ToString()))
#             self.raumnummer = ''
#         self.list = []
#         self.regelventil = None
#         self.sechswege = None
#         self.Wirkungsort(self.elem)
#         self.hl = get_value(self.elem.get_Parameter(Guid('49d7278f-3e2b-4f2e-b619-43e856e15be7')))
#         self.kl = get_value(self.elem.get_Parameter(Guid('fd3b6adb-78ec-413e-9fb4-dde3d1cea6ee')))
#         if self.typ == 'HK':
#             if not self.sechswege:
#                 logger.error('Kein 6WV für Element {}'.format(self.elemid.ToString()))
#         if not self.regelventil:
#             logger.error('Kein RV für Element {}'.format(self.elemid.ToString()))


    
#     def gettyp(self):
#         if self.name.upper().find('KÜHL') != -1:
#             self.typ = 'HK'
#             return 'HK'
#         else:
#             self.typ = 'H'
#             return 'H'
    
#     def Wirkungsort(self, elem):
#         if self.gettyp() == 'HK':
#             self.Wirkungsort_HK(elem)
#         else:
#             self.Wirkungsort_H(elem)

#     def Wirkungsort_HK(self, elem):
#         '''Ermittlung der Ventil Wirkungsorte'''
#         if self.sechswege:
#             return
#         elemid = elem.Id.ToString()
#         self.list.append(elemid)
#         cate = elem.Category.Name
#         if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Rohrdämmung', 'Luftkanaldämmung außen',
#                         'Luftkanaldämmung innen']:
#             conns = None
#             try:
#                 conns = elem.MEPModel.ConnectorManager.Connectors
#             except:
#                 try:
#                     conns = elem.ConnectorManager.Connectors
#                 except:
#                     pass

#             if conns:

#                 if conns.Size > 8:

#                     return
#                 for conn in conns:
#                     if elemid == self.elemid.ToString():

#                         try:
#                             if conn.PipeSystemType.ToString() == 'ReturnHydronic':
#                                 continue
#                         except:
#                             continue
#                     refs = conn.AllRefs
#                     for ref in refs:
#                         owner = ref.Owner

#                         if not owner.Id.ToString() in self.list:
#                             if owner.Category.Name == 'HLS-Bauteile':
#                                 return
#                             if owner.Category.Name == 'Rohrzubehör':
#                                 faminame = owner.Symbol.FamilyName
#                                 if faminame.find('Regelventil') != -1:
#                                     self.regelventil = owner.Id.ToString()
#                                 elif faminame.find('6_Wege') != -1:
#                                     self.sechswege = owner.Id.ToString()
#                                     return


#                             self.Wirkungsort_HK(owner)
    
#     def Wirkungsort_H(self, elem):
#         '''Ermittlung der Ventil Wirkungsorte'''
#         if self.regelventil:
#             return
#         elemid = elem.Id.ToString()
#         self.list.append(elemid)
#         cate = elem.Category.Name
#         if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Rohrdämmung', 'Luftkanaldämmung außen',
#                         'Luftkanaldämmung innen']:
#             conns = None
#             try:
#                 conns = elem.MEPModel.ConnectorManager.Connectors
#             except:
#                 try:
#                     conns = elem.ConnectorManager.Connectors
#                 except:
#                     pass

#             if conns:

#                 if conns.Size > 8:
#                     self.regelventil = None
#                     return
#                 for conn in conns:
#                     if elemid == self.elemid.ToString():

#                         try:
#                             if conn.PipeSystemType.ToString() == 'ReturnHydronic':
#                                 continue
#                         except:
#                             continue
#                     refs = conn.AllRefs
#                     for ref in refs:
#                         owner = ref.Owner

#                         if not owner.Id.ToString() in self.list:
#                             if owner.Category.Name == 'HLS-Bauteile':
#                                 return
#                             if owner.Category.Name == 'Rohrzubehör':
#                                 faminame = owner.Symbol.FamilyName
#                                 if faminame.find('Regelventil') != -1:
#                                     self.regelventil = owner.Id.ToString()
#                                     return

#                             self.Wirkungsort_H(owner)

# # wege6 = {}
# # rv = {}
# # for el in hksegel_coll:
# #     hksegel = HKSegel(el.Id)
# #     if hksegel.sechswege:
# #         if hksegel.sechswege not in wege6.keys():
# #             wege6[hksegel.sechswege] = []
# #         wege6[hksegel.sechswege].append(hksegel)
# #     if hksegel.regelventil:
# #         if hksegel.regelventil not in rv.keys():
# #             rv[hksegel.regelventil] = []
# #         rv[hksegel.regelventil].append(hksegel)

# class Ventil:
#     def __init__(self,elemid,liste):
#         self.elemid = elemid
#         self.elem = doc.GetElement(DB.ElementId(int(self.elemid)))
#         try:
#             self.einbauort = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
#         except:
#             self.einbauort = ''
        
#         self.liste = liste
#         self.ort = ''
#         self.ort_liste = []
#         self.hl = 0
#         self.kl = 0
#         self.get_summe()
#         self.mc = None

#     def get_MC(self):
#         if self.einbauort:
#             if self.einbauort in dict_mc.keys():
#                 mcs = dict_mc[self.einbauort]
#                 for el in mcs:
#                     if el.Location.Point.DistanceTo(self.elem.Location.Point) <=200/304.8:
#                         self.mc = el
#                         return
#         for el in MCNode_coll:
#             if el.Location.Point.DistanceTo(self.elem.Location.Point) <=200/304.8:
#                 self.mc = el
#                 return
#         if not self.mc:
#             logger.error('Kein MC Connection für element {}'.format(self.elemid))

#     def werte_schreiben(self):
#         self.elem.LookupParameter('IGF_X_Wirkungsort').Set(self.ort)
#         self.elem.LookupParameter('IGF_X_Einbauort').Set(self.einbauort)
    
#     def werte_schreiben_MC(self):
#         # self.mc.get_Parameter(Guid('49d7278f-3e2b-4f2e-b619-43e856e15be7')).SetValueString(str(round(self.hl)))
#         # self.mc.LookupParameter('IGF_X_Wirkungsort').Set(self.ort)
#         try:
#             nummer = self.elem.LookupParameter('IGF_X_Bauteilnummerierung').AsString()
#             self.mc.LookupParameter('IGF_X_Bauteilnummerierung').Set('DSH'+nummer[2:])
#         except:
#             print(self.elemid)
       
#     def get_summe(self):
#         for el in self.liste:
#             self.hl += el.hl
#             if el.raumnummer and el.raumnummer not in self.ort_liste:
#                 self.ort_liste.append(el.raumnummer)
#             if el.typ == 'HK':
#                 self.kl += el.kl
#         for el in sorted(self.ort_liste):
#             self.ort += el + ', '
#         if len(self.ort) > 2:
#             self.ort = self.ort[:-2]

# t = DB.Transaction(doc,'Leistung')
# t.Start()
# for el in wege6.keys():
#     vt = Ventil(el,wege6[el])
#     vt.get_MC()
#     vt.werte_schreiben()
#     vt.werte_schreiben_MC()

# # for el in rv.keys():
# #     vt = Ventil(el,rv[el])
# #     vt.werte_schreiben()
# t.Commit()
    