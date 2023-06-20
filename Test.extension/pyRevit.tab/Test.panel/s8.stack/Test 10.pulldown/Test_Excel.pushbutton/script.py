# coding: utf8
# import os
# from pyrevit import revit, UI, DB
# from pyrevit import script
# from IGF_Elementfilter import ElementFamilyTypeSeparatelyContains_Filter,ElementFamilyTypeSeparatelyEquals_Filter
import time
from rpw import revit,DB
__title__ = "Test_Excel"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"


# from IGF_log import getlog
# from excel._NPOI_2020 import getlog1
# t0 = time.time()
# try:
#     getlog(__title__)
# except:
#     pass
# t1 = time.time()

# try:
#     getlog1(__title__)
# except:
#     pass
# t2 = time.time()
# print(t1-t0,t2-t1,(t1-t0)/(t2-t1))

doc = revit.doc
uidoc = revit.uidoc

cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
connset = DB.ConnectorSet()
for el in cl:
    conns = el.MEPModel.ConnectorManager.Connectors
    for conn in conns:
        if conn.Direction.ToString() == 'Out' and conn.MEPSystem == None:
            connset.Insert(conn)
t = DB.Transaction(doc,'System')
t.Start()
doc.GetElement(DB.ElementId(14613235)).Add(connset)
t.Commit()

t = DB.Transaction(doc,'System')
t.Start()
connset = DB.ConnectorSet()
rohr = None
zube = None
for el in cl:
    if el.Category.Name == 'Rohre':
        rohr = el
    else:
        zube = el
conns = zube.MEPModel.ConnectorManager.Connectors
conn0 = None

for conn in conns:
    if conn.MEPSystem:
        conn0 = conn
    else:
        connset.Insert(conn)
conn0.MEPSystem.Add(rohr.ConnectorManager.Connectors)
t.Commit()

t = DB.Transaction(doc,'System')
t.Start()
cl = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
for el in cl:
    bearbeitungsbereich = el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
    doc.GetWorksetTable().SetActiveWorksetId(DB.WorksetId(bearbeitungsbereich))
    try:
        conns = el.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                refs = conn.AllRefs
                for ref in refs:
                    if ref.Owner.Category.Name == 'Rohrzubehör':
                        conn.DisconnectFrom(ref)
                        ref.Owner.Location.Move(conn.CoordinateSystem.BasisZ * 10 / 304.8)
                        try:DB.Plumbing.Pipe.Create(doc,DB.ElementId(11593842),el.LevelId,conn,ref)
                        except:print(el.Id)
                        break
    except:print(el.Id)

t.Commit()  





# logger = script.get_logger()

# DICT_MEP_NUMMER_NAME  = {}

# #hksegel_coll = ElementFamilyTypeSeparatelyEquals_Filter('_H_IGF_422_MC_Kermi Therm X2 Profil-K Type22', 'Kermi Therm X2 Profil-K Type22 1100x300_15',False, DB.BuiltInCategory.OST_MechanicalEquipment)
# def get_MEP_NUMMER_NAMR():
#     spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
#     for el in spaces:
#         DICT_MEP_NUMMER_NAME[el.Number] = el


# class HKSegel:
#     def __init__(self, elemid):
#         self.elemid = elemid
#         self.elem = doc.GetElement(self.elemid)
#         self.name = self.elem.Symbol.FamilyName + ': ' + self.elem.Name
#         self.raumname = ''
#         self.phase = doc.GetElement(self.elem.CreatedPhaseId)
#         self.raum = None
#         self.raumtyp = ''
#         self.GetRaum()
#         if self.raum:
#             self.raumtyp = self.raum.LookupParameter('Bedingungstyp').AsValueString()
#         self.info = ''
        
#         # if self.raumtyp == 'Beheizt und gekühlt':
#         #     if self.elem.Name.find('EHV') == -1:
#         #         print(self.elem.Id.ToString(),'Mit EHV')
#         #         self.info = 'mit'
#         # else:
#         #     if self.elem.Name.find('EHV') != -1:
#         #         print(self.elem.Id.ToString(),'Ohne EHV')
#         #         self.info = 'ohne'



#     def Muster_Pruefen(self):
#         '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
#         try:
#             bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
#             if bb == 'KG4xx_Musterbereich':return True
#             else:return False
#         except:return False 

    
#     def changetyp(self):
#         self.elem.ChangeTypeId(DB.ElementId(33921033))
#         #33921013,33921015,33921033,33921035
#         self.elem.LookupParameter('Supply_connector_position').Set(1)
#         self.elem.LookupParameter('Return_connector_position').Set(2)
#         # ids = self.elem.Symbol.Family.GetFamilySymbolIds()
#         # dict_ = {}
#         # for Id in ids:
#         #     symbol = doc.GetElement(Id)
#         #     name = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
#         #     dict_[name] = Id
#         # if self.info == 'mit':
#         #     name_neu = self.elem.Name+'_EHV'
#         #     if name_neu in dict_.keys():
#         #         self.elem.ChangeTypeId(dict_[name_neu])
#         #     else:
#         #         typ = self.elem.Symbol.Duplicate(name_neu)
#         #         doc.Regenerate()
                
    
#     def GetRaum(self):
#         try:
#             self.raum = self.elem.Space[self.phase]
#             self.raum.Number
            
#         except:
#             if not self.Muster_Pruefen():
#                 try:
#                     logger.error('Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elemid.ToString()))
#                     # param_einbauort = self.get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
#                     # if param_einbauort not in DICT_MEP_NUMMER_NAME.keys():
#                     #     logger.error('Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elemid.ToString()))
#                     # else:
#                     #     self.raum = DICT_MEP_NUMMER_NAME[param_einbauort]
#                 except:pass
#             return


# Liste = [] 
# # '_H_IGF_422_Charleston-4-1200-3000','_H_IGF_422_MC_Charleston-2-1200-3000',
# # familien = [
# # '_H_IGF_422_MC_Kermi Profil-K Type11','_H_IGF_422_MC_Kermi Therm X2 Profil-K Type22',
# # '_H_IGF_422_MC_Kermi Therm X2 Profil-K Type33','_H_IGF_422_MC_Kermi Therm X2 Profil-V Type12',
# # 'Kermi Plan-V Hygiene Type10','Kermi Therm X2 Plan-V Hygiene Type20','Kermi Therm X2 Plan-V Hygiene Type30']
# # for fam in familien:
# #     coll = ElementFamilyTypeSeparatelyContains_Filter(fam, '',False, DB.BuiltInCategory.OST_MechanicalEquipment)

# #     for el in coll:
# #         hksegel = HKSegel(el.Id)
# #         Liste.append(hksegel)


# t = DB.Transaction(doc,'Heizkörper wechseln')
# t.Start()
# for el in uidoc.Selection.GetElementIds():
#     # if el.raumtyp == 'Beheizt und gekühlt':
#     try:
#         HKSegel(el).changetyp()
#     except Exception as e:
#         print(e)

# t.Commit()
    