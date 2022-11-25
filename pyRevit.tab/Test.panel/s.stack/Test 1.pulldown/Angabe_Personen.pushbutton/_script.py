# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
from pyrevit.forms import WPFWindow

__title__ = "Angabe MEP"
__doc__ = """
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass
revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()


revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el
rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()

rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()

doc = revit.doc


def Raumnummer():
    param_equality=DB.FilterStringContains()
    param_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    param_prov=DB.ParameterValueProvider(param_id)
    param_value_rule=DB.FilterStringRule(param_prov,param_equality,'stuhl',True)
    param_filter = DB.ElementParameterFilter(param_value_rule)
    return param_filter

stuhl = DB.FilteredElementCollector(rvtdoc).WherePasses(Raumnummer()).WhereElementIsNotElementType().ToElements()#.OfCategory(DB.BuiltInCategory.OST_Furniture)
DICT_Anzahl = {}
for el in stuhl:
    roomid = el.get_Parameter(DB.BuiltInParameter.ELEM_ROOM_ID).AsValueString()
    room = rvtdoc.GetElement(DB.ElementId(int(roomid)))
    if not room:
        continue

    number = room.Number
    name = room.LookupParameter('Name').AsString()

    if number not in DICT_Anzahl.keys():
        DICT_Anzahl[number] = [1,name]
    else:
        DICT_Anzahl[number][0] += 1


class MEPRaum:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.stuhlanzahl = 0
        self.raumname = self.elem.LookupParameter('Name').AsString()
        if self.elem.Number in DICT_Anzahl.keys():
            if self.elem.LookupParameter('Name').AsString() == DICT_Anzahl[self.elem.Number][1]:self.stuhlanzahl = DICT_Anzahl[self.elem.Number][0]
        self.beleuchtung = '10'
        

    @property
    def gekuhlt(self):
        return self.elem.LookupParameter('Bedingungstyp').AsValueString().upper().find('GEKÜHLT') != -1   
    
    def get_Raumtyp(self):
        
            
    
        if self.gekuhlt:
            if self.elem.LookupParameter('Personenzahl').AsValueString() == '0' and self.stuhlanzahl != 0:
                try:
                    self.elem.LookupParameter('Personenzahl').SetValueString(str(self.stuhlanzahl))

                    self.elem.LookupParameter('Eingabe MEP').Set(0)
                except Exception as e:
                    
                    logger.error(e)
                    logger.error(self.elemid,self.raumname)

# spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

# t = DB.Transaction(doc,'MEP')
# t.Start()
# for el in spaces_collector:
#     raumname = el.LookupParameter('Name').AsString().upper()
#     if el.LookupParameter('Bedingungstyp').AsValueString().find('Nicht konditioniert') != -1:el.LookupParameter('Eingabe MEP').Set(1)  
#     # if raumname.find('KÜCHE') != -1 or raumname.find('LAGER') != -1:
#     #     el.LookupParameter('Angegebene Versorgungsleistung').SetValueString(el.LookupParameter('IGF_K_KühllastLaborRaum').AsValueString())
#     #     print(el.Number+' - '+raumname)
#     #     el.LookupParameter('Eingabe MEP').Set(1)
#     #     el.LookupParameter('Personenzahl').SetValueString(str(0))

# t.Commit()



spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()

t = DB.Transaction(doc,'MEP')
t.Start()
for el in spaces_collector:
    MEPRaum(el).get_Raumtyp()
t.Commit()

# UI.TaskDialog.Show(__title__, "Bauteile, die Parameter 'IGF_HLS_Brandschott' den Wert 999 sind, werden nicht berücksichtigt!")
# # Option
# opt = DB.Options()
# opt.ComputeReferences = True
# opt.IncludeNonVisibleObjects = True

# # SolidCurveIntersectionOptions
# opt1 = DB.SolidCurveIntersectionOptions()
# opt1.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsOutside
# opt2 = DB.SolidCurveIntersectionOptions()
# opt2.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside

# def getSolids(elem):
#     lstSolid = []
#     ge = elem.get_Geometry(opt)
#     if ge != None:
#         lstSolid.extend(GetSolid(ge))
#     return lstSolid

# def GetSolid(GeoEle):
#     lstSolid = []
#     for el in GeoEle:
#         if el.GetType().ToString() == 'Autodesk.Revit.DB.Solid':
#             if el.SurfaceArea > 0 and el.Volume > 0 and el.Faces.Size > 1 and el.Edges.Size > 1:
#                 lstSolid.append(el)
#         elif el.GetType().ToString() == 'Autodesk.Revit.DB.GeometryInstance':
#             ge = el.GetInstanceGeometry()
#             lstSolid.extend(GetSolid(ge))
#     return lstSolid

# def TransformSolid(elem):
#     m_lstModels = []
#     listSolids = getSolids(elem)
#     for solid in listSolids:
#         tempSolid = solid
#         tempSolid = DB.SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
#         m_lstModels.append(tempSolid)
#     return m_lstModels

# revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
# revitLinks = revitLinks_collector.ToElementIds()

# system_rohr = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()
# system_rohr_dict = {}

# def coll2dict(coll,dict):
#     for el in coll:
#         type = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
#         if type in dict.Keys:
#             dict[type].append(el.Id)
#         else:
#             dict[type] = [el.Id]

# coll2dict(system_rohr,system_rohr_dict)
# system_rohr.Dispose()

# revitLinksDict = {}
# for el in revitLinks_collector:
#     revitLinksDict[el.Name] = el
# revitLinks_collector.Dispose()
# rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
# rvtdoc = None
# if not rvtLink:
#     logger.error("Keine Revitverknüpfung gewählt")
#     script.exit()

# rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
# if not rvtdoc:
#     logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
#     script.exit()

# walls = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
# wallsName = []
# for el in walls:
#     if not el.Name in wallsName:
#         wallsName.append(el.Name)
# BrandWalls = forms.SelectFromList.show(wallsName,multiselect=True, button_name='Select Walls')

# BrandWallEles = []
# for el in walls:
#     if el.Name in BrandWalls:
#         BrandWallEles.append(el)

# class System(object):
#     def __init__(self):
#         self.checked = False
#         self.SystemName = ''
#         self.TypName = ''

#     @property
#     def TypName(self):
#         return self._TypName
#     @TypName.setter
#     def TypName(self, value):
#         self._TypName = value
#     @property
#     def checked(self):
#         return self._checked
#     @checked.setter
#     def checked(self, value):
#         self._checked = value
#     @property
#     def ElementId(self):
#         return self._ElementId
#     @ElementId.setter
#     def ElementId(self, value):
#         self._ElementId = value

# Liste_Rohr = ObservableCollection[System]()

# for key in system_rohr_dict.Keys:
#     temp_system = System()
#     temp_system.TypName = key
#     temp_system.ElementId = system_rohr_dict[key]
#     Liste_Rohr.Add(temp_system)

# # GUI Systemauswahl
# class Systemauswahl(WPFWindow):
#     def __init__(self, xaml_file_name,liste_Rohr):
#         self.liste_Rohr = liste_Rohr
#         WPFWindow.__init__(self, xaml_file_name)
#         self.tempcoll = ObservableCollection[System]()
#         self.altdatagrid = None

#         try:
#             self.dataGrid.ItemsSource = liste_Rohr
#             self.altdatagrid = liste_Rohr
#         except Exception as e:
#             logger.error(e)

#         self.suche.TextChanged += self.search_txt_changed

#     def search_txt_changed(self, sender, args):
#         """Handle text change in search box."""
#         self.tempcoll.Clear()
#         text_typ = self.suche.Text
#         if text_typ in ['',None]:
#             self.dataGrid.ItemsSource = self.altdatagrid

#         else:
#             if text_typ == None:
#                 text_typ = ''
#             for item in self.altdatagrid:
#                 if item.TypName.find(text_typ) != -1:
#                     self.tempcoll.Add(item)
#             self.dataGrid.ItemsSource = self.tempcoll
#         self.dataGrid.Items.Refresh()

#     def checkall(self,sender,args):
#         for item in self.dataGrid.Items:
#             item.checked = True
#         self.dataGrid.Items.Refresh()

#     def uncheckall(self,sender,args):
#         for item in self.dataGrid.Items:
#             item.checked = False
#         self.dataGrid.Items.Refresh()

#     def toggleall(self,sender,args):
#         for item in self.dataGrid.Items:
#             value = item.checked
#             item.checked = not value
#         self.dataGrid.Items.Refresh()

#     def auswahl(self,sender,args):
#         self.Close()

# Systemwindows = Systemauswahl("System.xaml",Liste_Rohr)
# try:
#     Systemwindows.ShowDialog()
# except Exception as e:
#     logger.error(e)
#     Systemwindows.Close()
#     script.exit()


# class System_Liste:
#     def __init__(self,systemtyp,system):
#         self.systemtyp = systemtyp
#         self.system = system
#         self.bauteile = []

#     def get_bauteile(self):
#         for el in self.system.PipingNetwork:
#             cate = el.Category.Name
#             if cate in ['Rohre']:
#                 if el.Id.ToString() not in self.bauteile:
#                     self.bauteile.append(el.Id.ToString())


# liste_system = []
# for el in Liste_Rohr:
#     if el.checked == True:
#         temp_system = System_Liste(el.TypName,None)
#         for elemid in el.ElementId:
#             elem = doc.GetElement(elemid)
#             temp_system.system = elem
#             temp_system.get_bauteile()
#         liste_system.append(temp_system)
# if len(liste_system) == 0:
#     logger.info('Keine System ausgewählt')
#     script.exit()

# class Bauteil:
#     def __init__(self,elemid,liste_solid):
#         self.elemid = elemid
#         self.elem = doc.GetElement(DB.ElementId(int(elemid)))
#         self.anzahl_vorab = self.elem.LookupParameter('IGF_HLS_Brandschott').AsValueString()
#         self.line_liste = []
#         self.liste_solid = liste_solid
#         self.anzahl = 0
#         self.get_line_liste()
#         self.get_anzahl()
        
    
#     def get_line_liste(self):
#         conns = self.elem.ConnectorManager.Connectors
#         liste_punkte = []
#         for conn in conns:
#             liste_punkte.append(conn.Origin)
#         for n in range(len(liste_punkte)):
#             for j in range(n+1,len(liste_punkte)):
#                 try:
#                     self.line_liste.append(DB.Line.CreateBound(liste_punkte[n], liste_punkte[j]))
#                 except:
#                     pass

#     def get_anzahl(self):
#         if self.anzahl_vorab == '999':
#             self.anzahl = 999
#             return

#         line =  self.line_liste[0]
#         for item in self.liste_solid:
#             elelink = item[1]
#             for solid in elelink:
#                 result1 = solid.IntersectWithCurve(line,opt1)
#                 result2 = solid.IntersectWithCurve(line,opt2)
#                 if result1.SegmentCount > 0 and result2.SegmentCount > 0:
#                     self.anzahl += 1
#                     result1.Dispose()
#                     result2.Dispose() 
#                     break
#                 result1.Dispose()
#                 result2.Dispose() 

#     def line_entfernen(self):
#         for line in self.line_liste:
#             line.Dispose()
    
#     def wert_schreiben(self):
#         try:
#             self.elem.LookupParameter('IGF_HLS_Brandschott').Set(self.anzahl)
#         except:
#             pass


# # RvtLinkElem
# RvtLinkElemSolids = []
# # ProElemCurve = []
# step = int(len(BrandWallEles)/200)
# with forms.ProgressBar(title='{value}/{max_value} Wände in Revitverknüpfung',cancellable=True, step=step) as pb:
#     n_1 = 0
#     for ele in BrandWallEles:
#         if pb.cancelled:
#             script.exit()
#         n_1 += 1
#         pb.update_progress(n_1, len(BrandWallEles))
#         models = TransformSolid(ele)
#         RvtLinkElemSolids.append([ele,models])

# rvtdoc.Dispose()

# liste_bauteile = []
# dict_bauteile = {}
# Nichtbearbeitet = [e.systemtyp for e in liste_system]
# Bearbeitet = []     
# with forms.ProgressBar(title='',cancellable=True, step=30) as pb2:
#     for n0,system in enumerate(liste_system):
#         pb2.title = 'Brandschott zählen --- {value}/{max_value} Rohre in ' + str(n0+1) + '/' + str(len(liste_system)) + ' Systeme ---- ' + system.systemtyp
#         pb2.step = int(len(system.bauteile)/1000)+10
#         if pb2.cancelled:
#             break
#         temp_liste = []
#         for n1,bauteilid in enumerate(system.bauteile):
#             if pb2.cancelled:
#                 break
#             bauteil = Bauteil(bauteilid,RvtLinkElemSolids)
#             bauteil.line_entfernen()
#             liste_bauteile.append(bauteil)
#             temp_liste.append(bauteil)
#             pb2.update_progress(n1+1,len(system.bauteile))
#         dict_bauteile[system.systemtyp] = temp_liste


# t = DB.Transaction(doc,'Brandschott zählen')
# t.Start()
# if forms.alert('Anzahl schreiben?',yes=True,no=True,ok=False):
#     with forms.ProgressBar(title='{value}/{max_value} Luftkanäle und Luftkanalformteile',cancellable=True, step=int(len(liste_bauteile)/1000)+10) as pb3:
#         n = 0
#         for systemtyp in dict_bauteile.keys():
#             for elem in dict_bauteile[systemtyp]:
#                 n+=1
#                 if pb3.cancelled:
#                     if forms.alert('bisherige Änderung behalten?',yes=True,no=True,ok=False):
#                         t.Commit()
#                         logger.info('Folgenede Systeme sind bereits bearbeitet.')
#                         for el in Bearbeitet:
#                             logger.info(el)
#                         logger.info('---------------------------------------')
#                         logger.info('Folgenede Systeme sind nicht bearbeitet.')
#                         for el in Nichtbearbeitet:
#                             logger.info(el)
#                     else:
#                         t.RollBack()
                    
#                     script.exit()
#                 elem.wert_schreiben()
#                 pb3.update_progress(n, len(liste_bauteile))
#             Bearbeitet.append(systemtyp)
#             Nichtbearbeitet.remove(systemtyp)
            
# t.Commit()

# for item in RvtLinkElemSolids:
#     for solid in item[1]:
#         solid.Dispose()