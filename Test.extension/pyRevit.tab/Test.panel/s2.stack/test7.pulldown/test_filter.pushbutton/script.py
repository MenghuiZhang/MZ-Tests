# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from rpw import DB,revit
# from eventhandler import ExternalEvent,BESCHIFTUNG2
import os
import clr
from System.Collections.Generic import List
from color_conversion import rgb2hsl,hsl2rgb
from System import Byte


__title__ = "Test"
__doc__ = """

Beschriftung ausrichten
gilt nur in Grundriss.

1. auf X(Y) ausrichten.
2. fixierte Beschriftung auswählen
3. zu verschiebenen Beschriftungen auswählen
4. Fertig stellen klicken.

[2022.08.08]
Version: 1.1

"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass


doc = revit.doc


all_lines = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
LINES = {i.Name: i  for i in all_lines.SubCategories if i.Id.ToString() not in ['-2000066','-2000831','-2000079','-2009018','-2000045','-2009019','-2000077','-2000065']}

systemtyp = DB.FilteredElementCollector(doc).OfClass(DB.MEPSystemType).ToElements()
SYSTEMTYPE = {i.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():i for i in systemtyp}

Filters_coll = DB.FilteredElementCollector(doc).OfClass(DB.FilterElement).ToElements()
FILTERS = {i.Name :i for i in Filters_coll}

systemline = {}
for name in SYSTEMTYPE.keys():
    if name in LINES.keys():
        systemline[name] = LINES[name]
    else:
        systemline[name] = None

Filters = {}
for name in SYSTEMTYPE.keys():
    if name in FILTERS.keys():
        Filters[name] = FILTERS[name]
    else:
        Filters[name] = None

# t = DB.Transaction(doc,'Linestyle erstellen')
# t.Start()
# for name in sorted(systemline.keys()):
#     line = systemline[name]
#     if line:
#         line.SetLineWeight(4,DB.GraphicsStyleType.Projection)
#         line.LineColor = SYSTEMTYPE[name].LineColor
#         if SYSTEMTYPE[name].LinePatternId.IntegerValue != -1:
#             print(SYSTEMTYPE[name].LinePatternId,name )
#             line.SetLinePatternId( SYSTEMTYPE[name].LinePatternId ,DB.GraphicsStyleType.Projection )
#         print(name + ' angepasst')
#     else:
#         line = doc.Settings.Categories.NewSubcategory(all_lines,name)
#         doc.Regenerate()
#         line.SetLineWeight(4,DB.GraphicsStyleType.Projection)
#         line.LineColor = SYSTEMTYPE[name].LineColor
#         if SYSTEMTYPE[name].LinePatternId.IntegerValue != -1:
#             print(SYSTEMTYPE[name].LinePatternId,name )
#             line.SetLinePatternId( SYSTEMTYPE[name].LinePatternId ,DB.GraphicsStyleType.Projection )
#             break
#         systemline[name] = line
#         print(name + ' erstellt')
        

# t.Commit()


# t = DB.Transaction(doc,'ViewFilter erstellen')
# t.Start()
# Cate_Rohr = List[DB.ElementId]()
# Cate_Rohr.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_PipeInsulations).Id)
# Cate_Rohr.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_PipeAccessory).Id)
# Cate_Rohr.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_FlexPipeCurves).Id)
# Cate_Rohr.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_PipeFitting).Id)
# Cate_Rohr.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_PipeCurves).Id)

# Cate_Luft = List[DB.ElementId]()
# Cate_Luft.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DuctInsulations).Id)
# Cate_Luft.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_FlexDuctCurves).Id)
# Cate_Luft.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DuctAccessory).Id)
# Cate_Luft.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DuctFitting).Id)
# Cate_Luft.Add(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_DuctCurves).Id)

# filterRules = List[DB.FilterRule]()
# name = '20_D_Rücklauf 95°C Reindampf-Kondensat'
# # for name in sorted(Filters.keys()):
# _filter = Filters[name]
# # if _filter:
# #     continue
# # else:
# if SYSTEMTYPE[name].GetType().Name == 'PipingSystemType':
#     _filter = DB.ParameterFilterElement.Create(doc, name, Cate_Rohr)
#     paramelemid = DB.ElementId(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
# else:
#     _filter = DB.ParameterFilterElement.Create(doc, name, Cate_Luft)
#     paramelemid = DB.ElementId(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
# doc.Regenerate()
# filterRules.Add(DB.ParameterFilterRuleFactory.CreateEqualsRule(paramelemid, SYSTEMTYPE[name].Id))
# elemFilter = DB.ElementParameterFilter(filterRules)
# _filter.SetElementFilter(elemFilter)
# Filters[name] = _filter
# print(name + ' Filter erstellt')
#     # break

# t.Commit()

t = DB.Transaction(doc,'Filter hinzufügen')
t.Start()

view = revit.active_view
if view.ViewTemplateId.IntegerValue != -1:
    view = doc.GetElement(view.ViewTemplateId)

filter_liste = {doc.GetElement(filterid).Name: doc.GetElement(filterid) for filterid in view.GetFilters()}
name = '20_D_Rücklauf 95°C Reindampf-Kondensat'
if '20_D_Rücklauf 95°C Reindampf-Kondensat' not in filter_liste.keys():
    view.AddFilter(Filters[name].Id)

    overrideGraphicSettings = DB.OverrideGraphicSettings()
    if SYSTEMTYPE[name].LinePatternId:
        overrideGraphicSettings.SetProjectionLinePatternId(SYSTEMTYPE[name].LinePatternId)
    if SYSTEMTYPE[name].LineColor:
        h,s,l = rgb2hsl(SYSTEMTYPE[name].LineColor.Red,SYSTEMTYPE[name].LineColor.Green,SYSTEMTYPE[name].LineColor.Blue)
        r,g,b = hsl2rgb(h,s,l-40)
        overrideGraphicSettings.SetProjectionLineColor(DB.Color(Byte(r),Byte(g),Byte(b)))
        overrideGraphicSettings.SetSurfaceForegroundPatternColor(SYSTEMTYPE[name].LineColor)
    # overrideGraphicSettings.SetSurfaceForegroundPatternId()
    view.SetFilterOverrides(Filters[name].Id,overrideGraphicSettings)
    

t.Commit()


# # class GUI(forms.WPFWindow):
# #     def __init__(self):
# #         self.Beschriftung1 = BESCHIFTUNG2()
# #         self.BeschriftungEvent1 = ExternalEvent.Create(self.Beschriftung1)
# #         forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
# #         self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
         
# #     def manuell(self, sender, args):
# #         self.BeschriftungEvent1.Raise()   

# # gui = GUI()
# # gui.Beschriftung1.GUI = gui
# # gui.Show()