# coding: utf8
from pyrevit import script, forms,revit
from Autodesk.Revit.DB import Transaction,FilteredElementCollector,BuiltInCategory,BuiltInParameter
from pyIGF_logInfo import getlog

__title__ = "1.41 Trassenzugehörigkeit_MedForce"
__doc__ = """IGF_Trassenzugehörigkeit für Projekt MedForce
Kategorien: HLS-Bauteile, Rohre, Rohrzubehör, Luftkanäle, Luftkanalzubehör,
Luftdurchlässe, Kabeltrassen, Leerrohr
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

HLS_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
HLS = HLS_collector.ToElementIds()
HLS_collector.Dispose()

Rohre_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()
Rohre = Rohre_collector.ToElementIds()
Rohre_collector.Dispose()

Luft_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType()
Luft = Luft_collector.ToElementIds()
Luft_collector.Dispose()

Terminal_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
Terminal = Terminal_collector.ToElementIds()
Terminal_collector.Dispose()

Kabel_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType()
Kabel = Kabel_collector.ToElementIds()
Kabel_collector.Dispose()

Leer_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType()
Leer = Leer_collector.ToElementIds()
Leer_collector.Dispose()

Kanalzu_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType()
Kanalzu = Kanalzu_collector.ToElementIds()
Kanalzu_collector.Dispose()

Rohrzu_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType()
Rohrzu = Rohrzu_collector.ToElementIds()
Rohrzu_collector.Dispose()

Kategorien = ['HLS-Bauteile', 'Rohre', 'Rohrzubehör', 'Luftkanäle', 'Luftkanalzubehör',
'Luftdurchlässe', 'Kabeltrassen', 'Leerrohr']
Kategorien.sort()
Cates = forms.SelectFromList.show(Kategorien,multiselect=True, button_name='Select Category')

def Ebenen(ebenename):
    out = ''
    if ebenename.find('EG') != -1:
        out = 'EG'
    elif ebenename.find('NN') != -1:
        out = 'EG'
    elif ebenename.find('SAN') != -1:
        out = 'EG'
    elif ebenename.find('1.UG') != -1:
        out = '1.UG'
    elif ebenename.find('2.UG') != -1:
        out = '2.UG'
    elif ebenename.find('1.OG') != -1:
        out = '1.OG'
    elif ebenename.find('2.OG') != -1:
        out = '2.OG'
    elif ebenename.find('3.OG') != -1:
        out = '3.OG'
    else:
        out = '4.OG'
    return out

def DatenSchreiben(Category,ids):
    title = '{value}/{max_value} Bauteile in Kategorie ' + Category
    step = int(len(ids)/200)+1
    with forms.ProgressBar(title=title, cancellable=True, step=step) as pb:
        for n,id in enumerate(ids):
            if pb.cancelled:
                t.RollBack()
                script.exit()
            pb.update_progress(n, len(ids))
            elem = doc.GetElement(id)
            Ebene = ''
            try:
                Ebene = elem.LookupParameter('CAx_Trassenbezugsebene').AsString()
            except:
                try:
                    Ebene = doc.GetElement(elem.LevelId).Name
                except:
                    try:
                        Ebene = elem.LookupParameter('Bauteillistenebene').AsValueString()
                    except:
                        try:
                            FamilyAndType = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
                        except:
                            FamilyAndType = ''
                        logger.error('Keine Ebene. ElementId: {}, Kategorie: {}, Familie: {}'.format(id,Category,FamilyAndType))
            if not Ebene:
                continue
            new_ebene = Ebenen(Ebene)
            param = elem.LookupParameter('IGF_Trassenzugehörigkeit')
            if param:
                param.Set(new_ebene)
            else:
                logger.error('Kein Parameter IGF_Trassenzugehörigkeit in Kategorie {}'.format(Category))

t = Transaction(doc,'Trassenzugehörigkeit')
t.Start()
if 'HLS-Bauteile' in Cates:
    DatenSchreiben('HLS_Bauteile',HLS)
if 'Rohre' in Cates:
    DatenSchreiben('Rohre',Rohre)
if 'Luftkanäle' in Cates:
    DatenSchreiben('Luftkanäle',Luft)
if 'Kabeltrassen' in Cates:
    DatenSchreiben('Kabeltrassen',Kabel)
if 'Leerrohr' in Cates:
    DatenSchreiben('Leerrohr',Leer)
if 'Rohrzubehör' in Cates:
    DatenSchreiben('Rohrzubehör',Rohrzu)
if 'Luftkanalzubehör' in Cates:
    DatenSchreiben('Luftkanalzubehör',Kanalzu)
if 'Luftdurchlässe' in Cates:
    DatenSchreiben('Luftdurchlässe',Terminal)
t.Commit()