# coding: utf8
from pyrevit import script, forms,revit
from Autodesk.Revit.DB import *

__title__ = "ermittelt Deckentyp für Revi-Öffnungen"
__doc__ = """Revi-Öffnungen(Deckenspiegel)
Parameter: IGF_BIG_Basisbauteil Revi-öffnungen
Kategorie: Allgemeines Modell"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config = script.get_config()

uidoc = revit.uidoc
doc = revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)

# Wände
revitLinks_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

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
Cellings = FilteredElementCollector(rvtdoc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType()
CellingsName = []
for el in Cellings:
    if not el.Name in CellingsName:
        CellingsName.append(el.Name)
Cellings_select = forms.SelectFromList.show(CellingsName,
multiselect=True, button_name='Select Decken')

cellingsListe = []
for el in Cellings:
    if el.Name in Cellings_select:
        cellingsListe.append(el)

# Allgemeines Modell

Bauteile = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType()
Bauteile_Ids = Bauteile.ToElementIds()
Bauteile.Dispose()

# revi-Öffnungen
Liste_Bauteile = []
for el in Bauteile_Ids:
    elem = doc.GetElement(el)
    Family = elem.Symbol.FamilyName
    if Family == 'Revi-Öffnungen':
        Liste_Bauteile.append(elem)


def getSolids(elem):
    lstSolid = []
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    ge = elem.get_Geometry(opt)
    if ge != None:
        lstSolid.extend(GetSolid(ge))
    return lstSolid

def GetSolid(GeoEle):
    lstSolid = []
    for el in GeoEle:
        if el.GetType().ToString() == 'Autodesk.Revit.DB.Solid':
            if el.SurfaceArea > 0 and el.Volume > 0 and el.Faces.Size > 1 and el.Edges.Size > 1:
                lstSolid.append(el)
        elif el.GetType().ToString() == 'Autodesk.Revit.DB.GeometryInstance':
            ge = el.GetInstanceGeometry()
            lstSolid.extend(GetSolid(ge))
    return lstSolid

def TransformSolid(elem):
    m_lstModels = []
    listSolids = getSolids(elem)
    for solid in listSolids:
        tempSolid = solid
        tempSolid = SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
        m_lstModels.append(tempSolid)
    return m_lstModels

def HLSkurve(elem,erweite):
    hlscurve = None
    BB = elem.get_BoundingBox(None)
    list = []
    Cen_X = (BB.Max.X + BB.Min.X) / 2
    Cen_Y = (BB.Max.Y + BB.Min.Y) / 2
    list = [XYZ(Cen_X,Cen_Y,BB.Max.Z + erweite),XYZ(Cen_X,Cen_Y,BB.Min.Z - erweite)]
    hlscurve = Line.CreateBound(list[0], list[1])
    return hlscurve
    

def EbenenUmbenennen(ebene):
    out = ''
    if not ebene:
        out = ''
    if ebene.find('EG') != -1:
        out = 'EG'
    elif ebene.find('SAN') != -1:
        out = 'EG'
    elif ebene.find('1.OG') != -1:
        out = '1.OG'
    elif ebene.find('2.OG') != -1:
        out = '2.OG'
    elif ebene.find('3.OG') != -1:
        out = '3.OG'
    elif ebene.find('1.UG') != -1:
        out = '1.UG'
    elif ebene.find('2.UG') != -1:
        out = '2.UG'
    elif ebene.find('3.UG') != -1:
        out = '3.UG'
    else:
        out = '4.OG'
    return out


# RvtLinkElem
RvtLinkElemSolids = {}
# ProElemCurve = []
step = int(len(cellingsListe)/200)
with forms.ProgressBar(title='{value}/{max_value} Decken in RVT-Link Model',cancellable=True, step=step) as pb:
    for n_1, ele in enumerate(cellingsListe):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n_1, len(cellingsListe))
        models = TransformSolid(ele)
        ebenename = rvtdoc.GetElement(ele.LevelId).Name
        deckenName = ele.Name
        Ebenen = EbenenUmbenennen(ebenename)
        if not Ebenen in RvtLinkElemSolids.keys():
            RvtLinkElemSolids[Ebenen] = {}
        if not deckenName in RvtLinkElemSolids[Ebenen].keys():
            RvtLinkElemSolids[Ebenen][deckenName] = []
        if not models in RvtLinkElemSolids[Ebenen][deckenName]:
            RvtLinkElemSolids[Ebenen][deckenName].append(models)

rvtdoc.Dispose()
Datenbank = {}


title = '{value}/{max_value} Bauteile'
with forms.ProgressBar(title=title,cancellable=True, step=10) as pb2:
    for n_1,elem in enumerate(Liste_Bauteile):
        if pb2.cancelled:
            script.exit()
        pb2.update_progress(n_1+1, len(Liste_Bauteile))
        ebene = elem.LookupParameter('Bauteillistenebene').AsValueString()
        neu_ebene = EbenenUmbenennen(ebene)
        curve = HLSkurve(elem,0.2)
        if not curve:
            logger.error('Linie kann nicht erstellt werden. ID: ' + elem.Id.ToString())
            continue
        Klass = ''
        if neu_ebene in RvtLinkElemSolids.keys():
            models_dict = RvtLinkElemSolids[neu_ebene]
            
            for lin in [0.2,0.5,1,1.5]:
                curve = HLSkurve(elem,lin)
                if not elem.Id.ToString() in Datenbank.keys():
                    for klasse in models_dict.keys():
                        models = models_dict[klasse]
                        opt = SolidCurveIntersectionOptions()
                        for item in models:
                            if not Klass:
                                for solid in item:
                                    result = solid.IntersectWithCurve(curve,opt)
                                    if result.SegmentCount > 0:
                                        Klass = klasse
                                        Datenbank[elem.Id.ToString()] = Klass
            
        if not elem.Id.ToString() in Datenbank.keys():
            Datenbank[elem.Id.ToString()] = ''

# Daten schreiben
if forms.alert("Daten schreiben?", ok=False, yes=True, no=True):
    t = Transaction(doc,'Revi-Öffnungen')
    t.Start()
    eles = Datenbank.keys()[:]
    step3 = int(len(eles)/1000)+10
    with forms.ProgressBar(title='{value}/{max_value} Bauteile',cancellable=True, step=step3) as pb3:
        for n,el in enumerate(eles):
            if pb3.cancelled:
                t.RollBack()
                script.exit()
            pb3.update_progress(n, len(eles))
            doc.GetElement(ElementId(int(el))).LookupParameter('IGF_BIG_Basisbauteil Revi-öffnungen').Set(Datenbank[el])
    t.Commit()