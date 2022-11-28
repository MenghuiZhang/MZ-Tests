# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
import time

__title__ = "8.30 zählt nötige Brandschotts (nur für alle Rohre)"
__doc__ = """zählt nötige Brandschotts für Rohre
Paremeter: IGF_HLS_Brandschutz,IGF_HLS_Brandschott,IGF_HLS_Brandschott_prüfen

[2022.02.15]
Version: 1.3
"""
start = time.time()
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass

rohrs = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElementIds()


# rohrs = uidoc.Selection.GetElementIds()

# Option
opt = DB.Options()
opt.ComputeReferences = True
opt.IncludeNonVisibleObjects = True

# SolidCurveIntersectionOptions
opt1 = DB.SolidCurveIntersectionOptions()
opt1.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsOutside
opt2 = DB.SolidCurveIntersectionOptions()
opt2.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside


class Rohr(object):
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.Line = None
        self.Brandklass = {}
        self.Brandklass_Text = '' 
        self.Anzahl = 0
        self.Pruefen = 'Ja'
        self.richtung = self.get_Typ()
        
    def get_Line(self):
        conns = list(self.elem.ConnectorManager.Connectors)
        try:
            return DB.Line.CreateBound(conns[0].Origin, conns[1].Origin)
        except:
            return None
    
    def get_Brandklass_Text(self):

        try:
            anzahl = 0
            for cate in self.Brandklass:
                for typ in self.Brandklass[cate]:
                    for klass in self.Brandklass[cate][typ]:
                        self.Brandklass_Text += str(self.Brandklass[cate][typ][klass]) + 'x' + cate+'-'+typ+'-'+klass +', '
                        anzahl += self.Brandklass[cate][typ][klass]
            self.Brandklass_Text = self.Brandklass_Text[:-2] 
            if anzahl != self.Anzahl:
                self.Pruefen = 'Warnung: Anzahl unterschiedlich'
        except:
            logger.error(self.elemid)
            logger.error(self.Brandklass)
    
    def get_Typ(self):
        conns = list(self.elem.ConnectorManager.Connectors)
        z = abs(conns[0].Origin.Z- conns[1].Origin.Z)
        xy = ((conns[0].Origin.X- conns[1].Origin.X)**2 + (conns[0].Origin.Y- conns[1].Origin.Y)**2)**0.5
        if z < xy:
            return 'HORIZONTAL'
        else:
            return 'VERTICAL'
    
    def Ignorierne(self):
        try:
            if self.elem.LookupParameter('IGF_HLS_Brandschott_prüfen').AsString().upper().find('Nein') != -1:
                return True
            else:
                return False
        except:
            return False

    
    def wert_schreiben(self):
        if self.Brandklass_Text == None:
            self.Brandklass_Text = ""
        try:
            self.elem.LookupParameter('IGF_HLS_Brandschutz').Set(self.Brandklass_Text)
        except Exception as e:
            logger.error(e)
        
        try:
            self.elem.LookupParameter('IGF_HLS_Brandschott').Set(self.Anzahl)
        except Exception as e:
            logger.error(e)
        
        try:
            self.elem.LookupParameter('IGF_HLS_Brandschott_prüfen').Set(self.Pruefen)
        except Exception as e:
            logger.error(e)
    
    
class LinkedElement(object):
    def __init__(self,elemid,doc):
        self.elemid = elemid
        self.doc = doc
        self.elem = self.doc.GetElement(self.elemid)
        self.Solids = []
        self.TypName = self.elem.Name
        self.Typ = self.get_Type()
        self.Kategorie = self.elem.Category.Name
        self.Brandschutz = self.get_Brandschutz()
        if self.Kategorie == 'Geschossdecken' and self.Typ == '???':
            self.Brandschutz = None
    def get_Type(self):
        if self.TypName.upper().find('TREPPE') != -1:
            return 'TREPPE'
        if self.TypName.upper().find('BETON') != -1:
            return 'Beton'
        elif self.TypName.upper().find('GIPS') != -1 or self.TypName.upper().find('TROCKENBAU') != -1 or self.TypName.upper().find('METALL') != -1 or self.TypName.upper().find('RASTER') != -1 or self.TypName.upper().find('TK') != -1:
            return 'Trockenbau'
        elif self.TypName.upper().find('MAUERWERK') != -1 or self.TypName.upper().find('ZIEGEL') != -1:
            return 'Mauerwerk'
        
        else:
            Materiels = self.elem.GetMaterialIds(False)
            for m in Materiels:
                name = self.doc.GetElement(m).Name
                if name.upper().find('BETON') != -1:
                    return 'Beton'
                elif name.upper().find('GIPS') != -1 or self.TypName.upper().find('TROCKENBAU') != -1 or self.TypName.upper().find('METALL') != -1:
                    return 'Trockenbau'
                elif name.upper().find('MAUERWERK') != -1 or self.TypName.upper().find('ZIEGEL') != -1:
                    return 'Mauerwerk'
            
            return '???'

    def get_Brandschutz(self):
        if self.Kategorie == 'Geschossdecken':
            if self.TypName.upper().find('ENSCAPE') != -1:
                return None
            try:
                if self.elem.LookupParameter('Dicke').AsDouble()*0.3048 <= 0.061:
                    return None
            except:
                return None

            if self.TypName.upper().find('F120') != -1:
                return 'F120'
            elif self.TypName.upper().find('F90') != -1:
                return 'F90'
            elif self.TypName.upper().find('F60') != -1:
                return 'F60'
            elif self.TypName.upper().find('F30') != -1:
                return 'F30'
            elif self.TypName.upper().find('BW') != -1:
                return 'BW'
            else:
                return 'F90'
        else:
            if self.TypName.upper().find('F120') != -1:
                return 'F120'
            elif self.TypName.upper().find('F90') != -1:
                return 'F90'
            elif self.TypName.upper().find('F60') != -1:
                return 'F60'
            elif self.TypName.upper().find('F30') != -1:
                return 'F30'
            elif self.TypName.upper().find('BW') != -1:
                return 'BW'
            
            else:
                return None
    def Brandschutz_Pruefen(self):
        if self.Brandschutz:
            return True
        else:
            return False


revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el
revitLinks_collector.Dispose()
rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()

rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
def getSolids(elem):
    lstSolid = []
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
        tempSolid = DB.SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
        m_lstModels.append(tempSolid)
    return m_lstModels

if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()

walls = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElementIds()
decken = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElementIds()
geschossendecken = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElementIds()

Wall_liste = []
Decken_liste = []

# ProElemCurve = []
with forms.ProgressBar(title='{value}/{max_value} Wände in Revitverknüpfungsmodell',cancellable=True, step=int(len(walls)/200)+1) as pb:
    for n,elemid in enumerate(walls):
        if pb.cancelled:
            script.exit()
        
        pb.update_progress(n+1,len(walls))

        wall = LinkedElement(elemid,rvtdoc)
        if wall.Brandschutz_Pruefen():
            wall.Solids = TransformSolid(wall.elem)
            Wall_liste.append(wall)

with forms.ProgressBar(title='{value}/{max_value} Decken in Revitverknüpfungsmodell',cancellable=True, step=int(len(walls)/200)+1) as pb:
    for n,elemid in enumerate(decken):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n+1,len(decken))

        decke = LinkedElement(elemid,rvtdoc)
        if decke.Brandschutz_Pruefen():
            decke.Solids = TransformSolid(decke.elem)
            Decken_liste.append(decke)

with forms.ProgressBar(title='{value}/{max_value} Geschossendecken in Revitverknüpfungsmodell',cancellable=True, step=int(len(geschossendecken)/200)+1) as pb:
    for n,elemid in enumerate(geschossendecken):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n+1,len(geschossendecken))

        geschossendecke = LinkedElement(elemid,rvtdoc)
        if geschossendecke.Brandschutz_Pruefen():
            geschossendecke.Solids = TransformSolid(geschossendecke.elem)
            Decken_liste.append(geschossendecke)


gesamt_decke = DB.BooleanOperationsUtils.ExecuteBooleanOperation(Decken_liste[0].Solids[0],Decken_liste[0].Solids[0],DB.BooleanOperationsType.Union)
for n in range(1,len(Decken_liste[0].Solids)):
    temp = gesamt_decke
    try:
        gesamt_decke = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt_decke,Decken_liste[0].Solids[n],DB.BooleanOperationsType.Union)
        if temp != Decken_liste[0].Solids[0]:
            temp.Dispose()
    except Exception as e:
        logger.error(e)

for n in range(1,len(Decken_liste)):
    temp = gesamt_decke
    for el in Decken_liste[n].Solids:
        try:
            gesamt_decke = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt_decke,el,DB.BooleanOperationsType.Union)
            temp.Dispose()
        except Exception as e:
            logger.error(e)

gesamt_Wand = DB.BooleanOperationsUtils.ExecuteBooleanOperation(Wall_liste[0].Solids[0],Wall_liste[0].Solids[0],DB.BooleanOperationsType.Union)
for n in range(1,len(Wall_liste[0].Solids)):
    temp = gesamt_Wand
    try:
        gesamt_Wand = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt_Wand,Wall_liste[0].Solids[n],DB.BooleanOperationsType.Union)
        if temp != Wall_liste[0].Solids[0]:
            temp.Dispose()
    except Exception as e:
        logger.error(e)

for n in range(1,len(Wall_liste)):
    temp = gesamt_Wand
    for el in Wall_liste[n].Solids:
        try:
            gesamt_Wand = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt_Wand,el,DB.BooleanOperationsType.Union)
            temp.Dispose()
        except Exception as e:
            logger.error(e)

solid2_liste = getSolids(doc.GetElement(uidoc.Selection.GetElementIds()[0]))
for el in solid2_liste:
    temp = gesamt_Wand
    try:
        gesamt_Wand = DB.BooleanOperationsUtils.ExecuteBooleanOperation(gesamt_Wand,el,DB.BooleanOperationsType.Difference)
        temp.Dispose()
    except Exception as e:print(e)


rohre_liste = []

with forms.ProgressBar(title='{value}/{max_value} Rohre - Daten ermitteln',cancellable=True, step=30) as pb:
    for n,elemid in enumerate(rohrs):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n+1,len(rohrs))

        rohr = Rohr(elemid)
        rohr.Line = rohr.get_Line()
        if not rohr.Line:
            continue
        if rohr.Ignorierne():
            continue
        if rohr.richtung == 'HORIZONTAL':
            result1 = gesamt_Wand.IntersectWithCurve(rohr.Line,opt1)
            result2 = gesamt_Wand.IntersectWithCurve(rohr.Line,opt2)
            if result1.SegmentCount > 0 and result2.SegmentCount > 0:
                rohr.Anzahl = result2.SegmentCount
                if result1.SegmentCount - result2.SegmentCount != 1:
                    rohr.Pruefen = 'Warnung: Leitung endet in Objekt'
                result1.Dispose()
                result2.Dispose()
            else:
                result1.Dispose()
                result2.Dispose()
                rohr.get_Brandklass_Text()
                rohr.Line.Dispose()
                rohre_liste.append(rohr)
                continue

            for el in Wall_liste:
                solids = el.Solids
                if len(solids) == 0:
                    continue
                for solid in solids:
                    result1 = solid.IntersectWithCurve(rohr.Line,opt1)
                    result2 = solid.IntersectWithCurve(rohr.Line,opt2)
                    if result1.SegmentCount > 0 and result2.SegmentCount > 0:
                        if el.Kategorie not in rohr.Brandklass.keys():
                            rohr.Brandklass[el.Kategorie] = {}
                        if el.Typ not in rohr.Brandklass[el.Kategorie].keys():
                            rohr.Brandklass[el.Kategorie][el.Typ] = {}
                        if el.Brandschutz not in rohr.Brandklass[el.Kategorie][el.Typ].keys():
                            rohr.Brandklass[el.Kategorie][el.Typ][el.Brandschutz] = 1
                        else:
                            rohr.Brandklass[el.Kategorie][el.Typ][el.Brandschutz] += 1
                    result1.Dispose()
                    result2.Dispose()
        else:
            result1 = gesamt_decke.IntersectWithCurve(rohr.Line,opt1)
            result2 = gesamt_decke.IntersectWithCurve(rohr.Line,opt2)
            if result1.SegmentCount > 0 and result2.SegmentCount > 0:
                rohr.Anzahl = result2.SegmentCount
                if result1.SegmentCount - result2.SegmentCount != 1:
                    rohr.Pruefen = 'Warnung: Leitung endet in Objekt'
                result1.Dispose()
                result2.Dispose()
            else:
                result1.Dispose()
                result2.Dispose()
                rohr.get_Brandklass_Text()
                rohr.Line.Dispose()
                rohre_liste.append(rohr)
                continue
            for el in Decken_liste:
                
                solids = el.Solids
                if len(solids) == 0:
                    continue
                for solid in solids:
                    result1 = solid.IntersectWithCurve(rohr.Line,opt1)
                    result2 = solid.IntersectWithCurve(rohr.Line,opt2)
                    if result1.SegmentCount > 0 and result2.SegmentCount > 0:
                        if el.Kategorie not in rohr.Brandklass.keys():
                            rohr.Brandklass[el.Kategorie] = {}
                        if el.Typ not in rohr.Brandklass[el.Kategorie].keys():
                            rohr.Brandklass[el.Kategorie][el.Typ] = {}
                        if el.Brandschutz not in rohr.Brandklass[el.Kategorie][el.Typ].keys():
                            rohr.Brandklass[el.Kategorie][el.Typ][el.Brandschutz] = 1
                        else:
                            rohr.Brandklass[el.Kategorie][el.Typ][el.Brandschutz] += 1
                    result1.Dispose()
                    result2.Dispose()
        rohr.get_Brandklass_Text()
        rohr.Line.Dispose()
        rohre_liste.append(rohr)

t = DB.Transaction(doc,'Brandschott Rohre')
t.Start()
with forms.ProgressBar(title='{value}/{max_value} Daten schreiben',cancellable=True, step=int(len(rohre_liste)/200)+1) as pb:
    for n,rohr in enumerate(rohre_liste):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n+1,len(rohre_liste))
        rohr.wert_schreiben()
t.Commit()
