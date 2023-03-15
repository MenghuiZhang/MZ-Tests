# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import System
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from Autodesk.Revit.DB import *
clr.AddReference("RevitServices")
from RevitServices.Transactions import TransactionManager

start = time.time()

__title__ = "0.50 FamilyPara(IGF_Material) erstellen"
__doc__ = """FamilyPara(IGF_Material) erstellen. 
Kategorien: Luftkanalzubehör, Luftdurchlässe, Luftkanalformteile,
HLS-Bauteile, Rohrzubehör, Rohrformteile"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp

from pyIGF_logInfo import getlog
getlog(__title__)


Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.FamilyInstance))

LuftZubehoer = []
LuftFormteil = []
Luftdurchlass = []
HLS = []
Rohrzubehoer = []
Rohrformteile = []
LuftZubehoerid = []
LuftFormteilid = []
Luftdurchlassid = []
HLSid = []
Rohrzubehoerid = []
Rohrformteileid = []

for el in Family:
    familie = el.Symbol.Family
    category = familie.FamilyCategory.Name
    if category == 'Luftkanalzubehör':
        if not familie.Id in LuftZubehoerid:
            LuftZubehoer.append(familie)
            LuftZubehoerid.append(familie.Id)
    elif category == 'Luftkanalformteile':
        if not familie.Id in LuftFormteilid:
            LuftFormteil.append(familie)
            LuftFormteilid.append(familie.Id)
    elif category == 'Luftdurchlässe':
        if not familie.Id in Luftdurchlassid:
            Luftdurchlass.append(familie)
            Luftdurchlassid.append(familie.Id)
    elif category == 'HLS-Bauteile':
        if not familie.Id in HLSid:
            HLS.append(familie)
            HLSid.append(familie.Id)
    elif category == 'Rohrzubehör':
        if not familie.Id in Rohrzubehoerid:
            Rohrzubehoer.append(familie)
            Rohrzubehoerid.append(familie.Id)
    elif category == 'Rohrformteile':
        if not familie.Id in Rohrformteileid:
            Rohrformteile.append(familie)
            Rohrformteileid.append(familie.Id)


class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source = FamilySource.Family
        overwriteParameterValues = True
        return True


def AddFamilyParam(Familydaten,Kate):
    title = "{value}/{max_value} Familien in Kategorie "+ Kate
    with forms.ProgressBar(title=title, cancellable=True, step=1) as pb:
        n = 0
        trans1 = TransactionManager.Instance
        trans1.ForceCloseTransaction()
        for el in Familydaten:
            if pb.cancelled:
                script.exit()
            n = n + 1
            pb.update_progress(n, len(Familydaten))

            fdoc = None
            m_familyMgr = None
            try:
                fdoc = doc.EditFamily(el)
                m_familyMgr = fdoc.FamilyManager
                logger.info('Familie {} bearbeiten'.format(el.Name))
            except Exception as exc:
                logger.error(exc)
            if not fdoc:
                continue

            try:
                trans1.EnsureInTransaction(fdoc)
                pt = DB.ParameterType.Material
                pg = DB.BuiltInParameterGroup.PG_MATERIALS
                paramTw = m_familyMgr.AddParameter('IGF_Material', pg, pt, True)
                fdoc.Regenerate()
                trans1.ForceCloseTransaction()
                logger.info('Parameter von Familie {} wurde erstellt'.format(el.Name))
            except Exception as e:
                logger.error(el.Name +': '+e.ToString())
                trans1.ForceCloseTransaction()
            Material = m_familyMgr.get_Parameter('IGF_Material')
            assoFP2Mate(fdoc,Material)
            if n == 5:
                LoadFamily(el,doc,el.Name)

            logger.info(30*'-')

def LoadFamily(family,Prodoc,famName):
    famdoc = Prodoc.EditFamily(family)
    trans2 = TransactionManager.Instance
    trans2.EnsureInTransaction(Prodoc)
    for symbol in family.GetFamilySymbolIds():
        if not Prodoc.GetElement(symbol).IsActive:
            Prodoc.GetElement(symbol).Activate()
        break
    trans2.ForceCloseTransaction()
    try:
        famdoc.LoadFamily(Prodoc, FamilyLoadOptions())
        logger.info('Familie {} wurde geladen'.format(famName))
    except Exception as ex:
        logger.error(ex.ToString())
    famdoc.Close(False)

def assoFP2Mate(famdoc,Mate):
    Extrusion = DB.FilteredElementCollector(famdoc).OfClass(clr.GetClrType(DB.Extrusion)).WhereElementIsNotElementType()
    ExtrusionIds = Extrusion.ToElementIds()
    Blend = DB.FilteredElementCollector(famdoc).OfClass(clr.GetClrType(DB.Blend)).WhereElementIsNotElementType()
    BlendIds = Blend.ToElementIds()
    Sweep = DB.FilteredElementCollector(famdoc).OfClass(clr.GetClrType(DB.Sweep)).WhereElementIsNotElementType()
    SweepIds = Sweep.ToElementIds()
    SweptBlend = DB.FilteredElementCollector(famdoc).OfClass(clr.GetClrType(DB.SweptBlend)).WhereElementIsNotElementType()
    SweptBlendIds = SweptBlend.ToElementIds()
    Revolution = DB.FilteredElementCollector(famdoc).OfClass(clr.GetClrType(DB.Revolution)).WhereElementIsNotElementType()
    RevolutionIds = Revolution.ToElementIds()
    Extrusion.Dispose()
    Blend.Dispose()
    Sweep.Dispose()
    SweptBlend.Dispose()
    Revolution.Dispose()
    try:
        t = Transaction(famdoc,'BindingParameter')
        t.Start()
        if len(ExtrusionIds) != 0:
            Matezuweisen(famdoc,ExtrusionIds,Mate)
        if len(ExtrusionIds) != 0:
            Matezuweisen(famdoc,BlendIds,Mate)
        if len(ExtrusionIds) != 0:
            Matezuweisen(famdoc,SweepIds,Mate)
        if len(ExtrusionIds) != 0:
            Matezuweisen(famdoc,SweptBlendIds,Mate)
        if len(ExtrusionIds) != 0:
            Matezuweisen(famdoc,RevolutionIds,Mate)
        t.Commit()
    except Exception as e:
        logger.info(e)
def Matezuweisen(famdoc,ids,Mate):
    for id in ids:
        ele = famdoc.GetElement(id)
        familyMgr = famdoc.FamilyManager
        materialElementPara = ele.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        try:
            familyMgr.AssociateElementParameterToFamilyParameter(materialElementPara, Mate)
        except Exception as e:
            logger.error(e)
AddFamilyParam(LuftZubehoer,'Luftkanalzubehör')
# AddFamilyParam(LuftFormteil,'Luftkanalformteile')
# AddFamilyParam(Luftdurchlass,'Luftdurchlässe')
# AddFamilyParam(HLS,'HLS-Bauteile')
# AddFamilyParam(Rohrzubehoer,'Rohrzubehör')
# AddFamilyParam(Rohrformteile,'Rohrformteile')


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
