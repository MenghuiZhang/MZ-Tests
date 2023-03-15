# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
import System
import clr
import os
from pyrevit.forms import WPFWindow, SelectFromList
from Autodesk.Revit.DB import *
clr.AddReference("RevitServices")
from RevitServices.Transactions import TransactionManager

start = time.time()

__title__ = "0.55 Familien exportieren"
__doc__ = """Familien exportieren"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

from pyIGF_logInfo import getlog
getlog(__title__)


uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
app = rpw.revit.app
uiapp = rpw.revit.uiapp


Family = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))


class NameAedern(WPFWindow):
    def __init__(self, xaml_file_name,Familie):
        self.FamilieName = Familie.Name
        WPFWindow.__init__(self, xaml_file_name)
        self._set_comboboxes()
        self.Path = None
        self.Ordner = None

    def _set_comboboxes(self):
        self.AltName.Text = self.FamilieName
        self.NeuName.Text = '_L_IGF_430_' + self.FamilieName +'.rfa'

    @property
    def neuNameID(self):
        p = None
        if self.NeuName.Text:
            p = self.NeuName.Text
        return p

    def run(self, sender, args):
        if self.neuNameID:
            Name = self.neuNameID
            OrdnerName = Name.split('_')
            Ordner1 = '_' + OrdnerName[1]
            Ordner2 = None
            if len(OrdnerName[3]) == 3:
                Ordner2 = OrdnerName[3]
            if Ordner2:
                path = 'R:\\Vorlagen\\_IGF\\_Familien'+'\\'+ Ordner1 + '\\'+ Ordner2 + '\\'+ Name
                self.Ordner = 'R:\\Vorlagen\\_IGF\\_Familien'+'\\'+ Ordner1 + '\\'+ Ordner2
                self.Path = path
            else:
                path = 'R:\\Vorlagen\\_IGF\\_Familien'+'\\'+ Ordner1 + '\\'+ Name
                self.Path = path
                self.Ordner = 'R:\\Vorlagen\\_IGF\\_Familien'+'\\'+ Ordner1

        self.Close()
    def skip(self, sender, args):
        self.Close()
    def cancel(self, sender, args):
        self.Close()
        script.exit()


def mkdir(path):
    path=path.strip()
    isExists=os.path.exists(path)
    if not isExists:
        os.makedirs(path)

LuftZubehoer = []
LuftFormteil = []
Luftdurchlass = []
HLS = []
Rohrzubehoer = []
Rohrformteile = []


for el in Family:
    familie = el
    category = familie.FamilyCategory.Name
    if category == 'Luftkanalzubehör':
        LuftZubehoer.append(familie)
    elif category == 'Luftkanalformteile':
        LuftFormteil.append(familie)
    elif category == 'Luftdurchlässe':
        Luftdurchlass.append(familie)
    elif category == 'HLS-Bauteile':
        HLS.append(familie)
    elif category == 'Rohrzubehör':
        Rohrzubehoer.append(familie)
    elif category == 'Rohrformteile':
        Rohrformteile.append(familie)

saveAsOptions = SaveAsOptions()
saveAsOptions.OverwriteExistingFile = True
saveAsOptions.MaximumBackups = 3
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
            if not el.IsEditable:
                continue

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

            paramTw = None
            try:
                trans1.EnsureInTransaction(fdoc)
                pt = DB.ParameterType.Material
                pg = DB.BuiltInParameterGroup.PG_MATERIALS
                paramTw = m_familyMgr.AddParameter('IGF_Material', pg, pt, True)
                fdoc.Regenerate()
                trans1.ForceCloseTransaction()
                logger.info('Parameter von Familie {} wurde erstellt'.format(el.Name))
            except Exception as e:
                paramTw = m_familyMgr.get_Parameter('IGF_Material')
                trans1.ForceCloseTransaction()

            assoFP2Mate(fdoc,paramTw)

            Nameaendern = NameAedern('window.xaml',el)
            Nameaendern.ShowDialog()
            FamilyPath = Nameaendern.Path
            ordner = Nameaendern.Ordner
            FamName = Nameaendern.neuNameID
            if not FamilyPath:
                continue
            try:
                mkdir(ordner)
            except:
                pass
            try:
                fdoc.SaveAs(FamilyPath)
            except Exception as e:
                fdoc.SaveAs(FamilyPath,saveAsOptions)
                bachup = FamilyPath[:-4] + '.0001.rfa'
                os.remove(bachup)


            # fami = doc.Application.OpenDocumentFile(FamilyPath)
            # Material = fami.FamilyManager.get_Parameter('IGF_Material')
            # assoFP2Mate(fdoc,Material)
            #
            # fdoc.SaveAs(FamilyPath,saveAsOptions)
            # bachup = FamilyPath[:-4] + '.0001.rfa'
            # os.remove(bachup)


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
AddFamilyParam(LuftFormteil,'Luftkanalformteile')
AddFamilyParam(Luftdurchlass,'Luftdurchlässe')
AddFamilyParam(HLS,'HLS-Bauteile')
AddFamilyParam(Rohrzubehoer,'Rohrzubehör')
AddFamilyParam(Rohrformteile,'Rohrformteile')


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
