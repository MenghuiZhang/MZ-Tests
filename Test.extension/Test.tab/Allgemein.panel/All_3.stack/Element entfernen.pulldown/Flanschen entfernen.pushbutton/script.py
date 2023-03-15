# coding: utf8
from rpw import revit,DB
from pyrevit import script, forms
from System.Collections.Generic import List



__title__ = "Flansche entfernen"
__doc__ = """

[2022.04.13]
Version: 1.1
"""
__authors__ = "Menghui Zhang"


logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc

class Flansch:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.IsFlansch = False
        try:
            if self.elem.Symbol.FamilyName != 'MAGI-FLG-WELD-COLLAR-PN16':
                return
            self.IsFlansch = True

        except:
            return
        
        self.rohr = None
        self.bogen = None
        
    def get_bogenundrohe(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                refs = conn.AllRefs
                for ref in refs:
                    owner = ref.Owner
                    if owner.Category.Id.IntegerValue == -2008044:
                        self.rohr = ref
                    elif owner.Category.Id.IntegerValue == -2008049:
                        self.bogen = ref
    
    def verbinden(self):
        try:
            if self.rohr and self.bogen:
                try:
                    self.rohr.Origin = self.bogen.Origin
                    self.bogen.ConnectTo(self.rohr)
                except Exception as e:
                    logger.error(e)
            else:
                logger.error(self.elemid.ToString())
        except:
            logger.error(self.elemid.ToString())

# Liste = uidoc.Selection.GetElementIds()
Liste = []
Liste_elemid = []
for el in uidoc.Selection.GetElementIds():
    fs = Flansch(el)
    if fs.IsFlansch:
        Liste_elemid.append(el)
        fs.get_bogenundrohe()
        Liste.append(fs)

t = DB.Transaction(doc,'Flanschen ersetzen')
t.Start()
doc.Delete(List[DB.ElementId](Liste_elemid))
doc.Regenerate()

with forms.ProgressBar(title="{value}/{max_value} Flansche",cancellable=True, step=1) as pb:

    for n,fs in enumerate(Liste):
        if pb.cancelled:
            t.RollBack()
            script.exit()
        
        pb.update_progress(n+1, len(Liste))
        fs.verbinden()
        doc.Regenerate()

t.Commit()


# t = DB.Transaction(doc,'Flanschen ersetzen')
# t.Start()
# with forms.ProgressBar(title="{value}/{max_value} Element",cancellable=True, step=1) as pb:

#     for n,elemid in enumerate(Liste):
#         if pb.cancelled:
#             t.RollBack()
#             script.exit()
        
#         pb.update_progress(n+1, len(Liste))
#         fs = Flansch(elemid)
#         if fs.IsFlansch:
#             fs.get_bogenundrohe()
#             doc.Delete(fs.elemid)
#             doc.Regenerate()
#             fs.verbinden()
#             doc.Regenerate()

# t.Commit()