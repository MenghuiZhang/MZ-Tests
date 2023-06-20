# coding: utf8
# from IGF_log import getlog,getloglocal
# from pyrevit import forms
from rpw import revit,DB
'''from eventhandler import VERBINDEN,ExternalEvent
import os'''


__title__ = "ULK Test"
__doc__ = """

[2022.04.13]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

# try:
#     getlog(__title__)
# except:
#     pass

# try:
#     getloglocal(__title__)
# except:
#     pass

doc = revit.doc
uidoc = revit.uidoc
t = DB.Transaction(doc,'text')
t.Start()
cl = uidoc.Selection.GetElementIds()
for elid in cl:
    print(elid)
    doc.Regenerate()
    el = doc.GetElement(elid)
    conns = el.MEPModel.ConnectorManager.Connectors
    for conn in conns:
        if conn.IsConnected:
            refs = conn.AllRefs
            for ref in refs:
                if ref.Owner.Category.Name in ['Rohrformteile']:
                    conns1 = ref.Owner.MEPModel.ConnectorManager.Connectors
                    if conns1.Size == 2:
                        for conn0 in conns1:
                            if not conn.IsConnectedTo(conn0):
                                connsize = str(int(round(conn0.Radius*304.8*2)))
                                typ_name = 'DN' + connsize
                                for typ in el.Symbol.Family.GetFamilySymbolIds():
                                    typname = doc.GetElement(typ).get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                                    # print(typname,typ_name)
                  
                                    if typname == typ_name:
                                        try:el.ChangeTypeId(typ)
                                        except:pass
t.Commit()
# class AktuelleBerechnung(forms.WPFWindow):
#     def __init__(self):
#         self.transitionfitting = VERBINDEN()
#         self.transitionfittingEvent = ExternalEvent.Create(self.transitionfitting)
#         forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
#         self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))

#     def erstellen(self, sender, args):
#         self.transitionfittingEvent.Raise()   
       
# wind = AktuelleBerechnung()
# wind.transitionfitting.GUI = wind
# wind.Show()
