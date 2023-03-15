# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw

__title__ = "0.03 Wert schreiben MEP Räume"
__doc__ = """

Wert schreiben
Parameter muss Exemplar sein.

für Ja/Nein Parameter bitte 1/0 eingeben.
1: Ja
0: Nein 

"""

__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc

try:
    getlog(__title__)
except:
    pass


# Kategorie_dict = {
#     'Allgemeines Modell':'',
#     'Detailelemente':'',
#     'Flexible Rohre':'',
#     'Flexkanäle':'',
#     'HLS-Bauteile':'',
#     'Luftdurchlässe':'',
#     'Luftkanalformteile':'',
#     'Luftkanalzubehör':'',
#     'Luftkanäle':'',
#     'MEP-Räume':'',   
# }

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

space = doc.GetElement(spaces[0])
param_liste = []
for param in space.Parameters:
    if not param.IsReadOnly and param.UserModifiable:
        param_liste.append(param.Definition.Name)
param_liste.sort()

class ParamsMEP(forms.WPFWindow):
    def __init__(self,xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.parameter.ItemsSource = param_liste
    
    def ok(self, sender, args):
        self.Close()
    def abbrechen(self, sender, args):
        self.Close()
        script.exit()

paramsmep = ParamsMEP('window.xaml')
try:
    paramsmep.ShowDialog()
except Exception as e:
    logger.error(e)
    paramsmep.Close()

para = paramsmep.parameter.Text
wert = paramsmep.wert.Text
t = DB.Transaction(doc,para)
t.Start()
with forms.ProgressBar(title = "{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb:
    for n, elem_id in enumerate(spaces):
        if pb.cancelled:
            t.RollBack()
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        raum = doc.GetElement(elem_id)
        param = raum.LookupParameter(para)
        if param.StorageType.ToString() == 'Integer':
            try:
                param.Set(int(wert))
            except:
                pass
        elif param.StorageType.ToString() == 'Double':
            try:
                param.SetValueString(str(wert))
            except:
                pass
        elif param.StorageType.ToString() == 'String':
            try:
                param.Set(str(wert))
            except:
                pass
t.Commit()
t.Dispose()