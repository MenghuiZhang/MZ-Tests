# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from rpw import revit, DB
from pyrevit import script, forms
from pyrevit.forms import WPFWindow

__title__ = "1.40 Höhenkoten der Ebene"
__doc__ = """
erstellt globale Parameter HöhenkotenEbene.
Bezuglich Projekt-NullPunktes und realen NullPunktes

Die folgenden Parameter werden erstellt.

IGF_EbeneName_ÜPN
IGF_EbeneName_ÜNN

[2021.11.25]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number

config = script.get_config(name+number+'Höhenkoten')

Levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
levels = Levels_collector.ToElementIds()

if not levels:
    logger.error("Keine Ebene in aktueller Projekt gefunden")
    script.exit()

levels_dict = {}
for el in Levels_collector:
    name = el.Name
    levels_dict[name] = el.Id

levels_liste = levels_dict.keys()[:]
levels_liste.sort()

class EbeneAuswahl(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self.ebenename.ItemsSource = levels_liste

        self.read_config()

    def read_config(self):
        try:
            if config.nn not in levels_liste:
                config.nn = ""
            else:
                try:
                    self.ebenename.Text = config.nn
                except:
                    self.ebenename.Text = config.nn = ""
        except:
            pass

        try:
            self.height.Text = config.height
                    
        except:
            self.height.Text = config.height = ""
        
        try:
            self.hoehe.IsChecked = config.hoehe
                    
        except:
            self.hoehe.IsChecked = config.hoehe = False
        
        try:
            self.name.IsChecked = config.name
                    
        except:
            self.name.IsChecked = config.name = False
        
    def write_config(self):
        config.height = self.height.Text
        config.nn = self.ebenename.Text
        config.hoehe = self.hoehe.IsChecked
        config.name = self.name.IsChecked
        script.save_config()

    def ok(self, sender, args):
        self.Close()
        self.write_config()

    def abbrechen(self, sender, args):
        self.Close()
        script.exit()

ebeneAuswahl = EbeneAuswahl('window.xaml')
try:
    ebeneAuswahl.ShowDialog()
except Exception as e:
    ebeneAuswahl.Close()
    logger.error(e)
    script.exit()

if ebeneAuswahl.hoehe.IsChecked:
    hoehe_nn = float(ebeneAuswahl.height.Text)
else:
    if ebeneAuswahl.name.IsChecked:
        nn = levels_dict[ebeneAuswahl.ebenename.Text]
        hoehe_nn = get_value(doc.GetElement(nn).LookupParameter('Ansicht'))
    else:
        logger.error('Nullpunkt nicht difiniert')
        script.exit()

class Ebenen:
    def __init__(self, element_id,_nn):
        self.element_id = element_id
        self.nn = _nn
        self.element = doc.GetElement(self.element_id)
        self.name = self.element.Name
        self.hoehe = get_value(self.element.LookupParameter('Ansicht'))
        self.istgebaeudegeschloss = get_value(self.element.LookupParameter('Gebäudegeschoss'))
        self.upn = self.hoehe
        self.unn = self.hoehe - self.nn
    
    def create_unn_upn(self):
        Unn = 'IGF_' + self.name + '_ÜNN'
        gp_unn = doc.GetElement(DB.GlobalParametersManager.FindByName(doc,Unn))
        if not gp_unn:
            gp_unn = DB.GlobalParameter.Create(doc,Unn,DB.ParameterType.Number)
        gp_unn.SetFormula(str(int(round(self.unn))))
        gp_unn.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS

        Upn = 'IGF_' + self.name + '_ÜPN'
        gp_upn = doc.GetElement(DB.GlobalParametersManager.FindByName(doc,Upn)) 
        if not gp_upn:
            gp_upn = DB.GlobalParameter.Create(doc,Upn,DB.ParameterType.Number)
        gp_upn.SetFormula(str(int(round(self.upn))))
        gp_upn.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS


Ebene_liste = []
with forms.ProgressBar(title='{value}/{max_value} Ebenen', cancellable=True, step=1) as pb:
    
    for n, level_id in enumerate(levels):
        if pb.cancelled:
            script.exit()
        
        pb.update_progress(n + 1, len(levels))
        if ebeneAuswahl.name:
            if levels_dict[ebeneAuswahl.ebenename.Text].ToString() == level_id.ToString():
                continue
                
        ebene = Ebenen(level_id,hoehe_nn)

        Ebene_liste.append(ebene)

if forms.alert('Globale Parameter erstellen?', ok=False, yes=True, no=True):
    t = DB.Transaction(doc,'Höhenkoten der Ebene')
    t.Start()
    gp_nn = doc.GetElement(DB.GlobalParametersManager.FindByName(doc,'IGF_Ebene_NN'))
    if not gp_nn:
        gp_nn = DB.GlobalParameter.Create(doc,'IGF_Ebene_NN1',DB.ParameterType.Number)
    gp_nn.SetFormula(str(int(round(hoehe_nn))))
    gp_nn.GetDefinition().ParameterGroup = DB.BuiltInParameterGroup.PG_CONSTRAINTS

    with forms.ProgressBar(title='{value}/{max_value} globale Parameter', cancellable=True, step=1) as pb1:
        for n, ebene in enumerate(Ebene_liste):
            if pb1.cancelled:
                t.RollBack()
                script.exit()

            pb1.update_progress(n + 1, len(Ebene_liste))
            ebene.create_unn_upn()
    t.Commit()
    t.Dispose()