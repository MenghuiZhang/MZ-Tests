# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms

__title__ = "0.40 Einbauebene ändern"
__doc__ = """

Die Bauteile werden mit den neue Ebene verknüpft.

[2021.12.13]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
uidoc = revit.uidoc
doc = revit.doc

Levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
levels = Levels_collector.ToElementIds()

if not levels:
    logger.error("Keine Ebene in aktueller Projekt gefunden")
    script.exit()

levels_dict = {}
for el in Levels_collector:
    name = el.Name
    levels_dict[name] = el.Id

Levels_collector.Dispose()

class Element_IGF:
    def __init__(self,elemid,ebene,d_hoehe,ebene_alt):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.ebene_alt = ebene_alt
        try:
            self.category = self.elem.Category.CategoryType.ToString()
        except:
            self.category = ''
        
        self.hoehe_Param_1 = self.elem.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
        self.hoehe_Param_2 = self.elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
        self.hoehe_versatz = self.elem.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)

        self.ebene_neu = ebene
        self.d_hoehe = d_hoehe
    
        if self.category == 'Model':
            if self.hoehe_Param_2 and self.hoehe_versatz:
                self.hoehe_neu = self.hoehe_versatz.AsDouble()+ self.d_hoehe
            else:
                self.hoehe_neu = None
    
    # def change_reference(self):
    #     try:
    #         references = self.elem.References
    #         reference_bauteil = None
    #         for ref in references:
    #             if ref.ElementId != self.ebene_alt:
    #                 reference_bauteil = ref
    #         references.Clear()
    #         references.Append(reference_bauteil)
    #         references.Append(DB.Reference(doc.GetElement(self.ebene_neu)))

    #     except Exception as e:
    #         logger.error(e)


    def werte_schreiben(self):
        if self.hoehe_Param_1:
            try:
                self.hoehe_Param_1.Set(self.ebene_neu)
            except:
                pass
        else:
            if self.hoehe_Param_2:
                try:
                    self.hoehe_Param_2.Set(self.ebene_neu)
                except:
                    pass
            if self.hoehe_versatz:
                try:
                    self.hoehe_versatz.Set(self.hoehe_neu)
                except:
                    pass
    


class ChangeEbene(forms.WPFWindow):
    def __init__(self, xaml_file_name,ItemsSource):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.ebene_alt.ItemsSource = sorted(ItemsSource.keys())
        self.ebene_neu.ItemsSource = sorted(ItemsSource.keys())
    
    def ok(self, sender, args):
        self.Close()
    def abbrechen(self, sender, args):
        self.Close()
        script.exit()
        
changeebene = ChangeEbene('window.xaml',levels_dict)
try:
    changeebene.ShowDialog()
except Exception as e1:
    logger.error(e1)
    changeebene.Close()
    script.exit()

Ebene_alt = changeebene.ebene_alt.Text
Ebene_neu = changeebene.ebene_neu.Text

elemid_ebene_alt = levels_dict[Ebene_alt]
elemid_ebene_neu = levels_dict[Ebene_neu]

ebene_alt_hoehe = doc.GetElement(elemid_ebene_alt).LookupParameter('Ansicht').AsDouble()
ebene_neu_hoehe = doc.GetElement(elemid_ebene_neu).LookupParameter('Ansicht').AsDouble()

d_ebene = ebene_alt_hoehe - ebene_neu_hoehe

Elementen = doc.GetElement(elemid_ebene_alt).GetDependentElements(None)


liste_element = []
with forms.ProgressBar(title = "{value}/{max_value} Element",cancellable=True, step=int(len(Elementen)/1000)+10) as pb:
    for n, elem_id in enumerate(Elementen):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(Elementen))
        elem = Element_IGF(elem_id,elemid_ebene_neu,d_ebene,ebene_alt_hoehe)
        if elem.category == 'Model':
            liste_element.append(elem)


if forms.alert('Ebene ändern?', ok=False, yes=True, no=True):
    t = DB.Transaction(doc,'test')
    t.Start()
    with forms.ProgressBar(title = "{value}/{max_value} Element",cancellable=True, step=int(len(liste_element)/1000)+10) as pb:
        for n, elem in enumerate(liste_element):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(liste_element))

            elem.werte_schreiben()
        
        
    t.Commit()