# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import Muster_Pruefen
from rpw import revit, DB
from pyrevit import script, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List

__title__ = "0.43 Einbauort, Familie basiert"
__doc__ = """

schreibt Einbauort(Raumnummer) in Bauteile ein.

Parameter:

IGF_X_Einbauort

"""
__authors__ = "Menghui Zhang"


try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc
PHASE = doc.Phases[0]

logger = script.get_logger()



dict_fam = {}
def get_dict_fam():
    categories = doc.Settings.Categories
    cate_model = []
    for el in categories:
        if el.CategoryType.ToString() == 'Model':
            cate_model.append(el.Id)

    categoryIds = List[DB.ElementId](cate_model)
    Filter = DB.ElementMulticategoryFilter(categoryIds)
    Fam_Collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).WherePasses(Filter)
    for el in Fam_Collector:
        cate = el.Category.Name
        Famname = el.Symbol.FamilyName
        if cate not in dict_fam:
            dict_fam[cate] = {}
        if Famname not in dict_fam[cate].keys():
            dict_fam[cate][Famname] = [el]
        else:
            dict_fam[cate][Famname].append(el)
get_dict_fam()

class Children(object):
    def __init__(self,name,checked,obj):
        self.name = name
        self.checked = checked
        self.object = obj
        self.children = ObservableCollection[Children]()
        self.parent = None
        self.expand = False

Liste_Cat = ObservableCollection[Children]()
def get_List_Cat():
    for el in sorted(dict_fam.keys()):
        Fams = dict_fam[el]
        cate = Children(el,False,Fams)
        for fam in sorted(Fams.keys()):
            Fam = Children(fam,False,Fams[fam])
            Fam.parent = cate
            cate.children.Add(Fam)         
        Liste_Cat.Add(cate)
get_List_Cat()

class Familieauswahl(forms.WPFWindow):
    def __init__(self,xaml_file_name,liste_module):
        self.liste_module = liste_module
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.treeView1.ItemsSource = liste_module
    
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        item = sender.DataContext
        for child in sender.DataContext.children:
            child.checked = Checked
            child.expand = True
        parent = item.parent
        
        if parent:
            parent.expand = True
            results = [child.checked for child in parent.children]
            if all(results):
                parent.checked = True
            elif any(results):
                parent.checked = None
            elif not any(results):
                parent.checked = False
        self.treeView1.Items.Refresh()

    
    def all(self,sender,args):
        for item in self.treeView1.Items:
            item.checked = True
            for panel in item.children:
                panel.checked = True
        self.treeView1.Items.Refresh()

    def kein(self,sender,args):
        for item in self.treeView1.Items:
            item.checked = False
            for panel in item.children:
                panel.checked = False
        self.treeView1.Items.Refresh()

    def zu(self,sender,args):
        for item in self.treeView1.Items:
            item.expand = False
        self.treeView1.Items.Refresh()

    def aus(self,sender,args):
        for item in self.treeView1.Items:
            item.expand = True
        self.treeView1.Items.Refresh()

    def ab(self,sender,args):
        self.Close()
        script.exit()
    
    def ok(self,sender,args):
        self.Close()

FamilienAuswahl = Familieauswahl('window.xaml',Liste_Cat)
try:
    FamilienAuswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienAuswahl.Close()
    script.exit()


class Bauteile(object):
    def __init__(self,elem):
        self.elem = elem
        try:
            self.nummer = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
        except:
            self.nummer = ''
    
    def wert_schreiben(self):
        try:
            self.elem.LookupParameter('IGF_X_Einbauort').Set(self.nummer)
        except:
            pass

bearbeitet = []
nichtbearbeitet = []
dict_alle = {}
for el in Liste_Cat:
    for ele in el.children:
        if ele.checked:
            nichtbearbeitet.append(ele.name)
            dict_alle[ele.name] = ele.object

def main():
    t = DB.Transaction(doc,'Einbauort')
    t.Start()
    if forms.alert('Einbauort schreiben?',yes=True,no=True,ok=False):
        with forms.ProgressBar(cancellable=True, step=10) as pb2:
            for n,bauteile in enumerate(dict_alle.keys()):
                systeme_title = bauteile + ' --- ' + str(n+1) + '/' + str(len(dict_alle.keys())) + ' Familien'
                pb2.title = '{value}/{max_value} Bauteile in Familien ' +  systeme_title
                pb2.step = int((len(dict_alle[bauteile])) /1000) + 10

                for n1,bauteil in enumerate(dict_alle[bauteile]):
                    if pb2.cancelled:
                        if forms.alert('bisherige Änderung behalten?',yes=True,no=True,ok=False):
                            t.Commit()
                            logger.info('Folgenede Familien sind bereits bearbeitet.')
                            for el in bearbeitet:
                                logger.info(el)
                            logger.error('---------------------------------------')
                            logger.warning('Folgenede Familien sind nicht bearbeitet.')
                            for el in nichtbearbeitet:
                                logger.warning(el)
                        else:
                            t.RollBack()
                        
                        return
                    if Muster_Pruefen(bauteil):continue
                    bauteil_neu = Bauteile(bauteil)
                    try:bauteil_neu.wert_schreiben()
                    except Exception as e:print(e)

                    pb2.update_progress(n1+1, len(dict_alle[bauteile]))
                bearbeitet.append(bauteile)
                nichtbearbeitet.remove(bauteile)
                
    t.Commit()
main()