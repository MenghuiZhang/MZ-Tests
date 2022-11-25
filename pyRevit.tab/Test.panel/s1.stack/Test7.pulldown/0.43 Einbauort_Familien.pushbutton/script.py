# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import Muster_Pruefen
from rpw import revit, DB
from pyrevit import script, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List

__title__ = "0.43 Einbauort für RVs"
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
    cate_model = []

    cate_model.append(DB.ElementId(DB.BuiltInCategory.OST_MechanicalEquipment))
    cate_model.append(DB.ElementId(DB.BuiltInCategory.OST_PlumbingFixtures))

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

familienAuswahl = Familieauswahl('window.xaml',Liste_Cat)
try:
    familienAuswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    familienAuswahl.Close()
    script.exit()


class ME:
    def __init__(self,elem):
        self.elem = elem
        self.liste_0 = []
        self.liste_1 = []
        self.vr = None
        self.rr = None
        try:
            self.nummer = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
        except:
            self.nummer = ''
        if familienAuswahl.vorlauf.IsChecked:
            self.get_vsr_vl(self.elem)  
        if familienAuswahl.ruecklauf.IsChecked:
            self.get_vsr_rl(self.elem)  


    def get_vsr_vl(self,elem):
        if self.vr:return
        elemid = elem.Id.ToString()
        self.liste_0.append(elemid)
        cate = elem.Category.Name

        

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                


                if conns.Size > 3 and cate != 'HLS-Bauteile':return
                for conn in conns:
                    if elemid == self.elem.Id.ToString():
                        
                        try:
                            if conn.PipeSystemType.ToString() != 'SupplyHydronic':
                                continue
                        except:continue
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste_0:
                            if owner.Category.Name == 'Rohrzubehör':
                                self.vr = owner
                                return
                                                                        
                            self.get_vsr_vl(owner)
        
    def get_vsr_rl(self,elem):
        if self.rr:return
        elemid = elem.Id.ToString()
        self.liste_1.append(elemid)
        cate = elem.Category.Name

        

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 3 and cate != 'HLS-Bauteile':return
                for conn in conns:
                    if elemid == self.elem.Id.ToString():
                        try:
                            if conn.PipeSystemType.ToString() != 'ReturnHydronic':
                                continue
                        except:continue
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste_1:
                            if owner.Category.Name == 'Rohrzubehör':
                                self.rr = owner
                                return                                      
                            self.get_vsr_rl(owner)

    def wert_schreiben_Alle(self):
        if self.vr:
            self.vr.LookupParameter('IGF_X_Einbauort').Set(self.nummer)
        if self.rr:
            self.rr.LookupParameter('IGF_X_Einbauort').Set(self.nummer)

    def wert_schreiben_Teil(self):
        if self.vr:
            a = self.vr.Space[doc.GetElement(self.vr.CreatedPhaseId)]
            if  a != None:
                self.vr.LookupParameter('IGF_X_Einbauort').Set(a.Number)
            else:
                self.vr.LookupParameter('IGF_X_Einbauort').Set(self.nummer)
          

        if self.rr:
            a = self.rr.Space[doc.GetElement(self.rr.CreatedPhaseId)]
            if  a != None:
                self.rr.LookupParameter('IGF_X_Einbauort').Set(a.Number)
            else:
                self.rr.LookupParameter('IGF_X_Einbauort').Set(self.nummer)
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
            for n,bauteile in enumerate(nichtbearbeitet):
                systeme_title = bauteile + ' --- ' + str(n+1) + '/' + str(len(nichtbearbeitet)) + ' Familien'
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
                    bauteil_neu = ME(bauteil)
                    if familienAuswahl.teil.IsChecked:
                        try:bauteil_neu.wert_schreiben_Teil()
                        except Exception as e:print(e)
                    else:
                        try:bauteil_neu.wert_schreiben_Alle()
                        except Exception as e:print(e)

                    pb2.update_progress(n1+1, len(dict_alle[bauteile]))
                bearbeitet.append(bauteile)
                nichtbearbeitet.remove(bauteile)
                
    t.Commit()
main()