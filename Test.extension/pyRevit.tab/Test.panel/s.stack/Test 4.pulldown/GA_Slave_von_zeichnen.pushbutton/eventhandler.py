# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter,ObjectType 
import Autodesk.Revit.DB as DB
from pyrevit import script, forms, revit
from System.Windows import Visibility 
from System.Collections.ObjectModel import ObservableCollection

class Filter(ISelectionFilter):
    def __init__(self,Familien_liste = []):
        self.Familien_liste = Familien_liste
    
    def AllowElement(self,elem):
        try:
            if elem.Symbol.FamilyName in self.Familien_liste:
                return True
            else:
                return False
        except:return False
    def AllowReference(self,ref,p):
        return False

class FilterMEP(ISelectionFilter):
   
    def AllowElement(self,elem):
        try:
            if elem.Category.Name == 'MEP-Räume':
                return True
            else:
                return False
        except:return False
    def AllowReference(self,ref,p):
        return False	
	
class FilterBSK(ISelectionFilter):
   
    def AllowElement(self,elem):
        try:
            if elem.Category.Name == 'Luftkanalzubehör' or elem.Category.Name == 'Luftdurchlässe':
                return True
            else:
                return False
        except:return False
    def AllowReference(self,ref,p):
        return False	

uidoc = revit.uidoc
doc = revit.doc
info = doc.ProjectInformation.Number+'_'+doc.ProjectInformation.Name
config = script.get_config(info+'GASlaveVon')

# erstellt ein Familie Collector
Familie_coll = DB.FilteredElementCollector(doc).OfClass(DB.Family)
Familieids = Familie_coll.ToElementIds()
Familie_coll.Dispose()

Liste_Kate = ['Elektrische Ausstattung','HLS-Bauteile','Luftdurchlässe',
'Luftkanalzubehör','Rohrzubehör','Sanitärinstallationen']

Dict_Category_Familien = {}


class Familie(object):
    def __init__(self,name,Cate,checked = False):
        self.checked = checked
        self.Name = name
        self.Cate = Cate

class Kategorie(object):
    def __init__(self,Name):
        self.Name = Name
        self.Familien = ObservableCollection[Familie]()

Liste_Familie = ObservableCollection[Familie]()
for elid in Familieids:
    elem = doc.GetElement(elid)
    Dict_Cate_Fam = {}
    if elem.FamilyCategory.Name in Liste_Kate:
        famname = elem.Name  
        if elem.FamilyCategory.Name not in Dict_Cate_Fam.keys():
            Dict_Cate_Fam[elem.FamilyCategory.Name] = []
        
        Dict_Cate_Fam[elem.FamilyCategory.Name].append(famname)
    for cate in Dict_Cate_Fam.keys():
        fam_liste = sorted(Dict_Cate_Fam[cate])
        for fam in fam_liste:
            temp = Familie(fam,cate)
            Liste_Familie.Add(temp)
            if temp.Cate not in Dict_Category_Familien.keys():
                Dict_Category_Familien[temp.Cate] = []
            Dict_Category_Familien[temp.Cate].append(temp)

Liste_Kategorie = ObservableCollection[Kategorie]()
for el in Liste_Kate:
    temp = Kategorie(el)
    if el in Dict_Category_Familien.keys():
        temp.Familien = ObservableCollection[Familie](Dict_Category_Familien[el])
    Liste_Kategorie.Add(temp)

     
class FamilieUI(forms.WPFWindow):
    def __init__(self,Liste_Kategorie):
        self.Liste_Kategorie = Liste_Kategorie
        self.tempcoll = ObservableCollection[Familie]()
        forms.WPFWindow.__init__(self, "Familie.xaml")
        self.List_View_Cate.ItemsSource = self.Liste_Kategorie
        self.List_View_Cate.SelectedIndex = 0
        self.ListView_Familie.ItemsSource = self.List_View_Cate.SelectedItem.Familien
     
    def catechanged(self,sender,args):
        if self.List_View_Cate.SelectedIndex != -1:
            self.ListView_Familie.ItemsSource = self.List_View_Cate.SelectedItem.Familien
        
    def checkedchanged(self,sender,args):
        """bearbeitet checked changed""" 
        try:
            Checked = sender.IsChecked
            if self.ListView_Familie.SelectedItem is not None:
                if sender.DataContext in self.ListView_Familie.SelectedItems:
                    try:
                        for item in self.ListView_Familie.SelectedItems:
                            item.checked = Checked
                    except:
                        pass
            self.ListView_Familie.Items.Refresh()
        except:
            pass
    
    def search_txt_changed(self, sender, args):
        """bearbeitet Textänderungen im Suchfeld.""" 
        self.tempcoll.Clear()
        text_Fam = self.such.Text.upper()
        for item in self.List_View_Cate.SelectedItem.Familien:
            if item.Name.upper().find(text_Fam) != -1:
                self.tempcoll.Add(item)

        self.ListView_Familie.ItemsSource = self.tempcoll
        self.ListView_Familie.Items.Refresh()

    def ok(self,sender,args):
        self.Close()
    
    def keine(self,sender,args):
        for class_cate in self.Liste_Kategorie:
            for familie in class_cate.Familien:
                familie.checked = False
        self.ListView_Familie.Items.Refresh()

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        self.GUI.mepitem.IsEnabled = False
        self.GUI.bskitem.IsEnabled = False
        self.GUI.buttonfamsel.IsEnabled = False
        try:

            doc = uidoc.Document
            try:
                Liste = self.GUI.config.Liste_Familien
            except:
                Liste = []
            filters = Filter(Liste)
            elems = uidoc.Selection.PickElementsByRectangle(filters,'Familien auswählen') 
            if elems.Count == 0:
                TaskDialog.Show('Info.','Keinen Bauteil ausgewählt')
                self.GUI.button_select.IsEnabled = False
                self.GUI.mepitem.IsEnabled = True
                self.GUI.bskitem.IsEnabled = True
                self.GUI.buttonfamsel.IsEnabled = True
                return 
            self.GUI.button_select.IsEnabled = True
            self.GUI.elems = elems
            elemids = uidoc.Selection.GetElementIds()
            elemids.Clear()
            for el in elems:
                elemids.Add(el.Id)
            uidoc.Selection.SetElementIds(elemids)
            self.GUI.mepitem.IsEnabled = True
            self.GUI.bskitem.IsEnabled = True
            self.GUI.buttonfamsel.IsEnabled = True
            return
        except:
            self.GUI.mepitem.IsEnabled = True
            self.GUI.bskitem.IsEnabled = True
            self.GUI.buttonfamsel.IsEnabled = True

    def GetName(self):
        return "Bauteile auswählen"

class SCHREIBEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        text = self.GUI.textboxSlavevon.Text
        if text in [None,'']:
            TaskDialog.Show('Fehler.','den Wert von Slave von nicht eingegeben')
            return 
  
        elems = self.GUI.elems
        
        t = DB.Transaction(doc,'Slave von schreiebn')
        t.Start()
        for el in elems:
            try:
                el.LookupParameter('IGF_GA_Slave_von').Set(text)
            except:
                print(el.Id.ToString())
        t.Commit()
        self.GUI.button_select.IsEnabled = False

    def GetName(self):
        return "Slave von schreiben"

class SELECTMEP(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        self.GUI.slavevonitem.IsEnabled = False
        self.GUI.bskitem.IsEnabled = False

        try:
            doc = uidoc.Document
            filters = FilterMEP()
            elems = uidoc.Selection.PickElementsByRectangle(filters,'MEP-Räume auswählen') 
            if elems.Count == 0:
                TaskDialog.Show('Info.','Kein MEP Raum ausgewählt')
                self.GUI.buttonwritemep.IsEnabled = False
                self.GUI.slavevonitem.IsEnabled = True
                self.GUI.bskitem.IsEnabled = True
                return 
            self.GUI.buttonwritemep.IsEnabled = True
            self.GUI.elems = elems
            elemids = uidoc.Selection.GetElementIds()
            elemids.Clear()
            for el in elems:
                elemids.Add(el.Id)
            uidoc.Selection.SetElementIds(elemids)
            self.GUI.slavevonitem.IsEnabled = True
            self.GUI.bskitem.IsEnabled = True
            return
        except:
            self.GUI.slavevonitem.IsEnabled = True
            self.GUI.bskitem.IsEnabled = True


    def GetName(self):
        return "MEP-Räume auswählen"

class SCHREIBENMEP(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        text = self.GUI.textbox_bb.Text
        if text in [None,'']:
            TaskDialog.Show('Fehler.','den Wert von Brandbereich nicht eingegeben')
            return 
  
        elems = self.GUI.elems
        
        t = DB.Transaction(doc,'Brandbereich schreiebn')
        t.Start()
        for el in elems:
            try:
                el.LookupParameter('IGF_GA_Brandbereich').Set(text)
            except:
                print(el.Id.ToString())
        t.Commit()
        self.GUI.buttonwritemep.IsEnabled = False

    def GetName(self):
        return "Brandbereich schreiben"

class SCHREIBENBSK(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
    def Execute(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        try:
            self.GUI.slavevonitem.IsEnabled = False
            self.GUI.mepitem.IsEnabled = False

            
            while True:
                bsk = uidoc.Selection.PickObject(ObjectType.Element,FilterBSK(),'BSK auswählen')
                mep1 = uidoc.Selection.PickObject(ObjectType.Element,FilterMEP(),'MEP Raum 1 auswählen')
                mep2 = uidoc.Selection.PickObject(ObjectType.Element,'MEP Raum 2 auswählen')
                bb1 = doc.GetElement(mep1).LookupParameter('IGF_GA_Brandbereich').AsString()
                bb2 = doc.GetElement(mep2).LookupParameter('IGF_GA_Brandbereich').AsString()
                t = DB.Transaction(doc,'BSK Brandbereich')
                t.Start()
                try:
                    doc.GetElement(bsk).LookupParameter('IGF_GA_Brandbereich_1_Name').Set(bb1)
                    doc.GetElement(bsk).LookupParameter('IGF_GA_Brandbereich_2_Name').Set(bb2)
                except Exception as e:
                    t.RollBack()
                    print('Bitte überprüfen')
                    break
                t.Commit()
                t.Dispose()
        except Exception as e:

            self.GUI.slavevonitem.IsEnabled = True
            self.GUI.mepitem.IsEnabled = True

            pass
            

    def GetName(self):
        return "Brandbereich schreiben"
