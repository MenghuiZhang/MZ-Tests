# coding: utf8
from pyrevit import revit,DB
from pyrevit import script
from Autodesk.Revit.DB import *
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow

__title__ = "Plankopf ändern"
__doc__ = """Plankopf ändern"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc


planids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElementIds()

if not planids:
    logger.error('Keine Pläne in Projekt')
    script.exit()

plankopf_liste = []
plankopfid_dict = {}

# Planköpfe aus aktueller Projekt
TitleBlocks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
TitleBlocks = TitleBlocks_collector.ToElementIds()
TitleBlocks_collector.Dispose()

for el in TitleBlocks:
    elem = doc.GetElement(el)
    familyandtyp = elem.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
    plankopfid_dict[familyandtyp] = el
    plankopf_liste.append(familyandtyp)
        
plankopf_liste.sort()

class Plaene(object):
    def __init__(self):
        self.checked = False
        self.plannummer = ''
        self.altfamily = ''
        self.alttyp = ''
        self.neufamilyandtyp = ''
        self.plankopfid = ''
        self.Group = ''
        self.Mcate = ''
        self.Subcate = ''

    @property
    def Group(self):
        return self._Group
    @Group.setter
    def Group(self, value):
        self._Group = value
    @property
    def Mcate(self):
        return self._Mcate
    @Mcate.setter
    def Mcate(self, value):
        self._Mcate = value
    @property
    def Subcate(self):
        return self._Subcate
    @Subcate.setter
    def Subcate(self, value):
        self._Subcate = value
    @property
    def plannummer(self):
        return self._plannummer
    @plannummer.setter
    def plannummer(self, value):
        self._plannummer = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def altfamily(self):
        return self._altfamily
    @altfamily.setter
    def altfamily(self, value):
        self._altfamily = value
    @property
    def alttyp(self):
        return self._alttyp
    @alttyp.setter
    def alttyp(self, value):
        self._alttyp = value
    @property
    def neufamilyandtyp(self):
        return self._neufamilyandtyp
    @neufamilyandtyp.setter
    def neufamilyandtyp(self, value):
        self._neufamilyandtyp = value
    
    @property
    def plankopfid(self):
        return self._plankopfid
    @plankopfid.setter
    def plankopfid(self, value):
        self._plankopfid = value
    
    @property
    def planid(self):
        return self._planid
    @planid.setter
    def planid(self, value):
        self._planid = value

Liste_Plaene = ObservableCollection[Plaene]()

Filterplankopf = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_TitleBlocks)
Auswahldict = {}

for planid in planids:
    elem = doc.GetElement(planid)
    plankopf = elem.GetDependentElements(Filterplankopf)
    if plankopf.Count != 1:
       continue
    tempclass = Plaene()
    plannr = elem.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
    tempclass.plannummer = plannr
    tempclass.planid = planid
    for kopfid in plankopf:
        kopf = doc.GetElement(kopfid)
        tempclass.plankopfid = kopfid
        try:
            family = kopf.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
            typ = kopf.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            tempclass.altfamily = family
            tempclass.alttyp = typ
        except Exception as e:
            logger.error(e) 
        try:
            group = elem.LookupParameter('Sheet Group').AsString()
            if not group:
                group = '???'
            tempclass.Group = group
            mcate = elem.LookupParameter('Main Category').AsString()
            if not mcate:
                mcate = '???'
            tempclass.Mcate = mcate
            scate = elem.LookupParameter('Subcategory 1').AsString()
            if not scate:
                scate = '???'
            tempclass.Subcate = scate
            if not group in Auswahldict.Keys:
                Auswahldict[group] = {}
            if not mcate in Auswahldict[group].Keys:
                Auswahldict[group][mcate] = []
            if not scate in Auswahldict[group][mcate]:
                Auswahldict[group][mcate].append(scate)
        except Exception as e:
            logger.error(e) 
    Liste_Plaene.Add(tempclass)

keys1 = Auswahldict.Keys[:]
keys1.sort()
keys1.append('Keine')

keys2 = []
keys3 = []

# GUI Pläne
class PlaeneUI(WPFWindow):
    def __init__(self, xaml_file_name,liste_plaene):
        self.list_plaene = liste_plaene
        WPFWindow.__init__(self, xaml_file_name)
        self.dataGrid.ItemsSource = liste_plaene
        self.tempcoll = ObservableCollection[Plaene]()
        self.leercoll = ObservableCollection[Plaene]()
        self.dataGrid.Columns[7].ItemsSource = plankopf_liste
        self.dataGrid.Columns[1].ItemsSource = keys1
        self.altdatagrid = liste_plaene
        self.group.ItemsSource = keys1
        self.main.ItemsSource = keys2
        self.sub.ItemsSource = keys3
        self.prevalue1 = ''
        self.newvalue1= ''

        self.prevalue2 = ''
        self.newvalue2 = ''

        self.prevalue3 = ''
        self.newvalue3 = ''

        self.prevalue7 = ''
        self.newvalue7 = ''

        self.plansuche.TextChanged += self.auswahl_txt_changed
        self.group.SelectionChanged += self.auswahl_txt_changed
        self.group.SelectionChanged += self.auswahl_group_changed
        self.sub.SelectionChanged += self.auswahl_txt_changed
        self.main.SelectionChanged += self.auswahl_txt_changed
        self.main.SelectionChanged += self.auswahl_main_changed
    def auswahl_group_changed(self,sender,args):
        group = ''
        temp_key2 = []
        temp_key3 = []
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            temp_key2 = Auswahldict[group].Keys[:]
            temp_key2.sort()
            temp_key2.append('Keine')
            
            self.main.ItemsSource = temp_key2
            self.main.Text = 'Keine'
            self.sub.ItemsSource = temp_key3
        except:
            self.main.ItemsSource = temp_key2
            self.main.Text = ''
            self.sub.ItemsSource = temp_key3
            self.sub.Text = ''
    def auswahl_main_changed(self,sender,args):
        main = ''
        try:
            main = self.main.SelectedItem.ToString()
        except:
            pass
        group = ''
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            keys3 = Auswahldict[group][main][:]
            keys3.sort()
            keys3.append('Keine')
            
            self.sub.Text = ''
            self.sub.ItemsSource = keys3
        except:
            self.sub.Text = ''
            self.sub.ItemsSource = []

    def dataGrid_BeginningEdit(self,sender,args):
        if args.Column.DisplayIndex == 1:
            self.prevalue1 = args.Column.GetCellContent(args.Row).Text
        elif args.Column.DisplayIndex == 2:
            self.prevalue2 = args.Column.GetCellContent(args.Row).Text
        elif args.Column.DisplayIndex == 3:
            self.prevalue3 = args.Column.GetCellContent(args.Row).Text
        elif args.Column.DisplayIndex == 7:
            self.prevalue7 = args.Column.GetCellContent(args.Row).Text

    def dataGrid_CellEditEnding(self,sender,args):
        if args.Column.DisplayIndex == 0:
            value = args.Row.Item.checked
            args.Row.Item.checked = not value
        elif args.Column.DisplayIndex == 1:
            self.newvalue1 = args.EditingElement.Text
            if self.newvalue1 != self.prevalue1:
                args.Row.Item.checked = True
                args.Row.Item.Group = self.newvalue1
                self.prevalue1 = ''
        elif args.Column.DisplayIndex == 2:
            self.newvalue2 = args.EditingElement.Text
            if self.newvalue2 != self.prevalue2:
                args.Row.Item.checked = True
                args.Row.Item.Mcate = self.newvalue2
                self.prevalue2 = ''
                if not args.Row.Item.Mcate in Auswahldict[args.Row.Item.Group].Keys:
                    Auswahldict[args.Row.Item.Group][args.Row.Item.Mcate] = [args.Row.Item.Subcate]
        elif args.Column.DisplayIndex == 3:
            self.newvalue3 = args.EditingElement.Text
            if self.newvalue3 != self.prevalue3:
                args.Row.Item.checked = True
                args.Row.Item.Subcate = self.newvalue3
                self.prevalue3 = ''
                if not args.Row.Item.Subcate in Auswahldict[args.Row.Item.Group][args.Row.Item.Subcate]:
                    Auswahldict[args.Row.Item.Group][args.Row.Item.Mcate].append(args.Row.Item.Subcate)
        elif args.Column.DisplayIndex == 7:
            self.newvalue7 = args.EditingElement.Text
            if self.newvalue7 != self.prevalue7:
                args.Row.Item.checked = True
                args.Row.Item.neufamilyandtyp = self.newvalue7
                self.prevalue7 = ''
            for item in self.dataGrid.Items:
                if item.checked:
                    if item.neufamilyandtyp in ['', None]:
                        item.neufamilyandtyp = self.newvalue7

        self.dataGrid.ItemsSource = self.leercoll
        if self.tempcoll.Count >= 1:
            self.dataGrid.ItemsSource = self.tempcoll
        else:
            self.dataGrid.ItemsSource = self.list_plaene
        
    def auswahl_txt_changed(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.plansuche.Text
        group = ''
        sub = ''
        main = ''
        try:
            group = self.group.SelectedItem.ToString()
        except:
            pass
        try:
            sub = self.sub.SelectedItem.ToString()
        except:
            pass
        try:
            main = self.main.SelectedItem.ToString()
        except:
            pass
        
        if not group:
            group = self.group.Text = ''
        if not sub:
            sub = self.sub.Text = ''
        if not main:
            main = self.main.Text = ''
        
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid
            text_typ = self.plansuche.Text = ''

        for item in self.altdatagrid:
            if item.plannummer.find(text_typ) != -1:
                if group in ['','Keine']:
                    self.tempcoll.Add(item)
                else:
                    if item.Group == group:
                        if main in ['','Keine']:
                            self.tempcoll.Add(item)
                        else:
                            if main == item.Mcate:
                                if sub in ['','Keine']:
                                    self.tempcoll.Add(item)
                                else:
                                    if sub == item.Subcate:
                                        self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = True
        self.dataGrid.Items.Refresh()


    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = False
        self.dataGrid.Items.Refresh()


    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.checked
            item.checked = not value
        self.dataGrid.Items.Refresh()

    def aktu(self,sender,args):
        self.Close()
        tran = DB.Transaction(doc,'Pläne anpassen')
        tran.Start()
        for item in self.list_plaene:
            if item.checked == True:
                plankopf = doc.GetElement(item.plankopfid)
                if not item.neufamilyandtyp in ['', None]:
                    neuid = plankopfid_dict[item.neufamilyandtyp]
                    plankopf.ChangeTypeId(neuid)
                plan = doc.GetElement(item.planid)
                plan.LookupParameter('Sheet Group').Set(item.Group)
                plan.LookupParameter('Main Category').Set(item.Mcate)
                plan.LookupParameter('Subcategory 1').Set(item.Subcate)
        tran.Commit()
        
    def close(self,sender,args):
        self.Close()

Planfenster = PlaeneUI("window.xaml",Liste_Plaene)
try:Planfenster.ShowDialog()
except Exception as e:
    logger.error(e)
    Planfenster.Close()
    script.exit()