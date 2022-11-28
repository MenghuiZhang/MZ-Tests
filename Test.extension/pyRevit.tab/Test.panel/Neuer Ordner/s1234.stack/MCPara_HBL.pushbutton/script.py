# coding: utf8
from rpw import revit, UI, DB
from pyrevit import script, forms
from clr import GetClrType
from eventhandler import \
    FAMILIEN,\
    ITEMTEMPLATE,\
    PARAMETER,\
    EINGABE,\
    FAMILIEBEARBEITEN,\
    CheckBox, \
    GridViewColumn,\
    Binding,\
    DataTemplate,\
    FrameworkElementFactory,\
    ExternalEvent,\
    IS_Familien,\
    ObservableCollection,\
    TextBlock,\
    TextWrapping
            

__title__ = "Berechnungstypen zuweisen"
__doc__ = """

nicht für Revit 2021

"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
app = revit.app
uiapp = revit.uiapp

# {Familie:[Parameter],{Value(typname______111111),NeuFamTyp}}
config = script.get_config(doc.ProjectInformation.Name+\
    doc.ProjectInformation.Number+\
        'Berechnungstypen zuweisen')

a = []
a.index

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):                    
        forms.WPFWindow.__init__(self,'window.xaml')
        self.gridview = self.Datensatz.View
        self.Class_FAMILIEN = FAMILIEN
        self.Class_ITEMTEMPLATE = ITEMTEMPLATE
        self.Class_PARAMETER = PARAMETER
        self.Class_EINGABE = EINGABE
        self.familieb = FAMILIEBEARBEITEN()
        self.Class_CheckBox = CheckBox
        self.Class_GridViewColumn = GridViewColumn
        self.Class_Binding = Binding
        self.Class_DataTemplate = DataTemplate
        self.Class_FrameworkElementFactory = FrameworkElementFactory
        self.Class_TextBlock = TextBlock
        self.m_GetClrType = GetClrType
        self.wrap = TextWrapping.Wrap
        self.script = script
        
        self.config_datei = config
        self.IS_Familien = IS_Familien
        self.Familien.ItemsSource = self.IS_Familien
        self.liste = ObservableCollection[self.Class_ITEMTEMPLATE]()
        self.liste_temp = ObservableCollection[self.Class_ITEMTEMPLATE]()

        self.familiebevent = ExternalEvent.Create(self.familieb)

        # self.read_config0()
        self.read_config()

    def get_liste(self):
        alles_neu = False
        params = []
        temp_dict = {}
        
        for el in self.Familien.SelectedItem.IS_Params:
            if el.checked:
               params.append(el.name)
        if len(self.Familien.SelectedItem.Params) == len(params):
            for el in self.Familien.SelectedItem.Params:
                if el not in params:
                    alles_neu = True
                    break
        else:
            alles_neu = True
        
        if alles_neu == True:
            self.Familien.SelectedItem.Params = params
            self.Familien.SelectedItem.verteilung.Clear()
            if len(params) > 0:
                for el in self.Familien.SelectedItem.elems:
                    value = el.Name +'_______'
                    for para in params:
                        wert = el.LookupParameter(para).AsInteger()
                        value += str(wert)
                    if value not in temp_dict.keys():
                        temp_dict[value] = []
                    temp_dict[value].append(el)

                for el in sorted(temp_dict.keys()):
                    liste_temp = el.split('_______')
                    name = liste_temp[0]
                    value = liste_temp[1]
                    temp_item = self.Class_ITEMTEMPLATE(name,value,params,temp_dict[el],self.Familien.SelectedItem.dict_Families)
                    self.Familien.SelectedItem.verteilung.Add(temp_item)
            else:
                for el in self.Familien.SelectedItem.elems:
                    value = el.Name
                    if value not in temp_dict.keys():
                        temp_dict[value] = []
                        temp_dict[value].append(el)
                for name in sorted(temp_dict.keys()):
                    temp_item = self.Class_ITEMTEMPLATE(name,'',[],temp_dict[name],self.Familien.SelectedItem.dict_Families)
                    self.Familien.SelectedItem.verteilung.Add(temp_item)
            self.Familien.SelectedItem.Values = {}
            for el in temp_dict.keys():
                self.Familien.SelectedItem.Values[el] = ''

    def get_liste_default(self,Familie):
        temp_dict = {}
        Familie.verteilung.Clear()
        if len(Familie.Params) > 0:
            for el in Familie.IS_Params:
                if el.name in Familie.Params:
                    el.checked = True

            for el in Familie.elems:
                value = el.Name +'_______'
                for para in Familie.Params:
                    wert = el.LookupParameter(para).AsInteger()
                    value += str(wert)
                if value not in temp_dict.keys():
                    temp_dict[value] = []
                temp_dict[value].append(el)

            for el in sorted(temp_dict.keys()):
                liste_temp = el.split('_______')
                name = liste_temp[0]
                value = liste_temp[1]
                temp_item = self.Class_ITEMTEMPLATE(name,value,Familie.Params,temp_dict[el],Familie.dict_Families)
                Familie.verteilung.Add(temp_item)
                if len(Familie.Values.keys()) > 0:
                    if el in Familie.Values.keys():
                        try:
                            temp_item.elementtypeindex = temp_item.ListElementType.IndexOf(Familie.Values[el])
                        except:
                            temp_item.elementtypeindex = -1
        else:
            for el in Familie.elems:
                value = el.Name
                if value not in temp_dict.keys():
                    temp_dict[value] = []
                temp_dict[value].append(el)
            for name in sorted(temp_dict.keys()):
                temp_item = self.Class_ITEMTEMPLATE(name,'',[],temp_dict[name],Familie.dict_Families)
                Familie.verteilung.Add(temp_item)
                if len(Familie.Values.keys()) > 0:
                    if name in Familie.Values.keys():
                        try:
                            temp_item.elementtypeindex = temp_item.ListElementType.IndexOf(Familie.Values[name])
                        except:
                            temp_item.elementtypeindex = -1

    def set_up(self):
        while self.gridview.Columns.Count > 3:
            self.gridview.Columns.RemoveAt(3)
        for el in self.Familien.SelectedItem.Params:
            self.set_up_listview(el)

    def set_up_listview(self,Header):
        GridViewColumn_temp = self.Class_GridViewColumn()
        GridViewColumn_temp.Header = self.Class_TextBlock()
        GridViewColumn_temp.Width = 100
        GridViewColumn_temp.Header.Text = Header
        GridViewColumn_temp.Header.TextWrapping = self.wrap
        factory = self.Class_FrameworkElementFactory(self.m_GetClrType(self.Class_CheckBox))
        factory.SetBinding(self.Class_CheckBox.IsCheckedProperty,self.Class_Binding(Header))
        factory.SetValue(self.Class_CheckBox.IsEnabledProperty,False)
        CellTemplate_temp = self.Class_DataTemplate()
        CellTemplate_temp.VisualTree = factory
        GridViewColumn_temp.CellTemplate = CellTemplate_temp
        self.gridview.Columns.Add(GridViewColumn_temp)
  
    def familiechanged(self, sender, args):
        try:
            self.parameterbbbb.IsEnabled = True
            self.set_up()
            self.Datensatz.ItemsSource = self.Familien.SelectedItem.verteilung
        except Exception as e:
            print(e)
    
    # def read_config0(self):
    #     Familiendatei = {}
    #     try:
    #         Familiendatei = self.config_datei.Daten
    #         self.detailsDatei = Familiendatei[len(Familiendatei.keys())-1]
    #     except Exception as e:print(e)

    def read_config(self):
        Familiendatei = {}
        try:
            Familiendatei = self.config_datei.datei
        except Exception as e:
            print(e)
            
        try:
            for el in self.IS_Familien:
                if el.name in Familiendatei.keys():
                    try:
                        el.Params = Familiendatei[el.name][0]
                    except Exception as e:print(e)
                    try:
                        el.Values = Familiendatei[el.name][1]
                    except Exception as e:print(e)
                    try:
                        self.get_liste_default(el)
                    except Exception as e:print(e)
                else:self.get_liste_default(el)
        except Exception as e:
            print(e)

    # def write_config0(self):
    #     try:
    #         datei = {}
    #         for el in self.IS_Familien:
    #             datei[el.name] = [[],{}]
    #             datei[el.name][0] = sorted(el.Params)
    #             datei[el.name][1] = el.Values
    #         self.detailsDatei = datei
    #     except:print('Fehler','write_fonfig')
    
    def write_config(self):
        try:
            datei = {}
            try:
                for el in self.IS_Familien:
                    if el.name in self.config_datei.datei.keys():
                        temp_Params = self.config_datei.datei[el.name][0]
                        temp_Values = self.config_datei.datei[el.name][1]
                        neu_param = False
                        if len(temp_Params) == len(el.Params):
                            for para in temp_Params:
                                if para not in el.Params:
                                    neu_param = True
                                    break
                        else:
                            neu_param = True
                        if neu_param ==  True:
                            datei[el.name] = [[],{}]
                            datei[el.name][0] = sorted(el.Params)
                            datei[el.name][1] = el.Values
                        else:
                            datei[el.name] = [[],{}]
                            datei[el.name][0] = sorted(el.Params)
                            datei[el.name][1] = temp_Values
                            for Key in el.Values.keys():
                                datei[el.name][1][Key] = el.Values[Key]
                        
                self.config_datei.datei = datei
                self.script.save_config()
            except Exception as e:
                print(e)
        except:print('Fehler','write_fonfig')

    def start(self, sender, args):
        self.familiebevent.Raise()

    def ParameterArb(self, sender, args):
        try:
            wpfEingabe = self.Class_EINGABE()
            wpfEingabe.IS = self.Familien.SelectedItem.IS_Params
            wpfEingabe.Params.ItemsSource = wpfEingabe.IS
            wpfEingabe.ShowDialog()
            self.get_liste()
            self.set_up()
            self.Datensatz.ItemsSource = self.Familien.SelectedItem.verteilung
        except Exception as e:print(e)

    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        print(sender.DataContext,args.Source)
        if self.Datensatz.SelectedItem is not None:
            try:
                if sender.DataContext in self.Datensatz.SelectedItems:
                    for item in self.Datensatz.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.Datensatz.Items.Refresh()
                else:
                    pass
            except:
                pass

    def berechnungstypenchanged(self, sender, args):
        item = sender.DataContext
        try:
            self.Familien.SelectedItem.Values[item.defikey] = item.ListElementType[item.elementtypeindex]
        except:pass

wind = AktuelleBerechnung()
wind.familieb.GUI = wind
wind.Show()