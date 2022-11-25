# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_lib import get_value, Element_FamilyFilter
from IGF_log import getlog
from pyrevit import revit, UI, DB,script
from Autodesk.Revit.UI.Selection import ISelectionFilter
from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection




__title__ = "2.04 Daten von MEP in ULK schreiben"
__doc__ = """Kühlleistung von MEP-Räume in ULK schreiben

-------------------------
HK Category: HLS-Bauteile
-------------------------
Parameter in HK: manuell definiert, Exemplar Parameter
-------------------------
Parameter in MEP-Räume:
-------------------------
LIN_BA_CALCULATED_HEATING_LOAD: Heizlast Gesamt

IGF_H_HeizlastZuluft: Heizlast Zuluft

Heizlast für HK: LIN_BA_CALCULATED_HEATING_LOAD - IGF_H_HeizlastZuluft


!!!!!!!!!!!!!!!!!!!!!!!!
gleichmäßig: HK_KL = gesamt HK_HL / Anzahl

übernehmen: erst HL eingeben, dann in HK schreiben. 

"""

__authors__ = "Menghui Zhang"

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'ULK-Bauteile')
logger = script.get_logger()

try:
    getlog(__title__)
except:
    pass

class MEPRaumFilter(ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2003600':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

def PickMEP():
    try:
        re = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,MEPRaumFilter())
        elem = doc.GetElement(re)
        return elem
    except:
        pass

# HLS aus aktueller Projekt
HLS_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
HLSIds = HLS_collector.ToElementIds()


AllHLSNamen = []
AllHLSParam = []
for el in HLS_collector:
    Name = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
    if Name in AllHLSNamen:
        continue
    AllHLSNamen.append(Name)
for el in HLS_collector:
    params = el.Parameters
    for e in params:
        Name = e.Definition.Name
        if Name in AllHLSParam:
            continue
        AllHLSParam.append(Name)
    break

HLS_collector.Dispose()

AllHLSNamen.sort()
AllHLSParam.sort()

class K_Bauteile(object):
    def __init__(self,famname,param):
        self.Famname = famname
        self.KLeistung = param

class Familienauswahl(WPFWindow):
    def __init__(self, xaml_file_name):
        self.ListeK = ObservableCollection[K_Bauteile]()
        self.read_config()
        WPFWindow.__init__(self, xaml_file_name)
        try:
            self.dataGrid.ItemsSource = self.ListeK
            self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
            self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        except:
            pass
    
    def read_config(self):
        try:
            for el in config.KFamilie:
                if len(config.KFamilie) == 0:
                    break
                self.ListeK.Add(K_Bauteile(el[0],el[1]))
            
        except Exception as e:
            pass
    
    def write_config(self):
        try:
            liste = []
            for el in self.ListeK:
                liste.append([el.Famname,el.KLeistung])
            config.KFamilie = liste
        except:
            pass
        script.save_config()

    def cancel(self,sender,args):
        self.Close()
        script.exit()

    def OK(self,sender,args):
        self.Close()
        self.write_config()

    def Add(self,sender,args):
        temp_class = K_Bauteile('','')
        self.dataGrid.ItemsSource.Add(temp_class)
        self.dataGrid.Items.Refresh()
    def dele(self,sender,args):
        items = self.dataGrid.SelectedItem
        self.ListeK.Remove(items)

FamilienDialog = Familienauswahl('auswahl.xaml')
try:
    FamilienDialog.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienDialog.Close()
    script.exit()

dict_Bauteile = {}
for el in FamilienDialog.ListeK:
    dict_Bauteile[el.Famname] = el.KLeistung
#////////////////////////////////////////////////
class HeizL(object):
    def __init__(self,elemid,param):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.Param = self.elem.LookupParameter(param)
        self.werte = get_value(self.Param)
        self.Space = self.elem.Space[doc.Phases[0]]
        if self.Space:
            self.Spaceid = self.Space.Id.ToString()
        else:
            self.Spaceid = ''
        self.name = 'ULK'
        self.typ = self.elem.Name

class MEPRaum(object):
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        try:
            self.HL_Gesamt = get_value(self.elem.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD'))
        except:
            self.HL_Gesamt = 0
        try:
            self.HL_Zuluft = get_value(self.elem.LookupParameter('IGF_H_HeizlastZuluft'))
        except:
            self.HL_Zuluft = 0
        self.HL_HK = self.HL_Gesamt - self.HL_Zuluft

dict_Bauteile_MEP = {}
for el in dict_Bauteile.keys():
    coll = Element_FamilyFilter(el)
    for elem in coll:
        heizl = HeizL(elem.Id,dict_Bauteile[el])
        if heizl.Spaceid:
            if heizl.Spaceid not in dict_Bauteile_MEP.keys():
                dict_Bauteile_MEP[heizl.Spaceid] = ObservableCollection[HeizL]()
                dict_Bauteile_MEP[heizl.Spaceid].Add(heizl)
            else:
                dict_Bauteile_MEP[heizl.Spaceid].Add(heizl)

global WEITER
WEITER = False

class Familienbearbeiten(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self.elem = PickMEP()
        self.mepraum = ''
        self.dataGrid.ItemsSource = ObservableCollection[HeizL]()
        self.items = []
        self.set_up()

    def set_up(self):
        self.mepraum = MEPRaum(self.elem.Id)
        self.Nummer.Text = self.mepraum.elem.Number
        self.GesamtHL.Text = str(round(self.mepraum.HL_Gesamt))
        self.ZuluftHL.Text = str(round(self.mepraum.HL_Zuluft))
        self.HKHL.Text = str(round(self.mepraum.HL_HK))
        if self.elem.Id.ToString() in dict_Bauteile_MEP.keys():
            self.items = dict_Bauteile_MEP[self.elem.Id.ToString()]
            i = 0
            for elem in self.items:
                i += 1
                elem.name = 'HZK-' + str(i)
            self.dataGrid.ItemsSource = self.items
        else:
            UI.TaskDialog.Show(__title__,"Keine HZK im ausgewählten Raum gefunden! Bitte wählen Sie ein neu Raum aus.")
            self.elem = PickMEP()
            self.set_up()

    def close(self,sender,args):
        self.Close()

    def gleich(self,sender,args):
        t = DB.Transaction(doc,self.mepraum.elem.Number)
        t.Start()
        for el in self.items:
            el.werte = round(float(self.HKHL.Text) / self.items.Count)
            el.Param.SetValueString(str(el.werte))
        doc.Regenerate()
        t.Commit()
        t.Dispose()
        self.dataGrid.Items.Refresh()
    
    def weiter(self,sender,args):
        global WEITER
        WEITER = True
        self.Close()
        # self.elem = PickMEP()
        # self.set_up()
        # self.show_dialog()

    def overwrite(self,sender,args):
        self.dataGrid.Items.Refresh()
        t = DB.Transaction(doc,self.mepraum.elem.Number)
        t.Start()
        for el in self.items:
            el.Param.SetValueString(str(el.werte))
        doc.Regenerate()
        t.Commit()
        t.Dispose()

Bauteilebearbeiten = Familienbearbeiten('final_bearbeiten.xaml')
try:
    Bauteilebearbeiten.ShowDialog()
except Exception as e:
    logger.error(e)
    Bauteilebearbeiten.Close()
    script.exit()

while(WEITER):
    WEITER = False
    Bauteilebearbeiten = Familienbearbeiten('final_bearbeiten.xaml')
    try:
        Bauteilebearbeiten.ShowDialog()
    except Exception as e:
        logger.error(e)
        Bauteilebearbeiten.Close()
        script.exit()
