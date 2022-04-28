# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw,clr
clr.AddReference("PresentationFramework")
import time
from System.Windows.Controls import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ISelectionFilter
from pyrevit.forms import WPFWindow, SelectFromList
import System
import System.Windows
from System.Windows import Application, Window
from System.Collections.ObjectModel import *
import os.path as op
clr.AddReference("RevitServices")
from RevitServices.Transactions import TransactionManager
from System.Windows.Media import Colors, Brushes
from System.Windows import FontWeights, FontStyles



start = time.time()


__title__ = "2.81 Heiz-/Kühllast für DeS"
__doc__ = """Heiz-/Kühlleistung = Heiz-/Kühllast / Menge
Voraussetzung: Nur Deckensegel im Raum. Kein ULK/HZK. """
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
Familien_DeS_config = script.get_config()
from pyIGF_logInfo import getlog
getlog(__title__)

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
phase = list(doc.Phases)[-1]

# HLS aus aktueller Projekt
HLS_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
    .WhereElementIsNotElementType()
HLS = HLS_collector.ToElementIds()

AllHLSNamen = []
AllHLSParam = []
for el in HLS_collector:
    Name = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
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

AllHLSNamen.sort()
AllHLSParam.sort()
class HUK_Bauteile(object):
    def __init__(self):
        self.Element = ''
        self.HLeistungSchreiben = ''
        self.KLeistungSchreiben = ''

    @property
    def Element(self):
        return self._Element
    @Element.setter
    def Element(self, value):
        self._Element = value
    @property
    def HLeistungSchreiben(self):
        return self._HLeistungSchreiben
    @HLeistungSchreiben.setter
    def HLeistungSchreiben(self, value):
        self._HLeistungSchreiben = value
    @property
    def KLeistungSchreiben(self):
        return self._KLeistungSchreiben
    @KLeistungSchreiben.setter
    def KLeistungSchreiben(self, value):
        self._KLeistungSchreiben = value

Liste_HUK_Familie = ObservableCollection[HUK_Bauteile]()
Liste_HUKVerbinder_Familie = ObservableCollection[HUK_Bauteile]()

def ObservableCollection2Liste(Coll):
    out_list = []
    if Coll.Count > 0:
        for el in Coll:
            out_list.append([el.Element,
            el.HLeistungSchreiben,
            el.KLeistungSchreiben
            ])
    return out_list

def Liste2Class(liste):
    out_class = HUK_Bauteile()
    out_class.Element = str(liste[0])
    out_class.HLeistungSchreiben = str(liste[1])
    out_class.KLeistungSchreiben = str(liste[2])
    return out_class

def read_config():
    try:
        for el in Familien_DeS_config.liste_HUK_Familie:
            if len(Familien_DeS_config.liste_HUK_Familie) == 0:
                break
            Liste_HUK_Familie.Add(Liste2Class(el))

    except Exception as e:
        logger.error(e)
    try:
        for el in Familien_DeS_config.liste_HUKVer_Familie:
            if len(Familien_DeS_config.liste_HUKVer_Familie) == 0:
                break
            Liste_HUKVerbinder_Familie.Add(Liste2Class(el))

    except Exception as e:
        logger.error(e)



def write_config():
    Familien_DeS_config.liste_HUK_Familie = ObservableCollection2Liste(Liste_HUK_Familie)
    Familien_DeS_config.liste_HUKVer_Familie = ObservableCollection2Liste(Liste_HUKVerbinder_Familie)
    script.save_config()

read_config()

class FamilienBearbeiten(WPFWindow):
    def __init__(self, xaml_file_name,listeHuK,listeHuKVer):
        self.ListeHuK = listeHuK
        self.ListeHuKVer = listeHuKVer
        WPFWindow.__init__(self, xaml_file_name)
        try:
            #self.all.Background = Brushes.White
            #self.all.FontWeight = FontWeights.Bold
            #self.all.FontStyle = FontStyles.Italic
            self.dataGrid.ItemsSource = listeHuK
            self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
            self.dataGrid.Columns[1].ItemsSource = AllHLSParam
            self.dataGrid.Columns[2].ItemsSource = AllHLSParam
        except Exception as e:
            print(e)
            pass

    def cancel(self,sender,args):
        self.Close()
        self.ListeHuK.Clear()
        self.ListeHuKVer.Clear()
        read_config()


    def OK(self,sender,args):
        self.Close()

    def HundK(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeHuK
        self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
        self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        self.dataGrid.Columns[2].ItemsSource = AllHLSParam

    def HundKVer(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeHuKVer
        self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
        self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        self.dataGrid.Columns[2].ItemsSource = AllHLSParam


    def Add(self,sender,args):
        temp_class = HUK_Bauteile()
        self.dataGrid.ItemsSource.Add(temp_class)
        self.dataGrid.Items.Refresh()


FamilienDialog = FamilienBearbeiten('Test.xaml', Liste_HUK_Familie, Liste_HUKVerbinder_Familie)
FamilienDialog.ShowDialog()

write_config()

DeS_Familien = {}
DeS_Verbinder_Familien = {}

for el in Liste_HUK_Familie:
    DeS_Familien[el.Element] = [el.HLeistungSchreiben,el.KLeistungSchreiben]
for el in Liste_HUKVerbinder_Familie:
    DeS_Verbinder_Familien[el.Element] = [el.HLeistungSchreiben,el.KLeistungSchreiben]


Space_Dict = {}
Space_Collector = []
idListe = []
for el in HLS_collector:
    try:
        FamilieName = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()

        if FamilieName in DeS_Familien.keys():
            space = el.Space[phase]
            if not space:
                logger.error("Id: {}, Kein MEP-Raum".format(el.Id))
                idListe.append(el.Id)
            Nr = space.Number

            if not Nr in Space_Dict.keys():
                Space_Dict[Nr] = [el]
                Space_Collector.append(space)
            else:
                Space_Dict[Nr].append(el)

        elif FamilieName in DeS_Verbinder_Familien.keys():
            space = el.Space[phase]
            if not space:
                logger.error("Id: {}, Kein MEP-Raum".format(el.Id))
                idListe.append(el.Id)
            Nr = space.Number
            if not Nr in Space_Dict.keys():
                Space_Dict[Nr] = [el]
                Space_Collector.append(space)
            else:
                Space_Dict[Nr].append(el)

    except Exception as e:
        logger.error(e)

if any(idListe):
    sel = uidoc.Selection.GetElementIds()
    if forms.alert('Falsche Elements anzeigen?', ok=False, yes=True, no=True):
        for el in idListe:
            sel.Add(el)
            uidoc.Selection.SetElementIds(sel)
            sel.Remove(el)
            uidoc.ShowElements(el)

            if forms.alert('weiter anzeigen?', ok=False, yes=True, no=True):
                pass
            else:
                break

if forms.alert('Daten schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} MEP Räume',
                           cancellable=True, step=10) as pb1:
        n_1 = 0

        t = Transaction(doc, "gleichmäßig")
        t.Start()

        for Spaces in Space_Collector:
            if pb1.cancelled:
                t.RollBack()
                script.exit()
            n_1 += 1
            pb1.update_progress(n_1, len(Space_Collector))

            Nummer = Spaces.Number
            HL = Spaces.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
            KL = Spaces.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').AsValueString()
            hl = HL[:-1]
            kl = KL[:-1]
            logger.info(50*'*')
            logger.info('Nr: {}, Heizlast: {}, Kühllast: {}'.format(Nummer, HL, KL))
            logger.info(50*'-')
            elems = Space_Dict[Nummer]
            Verbinder = []
            for el in elems:
                FA = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if FA in DeS_Verbinder_Familien.keys():
                    Verbinder.append(el)
            al = 0
            Flaeche = 0
            for elemen in elems:
                L = elemen.LookupParameter('Plattenlänge').AsValueString()
                B = elemen.Symbol.LookupParameter('Plattenbreite').AsValueString()
                Flaeche += int(L)*int(B)
            for el in elems:
                if any(Verbinder):
                    nr = 0
                    Flae = 0
                    for eleme in Verbinder:
                        L = eleme.LookupParameter('Plattenlänge').AsValueString()
                        B = eleme.Symbol.LookupParameter('Plattenbreite').AsValueString()
                        Flae += int(L)*int(B)
                    for ele in Verbinder:
                        nr += 1
                        FA = ele.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                        LAE = int(ele.LookupParameter('Plattenlänge').AsValueString())
                        BRE = int(ele.Symbol.LookupParameter('Plattenbreite').AsValueString())
                        Hl = round(int(hl)*LAE*BRE/Flae)
                        Kl = round(int(kl)*LAE*BRE/Flae)
                        ele.LookupParameter(DeS_Verbinder_Familien[FA][0]).SetValueString(str(Hl))
                        ele.LookupParameter(DeS_Verbinder_Familien[FA][1]).SetValueString(str(Kl))
                        logger.info('DeS_Verbinder-{}: Heizleistung: {} W, Kühlleistung: {} W'.format(nr,Hl,Kl))
                    logger.info(50*'*')
                    break
                else:
                    al += 1
                    FA = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                    LAE = int(el.LookupParameter('Plattenlänge').AsValueString())
                    BRE = int(el.Symbol.LookupParameter('Plattenbreite').AsValueString())
                    Hl = round(int(hl)*LAE*BRE/Flaeche)
                    Kl = round(int(kl)*LAE*BRE/Flaeche)
                    el.LookupParameter(DeS_Familien[FA][0]).SetValueString(str(Hl))
                    el.LookupParameter(DeS_Familien[FA][1]).SetValueString(str(Kl))
                    logger.info('DeS-{}: Heizleistung: {} W, Kühlleistung: {} W'.format(al,Hl,Kl))
            if not any(Verbinder):
                logger.info(50*'*')

        t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
