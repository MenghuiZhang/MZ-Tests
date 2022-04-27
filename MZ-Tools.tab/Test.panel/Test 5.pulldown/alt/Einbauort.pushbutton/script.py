# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter

__title__ = "Einbauort, System basiert"
__doc__ = """
Schreibt Einbauort(Raumnummer) in Bauteile ein.

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

system_Luft = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
system_Rohr = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()

DICT_SYSTEM = {'LUFT':{},'ROHR':{}}
for elem in system_Luft:
    typ = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
    if not typ in DICT_SYSTEM['LUFT'].keys():
        DICT_SYSTEM['LUFT'][typ] = [elem]
    else:
        DICT_SYSTEM['LUFT'][typ].append(elem)
for elem in system_Rohr:
    typ = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
    if not typ in DICT_SYSTEM['ROHR'].keys():
        DICT_SYSTEM['ROHR'][typ] = [elem]
    else:
        DICT_SYSTEM['ROHR'][typ].append(elem)

system_Luft.Dispose()
system_Rohr.Dispose()

class MEP_System(object):
    def __init__(self,typ,liste_system):
        self.checked = False
        self.TypName = typ
        self.liste_system = liste_system
        self.liste_bauteile = []
        self.get_elemente()
        
    def get_elemente(self):
        for el in self.liste_system:
            try:
                elemente = el.DuctNetwork
            except:
                elemente = el.PipingNetwork

            for elem in elemente:
                cate = elem.Category.Id.ToString()
                # Leitung
                if cate not in ['-2008000','-2008020','-2008050','-2008044','-2008015','-2008122','-2008123','-2008124','-2008043']:
                    if elem.Id.ToString() not in self.liste_bauteile:
                        self.liste_bauteile.append(elem.Id.ToString())
               


LISTE_ALLE = ObservableCollection[MEP_System]()
LISTE_ROHR = ObservableCollection[MEP_System]()
LISTE_LUFT = ObservableCollection[MEP_System]()

for typ in sorted(DICT_SYSTEM['ROHR'].keys()):
    mepsystem = MEP_System(typ,DICT_SYSTEM['ROHR'][typ])
    LISTE_ALLE.Add(mepsystem)
    LISTE_ROHR.Add(mepsystem)

for typ in sorted(DICT_SYSTEM['LUFT'].keys()):
    mepsystem = MEP_System(typ,DICT_SYSTEM['LUFT'][typ])
    LISTE_ALLE.Add(mepsystem)
    LISTE_LUFT.Add(mepsystem)

class Systemauswahl(forms.WPFWindow):
    def __init__(self, xaml_file_name,liste_Rohr,liste_Luft):
        self.liste_Rohr = liste_Rohr
        self.liste_Luft = liste_Luft
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[MEP_System]()

        try:
            self.backAll()
            self.click(self.luft)
            self.ListView_Sys.ItemsSource = self.liste_Luft
            self.altdatagrid = self.liste_Luft
        except Exception as e:
            logger.error(e)

        self.suche.TextChanged += self.search_txt_changed

    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic
    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal
    def backAll(self):
        self.back(self.luft)
        self.back(self.rohr)
    
    def rohre(self,sender,args):
        self.backAll()
        self.click(self.rohr)
        self.ListView_Sys.ItemsSource = self.liste_Rohr
        self.altdatagrid = self.liste_Rohr
        self.ListView_Sys.Items.Refresh()
        self.suche.Text = ''

    def lueftung(self,sender,args):
        self.backAll()
        self.click(self.luft)
        self.ListView_Sys.ItemsSource = self.liste_Luft
        self.altdatagrid = self.liste_Luft
        self.ListView_Sys.Items.Refresh()
        self.suche.Text = ''
    
    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.ListView_Sys.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.TypName.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.ListView_Sys.ItemsSource = self.tempcoll
        self.ListView_Sys.Items.Refresh()

    def checkall(self,sender,args):
        for item in self.ListView_Sys.Items:
            item.checked = True
        self.ListView_Sys.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.ListView_Sys.Items:
            item.checked = False
        self.ListView_Sys.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.ListView_Sys.Items:
            value = item.checked
            item.checked = not value
        self.ListView_Sys.Items.Refresh()

    def ok(self,sender,args):
        self.Close()
    def abbrechen(self,sender,args):
        self.Close()

systemauswahl = Systemauswahl('window.xaml',LISTE_ROHR,LISTE_LUFT)
try:
    systemauswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    systemauswahl.Close()
    script.exit()

LISTE_NEU = []
for elem in LISTE_ALLE:
    if elem.checked:
        LISTE_NEU.append(elem)

if len(LISTE_NEU) == 0:
    script.exit()

class Element:
    def __init__(self,elemid):
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.raunnr = ''
        try:
            self.raunnr = self.elem.Space[PHASE].Number
        except Exception as e:
            pass
    def wert_schreiben(self):
        try:
            self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raunnr)
        except:
            pass

bearbeitet = []
nichtbearbeitet = LISTE_NEU[:]
t = DB.Transaction(doc,'Einbauort')
t.Start()
if forms.alert('Einbauort schreiben?',yes=True,no=True,ok=False):
    with forms.ProgressBar(title='Systeme --- Einbauort schreiben',cancellable=True, step=10) as pb2:
        for n,mepsystem in enumerate(LISTE_NEU):
            systeme_title = str(mepsystem.TypName) + ' --- ' + str(n+1) + '/' + str(len(LISTE_NEU)) + ' Systeme'
            pb2.title = '{value}/{max_value} Elemente in System ' +  systeme_title
            pb2.step = int((len(mepsystem.liste_bauteile)) /1000) + 10

            for n1,bauteilid in enumerate(mepsystem.liste_bauteile):
                if pb2.cancelled:
                    if forms.alert('bisherige Änderung behalten?',yes=True,no=True,ok=False):
                        t.Commit()
                        logger.info('Folgenede Systeme sind bereits bearbeitet.')
                        for el in bearbeitet:
                            logger.info(el.TypName)
                        logger.info('---------------------------------------')
                        logger.info('Folgenede Systeme sind nicht bearbeitet.')
                        for el in nichtbearbeitet:
                            logger.info(el.TypName)
                    else:
                        t.RollBack()
                    
                    script.exit()
                bauteil = Element(bauteilid)
                try:bauteil.wert_schreiben()
                except:pass

                pb2.update_progress(n1+1, len(mepsystem.liste_bauteile))
            bearbeitet.append(mepsystem)
            nichtbearbeitet.remove(mepsystem)
            
t.Commit()