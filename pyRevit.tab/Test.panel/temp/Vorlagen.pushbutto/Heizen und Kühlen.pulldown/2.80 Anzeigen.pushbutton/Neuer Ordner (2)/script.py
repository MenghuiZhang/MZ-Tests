# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw,clr
clr.AddReference("PresentationFramework")
import time
from System.Windows.Controls import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit.forms import WPFWindow, SelectFromList
import System
from System.Windows import Application, Window
from System.Collections.ObjectModel import *
import os.path as op



start = time.time()


__title__ = "2.80 Heiz-/Kühllast"
__doc__ = """Heiz-/Kühllast eines Raumes bzw Bauteile zeigen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc
phase = list(doc.Phases)[-1]

# HLS aus aktueller Projekt
HLS_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
    .WhereElementIsNotElementType()
HLS = HLS_collector.ToElementIds()

Familien = [
'_H_CAx FL Fußboden heizmatte',
'_H_CAx HZK Glieder RLVL unten',
'_H_CAx HZK Plane Platte RLVL unten',
'_H_CAx HZK Plane Platte VLRL Seite',
'_K_IGF_435_Deckensegel',
'_K_IGF_435_Deckensegelverbinder',
'_K_IGF_FG_Flex Geko LE',
'_K_IGF_FG_Flex Geko RE'
]
HZK_Familien=[
'_H_CAx FL Fußboden heizmatte',
'_H_CAx HZK Glieder RLVL unten',
'_H_CAx HZK Plane Platte RLVL unten',
'_H_CAx HZK Plane Platte VLRL Seite',
]
DeS_Familien=[
'_K_IGF_435_Deckensegel',
'_K_IGF_435_Deckensegelverbinder',
]
ULK_Familien=[
'_K_IGF_FG_Flex Geko LE',
'_K_IGF_FG_Flex Geko RE'
]

Space_H_Dict = {}
Space_K_Dict = {}
spcae_dict = []
for el in HLS_collector:
    try:
        FamilieName = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
        space = el.Space[phase]
        if not space:
            continue
        Nr = space.Number
        if FamilieName in HZK_Familien:
            if not Nr in Space_H_Dict.keys():
                Space_H_Dict[Nr] = [el]
            else:
                Space_H_Dict[Nr].append(el)
        elif FamilieName in DeS_Familien:
            if not Nr in Space_K_Dict.keys():
                Space_K_Dict[Nr] = [el]
                Space_H_Dict[Nr] = [el]
            else:
                Space_K_Dict[Nr].append(el)
                Space_H_Dict[Nr].append(el)
        elif FamilieName in ULK_Familien:
            if not Nr in Space_K_Dict.keys():
                Space_K_Dict[Nr] = [el]
            else:
                Space_K_Dict[Nr].append(el)

    except Exception as e:
        logger.error(e)


class Heiz(object):
    @property
    def Element(self):
        return self._Element
    @Element.setter
    def Element(self, value):
        self._Element = value
    @property
    def Leistung(self):
        return self._Leistung
    @Leistung.setter
    def Leistung(self, value):
        self._Leistung = value
class Kuehl(object):
    @property
    def Element(self):
        return self._Element
    @Element.setter
    def Element(self, value):
        self._Element = value
    @property
    def Leistung(self):
        return self._Leistung
    @Leistung.setter
    def Leistung(self, value):
        self._Leistung = value

class Heizwpf(WPFWindow):
    def __init__(self, xaml_file_name,liste_h,liste_k,raumnr,raumhl,raumkl):#eleliste
        self.ListeH = liste_h
        self.ListeK = liste_k
#        self.EleListe = eleliste
        self.RaumNr = raumnr
        self.RaumHL = raumhl
        self.RaumKL = raumkl
        WPFWindow.__init__(self, xaml_file_name)
#        self.dataGrid.ItemsSource = self.Liste
        self.Nummer.Text = self.RaumNr
        self.Heiz.Text = self.RaumHL
        self.Kuehl.Text = self.RaumKL

    def close(self,sender,args):
        self.Close()
    def gleich(self,sender,args):
        #selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()][0]
        # Daten(selection)
        self.Close()
        # t = Transaction(doc,'Heiz Gleichmäßig')
        # t.Start()
        # for el in self.EleListe:
        #     try:
        #         pass
        #     except Exception as e:
        #         logger.error(e)
        # t.Commit()

    def Heating(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeH
        self.dataGrid.Columns[0].Header = 'Heizen'
        self.dataGrid.Columns[1].Header = 'Heizleistung'

    def Cooling(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeK
        self.dataGrid.Columns[0].Header = 'Kühlen'
        self.dataGrid.Columns[1].Header = 'Kühlleistung'

    def cancel(self,sender,args):
        self.Close()
        total = time.time() - start
        logger.info("total time: {} {}".format(total, 100 * "_"))
        script.exit()
    # def reload(self,sender,args):
    #     self.Close()
    #     total = time.time() - start
    #     logger.info("total time: {} {}".format(total, 100 * "_"))
    #     script.exit()

def Daten(selec):
    Liste_H = ObservableCollection[Heiz]()
    Liste_K = ObservableCollection[Kuehl]()
    HL = None
    KL = None
    H_elemliste = []
    K_elemliste = []
    nr = None
    if selec.GetType().ToString() == 'Autodesk.Revit.DB.Mechanical.Space':
        nr = selec.Number
        if nr in Space_H_Dict.keys():
            HL = selec.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
            KL = selec.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').AsValueString()
            H_elemliste = Space_H_Dict[nr]
            hzk = 0
            des = 0
            for el in H_elemliste:
                HeizK = Heiz()
                if el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in HZK_Familien:
                    hzk += 1
                    Name = 'HZK-' + str(hzk)
                    HeizK.Element = Name
                    HeizK.Leistung = el.LookupParameter('CAx_Heizleistung').AsValueString()
                    Liste_H.Add(HeizK)
                elif el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in DeS_Familien:
                    des += 1
                    Name = 'DeS-' + str(des)
                    HeizK.Element = Name
                    HeizK.Leistung = el.LookupParameter('Leistung_Heizen').AsValueString()
                    Liste_H.Add(HeizK)
        elif nr in Space_K_Dict.keys():
            HL = selec.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
            KL = selec.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').AsValueString()
            K_elemliste = Space_K_Dict[nr]
            ulk = 0
            des = 0
            for el in K_elemliste:
                Cooling = Kuehl()
                if el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in ULK_Familien:
                    ulk += 1
                    Name = 'ULK-' + str(ulk)
                    Cooling.Element = Name
                    Cooling.Leistung = el.LookupParameter('Leistung_Kühlen').AsValueString()
                    Liste_K.Add(Cooling)
                elif el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in DeS_Familien:
                    des += 1
                    Name = 'DeS-' + str(des)
                    Cooling.Element = Name
                    Cooling.Leistung = el.LookupParameter('Leistung_Kühlen').AsValueString()
                    Liste_K.Add(Cooling)
    if any(H_elemliste) or any(K_elemliste):
        HeizWPF= Heizwpf('window_H.xaml',Liste_H,Liste_K,nr,HL,KL)
        HeizWPF.ShowDialog()

selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()][0]
selection_temp = selection
Daten(selection)
# while(selection == selection_temp):
#     selection_temp = selection
#     Daten(selection)
#     selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()][0]
    # if forms.alert("weiter?", ok=False, yes=True, no=True):
    #     selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()][0]
    # else:
    #     script.exit()

    # else:
    #     TaskDialog.Show(nr,'Kein HZK im gewählten Raum gefunden')




total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
