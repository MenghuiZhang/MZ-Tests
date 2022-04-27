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
from System.Windows import Application, Window
from System.Collections.ObjectModel import *
import os.path as op
#
# start = time.time()


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

HZK_Familien=[
'_H_CAx FL Fußboden heizmatte',
'Charleston-3-1200-3000',
'_H_IGF_MC_HZK Charleston RLVL unten',
'_H_IGF_MC_HZK Plane Platte RLVL unten',
'_H_CAx HZK Plane Platte VLRL Seite',
]
DeS_Familien=[
'_K_IGF_435_Deckensegel',
'_K_IGF_435_Deckensegelverbinder_Berechnung',
]
ULK_Familien=[
'_K_IGF_FG_Flex Geko LE',
'_K_IGF_FG_Flex Geko RE'
]

Space_H_Dict = {}
Space_K_Dict = {}

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
            else:
                Space_K_Dict[Nr].append(el)

            if not Nr in Space_H_Dict.keys():
                Space_H_Dict[Nr] = [el]
            else:
                Space_H_Dict[Nr].append(el)

        elif FamilieName in ULK_Familien:
            if not Nr in Space_K_Dict.keys():
                Space_K_Dict[Nr] = [el]
            else:
                Space_K_Dict[Nr].append(el)

    except Exception as e:
        logger.error(e)

class MEPRaumFilter(ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Name == 'MEP-Räume':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False


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

def PickMEP():
    try:
        re = uidoc.Selection.PickObject(Selection.ObjectType.Element,MEPRaumFilter())
        elem = doc.GetElement(re)
        Daten(elem)
    except:
        pass

class HeizUKuehl(WPFWindow):
    def __init__(self, xaml_file_name,liste_h,liste_k,raumnr,raumhl,raumkl,h_ele,K_ele):
        self.ListeH = liste_h
        self.ListeK = liste_k
        self.RaumNr = raumnr
        self.RaumHL = raumhl
        self.RaumKL = raumkl
        WPFWindow.__init__(self, xaml_file_name)
        self.Nummer.Text = self.RaumNr
        self.Heiz.Text = self.RaumHL
        self.Kuehl.Text = self.RaumKL
        self.Text = None
        self.EleListe = None
        self.HEleListe = h_ele
        self.KEleListe = K_ele
        self.Liste = None

    def close(self,sender,args):
        self.Close()

    def gleich(self,sender,args):
        t = Transaction(doc,self.Text)
        t.Start()
        if self.EleListe == self.HEleListe:
            hl = self.RaumHL[:-1]
            leistung = round(int(hl)/len(self.EleListe))
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                self.ListeH[n].Leistung = str(leistung) + ' W'
                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in HZK_Familien:
                    try:
                        #el.LookupParameter('IGF_H_HK_Leistung').Set(leistung)
                        el.LookupParameter('MC Piping Power').SetValueString(str(leistung))
                    except Exception as e:
                        pass
                elif Familie in DeS_Familien:
                    try:
                        #el.LookupParameter('IGF_DeS_Heizung').Set(leistung)
                        el.LookupParameter('MC Piping Power').SetValueString(str(leistung))
                    except Exception as e:
                        pass
        elif self.EleListe == self.KEleListe:
            kl = self.RaumKL[:-1]
            leistung = round(int(kl)/len(self.EleListe))
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                self.ListeK[n].Leistung = str(int(leistung)) + ' W'
                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in ULK_Familien:
                    try:
                        el.LookupParameter('MC Cooling Power').SetValueString(str(leistung))
                    except Exception as e:
                        pass

                elif Familie in DeS_Familien:
                    try:
                        el.LookupParameter('MC Cooling Power').SetValueString(str(leistung))
                    except Exception as e:
                        pass

        t.Commit()
        self.dataGrid.Items.Refresh()

    def Heating(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeH
        self.dataGrid.Columns[0].Header = 'Heizen'
        self.dataGrid.Columns[1].Header = 'Heizleistung'
        self.Text = 'Heizung Gleichmäßig'
        self.EleListe = self.HEleListe
        self.Liste = self.ListeH

    def Cooling(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeK
        self.dataGrid.Columns[0].Header = 'Kühlen'
        self.dataGrid.Columns[1].Header = 'Kühlleistung'
        self.Text = 'Kühlung Gleichmäßig'
        self.EleListe = self.KEleListe
        self.Liste = self.ListeK

    def overwrite(self,sender,args):
        t = Transaction(doc,'Überschreiben')
        t.Start()
        if self.EleListe == self.KEleListe:
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                leistung = self.Liste[n].Leistung
                leis = None
                try:
                    leis = int(leistung)
                except:
                    leis = int(leistung[:-1])

                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in ULK_Familien:
                    try:
                        el.LookupParameter('MC Cooling Power').SetValueString(str(leis))
                    except Exception as e:
                        logger.error(e)
                elif Familie in DeS_Familien:
                    try:
                        el.LookupParameter('MC Cooling Power').SetValueString(str(leis))
                    except Exception as e:
                        logger.error(e)

        elif self.EleListe == self.HEleListe:
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                leistung = self.Liste[n].Leistung
                leis = None
                try:
                    leis = int(leistung)
                except:
                    leis = int(leistung[:-1])

                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in HZK_Familien:
                    try:
                        el.LookupParameter('MC Piping Power').SetValueString(str(leis))
                    except Exception as e:
                        logger.error(e)
                elif Familie in DeS_Familien:
                    try:
                        el.LookupParameter('MC Piping Power').SetValueString(str(leis))
                    except Exception as e:
                        logger.error(e)
        t.Commit()

    def weiter(self,sender,args):
        self.Close()
        PickMEP()

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
                Hl = None

                if el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in HZK_Familien:
                    hzk += 1
                    Name = 'HZK-' + str(hzk)
                    HeizK.Element = Name
                    if el.LookupParameter('MC Piping Power'):
                        HLeistung = el.LookupParameter('MC Piping Power').AsValueString()
                        if HLeistung != '0 W':
                            HeizK.Leistung = HLeistung
                        else:
                            if el.LookupParameter('CAx_Heizleistung'):
                                HLeistung = el.LookupParameter('CAx_Heizleistung').AsValueString()
                                HeizK.Leistung = HLeistung
                    else:
                        if el.LookupParameter('CAx_Heizleistung'):
                            HLeistung = el.LookupParameter('CAx_Heizleistung').AsValueString()
                            HeizK.Leistung = HLeistung

                    Liste_H.Add(HeizK)
                elif el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in DeS_Familien:
                    des += 1
                    Name = 'DeS-' + str(des)
                    HeizK.Element = Name
                    if el.LookupParameter('MC Piping Power'):
                        HLeistung = el.LookupParameter('MC Piping Power').AsValueString()
                        if HLeistung != '0 W':
                            HeizK.Leistung = HLeistung
                        else:
                            if el.LookupParameter('Leistung_Heizen'):
                                HLeistung = el.LookupParameter('Leistung_Heizen').AsValueString()
                                HeizK.Leistung = HLeistung
                    else:
                        if el.LookupParameter('Leistung_Heizen'):
                            HLeistung = el.LookupParameter('Leistung_Heizen').AsValueString()
                            HeizK.Leistung = HLeistung

                    Liste_H.Add(HeizK)
        if nr in Space_K_Dict.keys():
            HL = selec.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
            KL = selec.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').AsValueString()
            K_elemliste = Space_K_Dict[nr]
            ulk = 0
            des = 0

            for el in K_elemliste:
                Kl = None
                Cooling = Kuehl()
                if el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in ULK_Familien:
                    ulk += 1
                    Name = 'ULK-' + str(ulk)
                    Cooling.Element = Name
                    if el.LookupParameter('MC Cooling Power'):
                        KLeistung = el.LookupParameter('MC Cooling Power').AsValueString()
                        if KLeistung != '0 W':
                            Cooling.Leistung = KLeistung
                        else:
                            if el.LookupParameter('Leistung_Kühlen'):
                                KLeistung = el.LookupParameter('Leistung_Kühlen').AsValueString()
                                Cooling.Leistung = KLeistung
                    else:
                        if el.LookupParameter('Leistung_Kühlen'):
                            KLeistung = el.LookupParameter('Leistung_Kühlen').AsValueString()
                            Cooling.Leistung = KLeistung

                    Liste_K.Add(Cooling)
                elif el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() in DeS_Familien:
                    des += 1
                    Name = 'DeS-' + str(des)
                    Cooling.Element = Name
                    if el.LookupParameter('MC Cooling Power'):
                        KLeistung = el.LookupParameter('MC Cooling Power').AsValueString()
                        if KLeistung != '0 W':
                            Cooling.Leistung = KLeistung
                        else:
                            if el.LookupParameter('Leistung_Kühlen'):
                                KLeistung = el.LookupParameter('Leistung_Kühlen').AsValueString()
                                Cooling.Leistung = KLeistung
                    else:
                        if el.LookupParameter('Leistung_Kühlen'):
                            KLeistung = el.LookupParameter('Leistung_Kühlen').AsValueString()
                            Cooling.Leistung = KLeistung

                    Liste_K.Add(Cooling)
    else:
        TaskDialog.Show("fehler","Keine MEP Raum ausgewählt!")

    if any(H_elemliste) or any(K_elemliste):
        HeizWPF= HeizUKuehl('window.xaml',Liste_H,Liste_K,nr,HL,KL,H_elemliste,K_elemliste)
        HeizWPF.ShowDialog()
    else:
        TaskDialog.Show("fehler","Keine HZK/ULK/DeS im ausgewählten Raum gefunden! Bitte wählen Sie ein neu Raum.")
        PickMEP()

PickMEP()

# total = time.time() - start
# logger.info("total time: {} {}".format(total, 100 * "_"))
