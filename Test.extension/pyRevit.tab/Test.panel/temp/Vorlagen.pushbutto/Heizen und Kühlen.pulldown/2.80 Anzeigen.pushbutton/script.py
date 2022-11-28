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


__title__ = "2.80 Heiz-/Kühllast"
__doc__ = """Heiz-/Kühllast eines Raumes bzw Bauteile zeigen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
Familien_config = script.get_config()

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

Liste_H_Familie = ObservableCollection[HUK_Bauteile]()
Liste_K_Familie = ObservableCollection[HUK_Bauteile]()
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
        for el in Familien_config.liste_H_Familie:
            if len(Familien_config.liste_H_Familie) == 0:
                break
            Liste_H_Familie.Add(Liste2Class(el))
        for el in Familien_config.liste_K_Familie:
            if len(Familien_config.liste_K_Familie) == 0:
                break
            Liste_K_Familie.Add(Liste2Class(el))
        for el in Familien_config.liste_HUK_Familie:
            if len(Familien_config.liste_HUK_Familie) == 0:
                break
            Liste_HUK_Familie.Add(Liste2Class(el))
        for el in Familien_config.liste_HUKVer_Familie:
            if len(Familien_config.liste_HUKVer_Familie) == 0:
                break
            Liste_HUKVerbinder_Familie.Add(Liste2Class(el))

    except Exception as e:
        pass


def write_config():
    Familien_config.liste_H_Familie = ObservableCollection2Liste(Liste_H_Familie)
    Familien_config.liste_K_Familie = ObservableCollection2Liste(Liste_K_Familie)
    Familien_config.liste_HUK_Familie = ObservableCollection2Liste(Liste_HUK_Familie)
    Familien_config.liste_HUKVer_Familie = ObservableCollection2Liste(Liste_HUKVerbinder_Familie)
    script.save_config()

read_config()

class FamilienBearbeiten(WPFWindow):
    def __init__(self, xaml_file_name,liste_h,liste_k,listeHuK,listeHuKVer):
        self.ListeH = liste_h
        self.ListeK = liste_k
        self.ListeHuK = listeHuK
        self.ListeHuKVer = listeHuKVer
        WPFWindow.__init__(self, xaml_file_name)
        try:
            self.dataGrid.ItemsSource = listeHuK
            self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
            self.dataGrid.Columns[1].ItemsSource = AllHLSParam
            self.dataGrid.Columns[2].ItemsSource = AllHLSParam
        except:
            pass

    def cancel(self,sender,args):
        self.Close()
        self.ListeH.Clear()
        self.ListeK.Clear()
        self.ListeHuK.Clear()
        self.ListeHuKVer.Clear()
        read_config()


    def OK(self,sender,args):
        self.Close()

    def Heating(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeH
        self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
        self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        self.dataGrid.Columns[2].ItemsSource = AllHLSParam

        self.dataGrid.Columns[1].Width = DataGridLength(300.0)
        self.dataGrid.Columns[2].Width = DataGridLength(0.0)



    def Cooling(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeK
        self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
        self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        self.dataGrid.Columns[2].ItemsSource = AllHLSParam
        self.dataGrid.Columns[1].Width = DataGridLength(0.0)
        self.dataGrid.Columns[2].Width = DataGridLength(300.0)


    def HundK(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeHuK
        self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
        self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        self.dataGrid.Columns[2].ItemsSource = AllHLSParam

        self.dataGrid.Columns[1].Width = DataGridLength(150.0)
        self.dataGrid.Columns[2].Width = DataGridLength(150.0)


    def HundKVer(self,sender,args):
        self.dataGrid.ItemsSource = self.ListeHuKVer
        self.dataGrid.Columns[0].ItemsSource = AllHLSNamen
        self.dataGrid.Columns[1].ItemsSource = AllHLSParam
        self.dataGrid.Columns[2].ItemsSource = AllHLSParam

        self.dataGrid.Columns[1].Width = DataGridLength(150.0)
        self.dataGrid.Columns[2].Width = DataGridLength(150.0)

    def Add(self,sender,args):
        temp_class = HUK_Bauteile()
        self.dataGrid.ItemsSource.Add(temp_class)
        self.dataGrid.Items.Refresh()


FamilienDialog = FamilienBearbeiten('Test.xaml', Liste_H_Familie, Liste_K_Familie, Liste_HUK_Familie, Liste_HUKVerbinder_Familie)
FamilienDialog.ShowDialog()

write_config()

HZK_Familien = {}
DeS_Familien = {}
DeS_Verbinder_Familien = {}
ULK_Familien = {}
for el in Liste_H_Familie:
    HZK_Familien[el.Element] = el.HLeistungSchreiben
for el in Liste_K_Familie:
    ULK_Familien[el.Element] = el.KLeistungSchreiben
for el in Liste_HUK_Familie:
    DeS_Familien[el.Element] = [el.HLeistungSchreiben,el.KLeistungSchreiben]
for el in Liste_HUKVerbinder_Familie:
    DeS_Verbinder_Familien[el.Element] = [el.HLeistungSchreiben,el.KLeistungSchreiben]

Space_H_Dict = {}
Space_K_Dict = {}

for el in HLS_collector:
    try:
        FamilieName = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
        space = el.Space[phase]
        if not space:
            continue
        Nr = space.Number
        if FamilieName in HZK_Familien.keys():
            if not Nr in Space_H_Dict.keys():
                Space_H_Dict[Nr] = [el]
            else:
                Space_H_Dict[Nr].append(el)
        elif FamilieName in DeS_Familien.keys():
            if not Nr in Space_K_Dict.keys():
                Space_K_Dict[Nr] = [el]
            else:
                Space_K_Dict[Nr].append(el)

            if not Nr in Space_H_Dict.keys():
                Space_H_Dict[Nr] = [el]
            else:
                Space_H_Dict[Nr].append(el)

        elif FamilieName in DeS_Verbinder_Familien.keys():
            if not Nr in Space_K_Dict.keys():
                Space_K_Dict[Nr] = [el]
            else:
                Space_K_Dict[Nr].append(el)

            if not Nr in Space_H_Dict.keys():
                Space_H_Dict[Nr] = [el]
            else:
                Space_H_Dict[Nr].append(el)

        elif FamilieName in ULK_Familien.keys():
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
        trans1 = TransactionManager.Instance
        trans1.ForceCloseTransaction()
        if self.Text == 'Heizung Gleichmäßig':
            t = Transaction(doc,self.Text)
            t.Start()
            hl = self.RaumHL[:-1]
            leistung = round(int(hl)/len(self.EleListe))
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                self.ListeH[n].Leistung = str(leistung) + ' W'
                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in HZK_Familien.keys():
                    try:
                        el.LookupParameter(HZK_Familien[Familie]).SetValueString(str(leistung))
                    except Exception as e:
                        pass
                elif Familie in DeS_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Familien[Familie][0]).SetValueString(str(leistung))
                    except Exception as e:
                        pass
                elif Familie in DeS_Verbinder_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Verbinder_Familien[Familie][0]).SetValueString(str(leistung))
                    except Exception as e:
                        pass
            doc.Regenerate()
            t.Commit()
        elif self.Text == 'Kühlung Gleichmäßig':
            t = Transaction(doc,self.Text)
            t.Start()
            kl = self.RaumKL[:-1]
            leistung = round(int(kl)/len(self.EleListe))
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                self.ListeK[n].Leistung = str(int(leistung)) + ' W'
                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in ULK_Familien.keys():
                    try:
                        el.LookupParameter(ULK_Familien[Familie]).SetValueString(str(leistung))
                    except Exception as e:
                        logger.error(e)

                elif Familie in DeS_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Familien[Familie][1]).SetValueString(str(leistung))
                    except Exception as e:
                        logger.error(e)
                elif Familie in DeS_Verbinder_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Verbinder_Familien[Familie][1]).SetValueString(str(leistung))
                    except Exception as e:
                        logger.error(e)
            doc.Regenerate()
            t.Commit()

        self.dataGrid.Items.Refresh()

    def Heating(self,sender,args):
        self.glei.Visibility = System.Windows.Visibility.Visible
        self.over.Visibility = System.Windows.Visibility.Visible
        self.dataGrid.ItemsSource = self.ListeH
        self.dataGrid.Columns[0].Header = 'Heizen'
        self.dataGrid.Columns[1].Header = 'Heizleistung'
        self.Text = 'Heizung Gleichmäßig'
        self.EleListe = self.HEleListe
        self.Liste = self.ListeH

    def Cooling(self,sender,args):
        self.glei.Visibility = System.Windows.Visibility.Visible
        self.over.Visibility = System.Windows.Visibility.Visible
        self.dataGrid.ItemsSource = self.ListeK
        self.dataGrid.Columns[0].Header = 'Kühlen'
        self.dataGrid.Columns[1].Header = 'Kühlleistung'
        self.Text = 'Kühlung Gleichmäßig'
        self.EleListe = self.KEleListe
        self.Liste = self.ListeK

    def overwrite(self,sender,args):
        trans1 = TransactionManager.Instance
        trans1.ForceCloseTransaction()

        if self.Text == 'Kühlung Gleichmäßig':
            t1 = Transaction(doc,'Überschreiben')
            t1.Start()
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                leistung = self.Liste[n].Leistung
                leis = None
                try:
                    leis = int(leistung)
                except:
                    leis = int(leistung[:-1])

                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in ULK_Familien.keys():
                    try:
                        el.LookupParameter(ULK_Familien[Familie]).SetValueString(str(leis))
                    except Exception as e:
                        pass
                elif Familie in DeS_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Familien[Familie][1]).SetValueString(str(leis))
                    except Exception as e:
                        pass
                elif Familie in DeS_Verbinder_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Verbinder_Familien[Familie][1]).SetValueString(str(leis))
                    except Exception as e:
                        pass
                try:
                    leis = int(self.Liste[n].Leistung)
                    self.Liste[n].Leistung = self.Liste[n].Leistung+' W'
                except:
                    pass
            doc.Regenerate()
            t1.Commit()


        elif self.Text == 'Heizung Gleichmäßig':
            t1 = Transaction(doc,'Überschreiben')
            t1.Start()
            for n in range(len(self.EleListe)):
                el = self.EleListe[n]
                leistung = self.Liste[n].Leistung
                leis = None
                try:
                    leis = int(leistung)
                except:
                    leis = int(leistung[:-1])

                Familie = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if Familie in HZK_Familien.keys():
                    try:
                        el.LookupParameter(HZK_Familien[Familie]).SetValueString(str(leis))
                    except Exception as e:
                        pass
                elif Familie in DeS_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Familien[Familie][0]).SetValueString(str(leis))
                    except Exception as e:
                        pass
                elif Familie in DeS_Verbinder_Familien.keys():
                    try:
                        el.LookupParameter(DeS_Verbinder_Familien[Familie][0]).SetValueString(str(leis))
                    except Exception as e:
                        pass
                try:
                    leis = int(self.Liste[n].Leistung)
                    self.Liste[n].Leistung = self.Liste[n].Leistung+' W'
                except:
                    pass
            doc.Regenerate()
            t1.Commit()
        self.dataGrid.Items.Refresh()

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
    H_elemliste_new = []
    K_elemliste_new = []
    nr = None
    if selec.GetType().ToString() == 'Autodesk.Revit.DB.Mechanical.Space':
        nr = selec.Number
        if nr in Space_H_Dict.keys():
            HL = selec.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
            KL = selec.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').AsValueString()
            H_elemliste = Space_H_Dict[nr]
            hzk = 0
            des = 0
            Verbinder = False
            for el in H_elemliste:
                el_Name = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if el_Name in DeS_Verbinder_Familien.keys():
                    Verbinder = True

            for el in H_elemliste:
                HeizK = Heiz()
                Hl = None
                el_Name = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()

                if el_Name in HZK_Familien.keys():
                    H_elemliste_new.append(el)
                    hzk += 1
                    Name = 'HZK-' + str(hzk)
                    HeizK.Element = Name
                    HeizK.Leistung = 0
                    try:
                        HLeistung = el.LookupParameter(HZK_Familien[el_Name]).AsValueString()
                        HeizK.Leistung = HLeistung
                    except:
                        pass
                    Liste_H.Add(HeizK)

                elif el_Name in DeS_Familien.keys():
                    if not Verbinder:
                        H_elemliste_new.append(el)
                        des += 1
                        Name = 'DeS-' + str(des)
                        HeizK.Element = Name
                        HeizK.Leistung = 0
                        try:
                            HLeistung = el.LookupParameter(DeS_Familien[el_Name][0]).AsValueString()
                            HeizK.Leistung = HLeistung
                        except:
                            pass

                        Liste_H.Add(HeizK)

                elif el_Name in DeS_Verbinder_Familien.keys():
                    H_elemliste_new.append(el)
                    des += 1
                    Name = 'DeS_Verbinder-' + str(des)
                    HeizK.Element = Name
                    HeizK.Leistung = 0
                    try:
                        HLeistung = el.LookupParameter(DeS_Verbinder_Familien[el_Name][0]).AsValueString()
                        HeizK.Leistung = HLeistung
                    except:
                        pass
                    Liste_H.Add(HeizK)

        if nr in Space_K_Dict.keys():
            HL = selec.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
            KL = selec.LookupParameter('LIN_BA_CALCULATED_COOLING_LOAD').AsValueString()
            K_elemliste = Space_K_Dict[nr]
            ulk = 0
            des = 0
            Verbinder = False
            for el in K_elemliste:
                el_Name = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                if el_Name in DeS_Verbinder_Familien.keys():
                    Verbinder = True

            for el in K_elemliste:
                el_Name = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                Kl = None
                Cooling = Kuehl()
                if el_Name in ULK_Familien.keys():
                    K_elemliste_new.append(el)
                    ulk += 1
                    Name = 'ULK-' + str(ulk)
                    Cooling.Element = Name
                    Cooling.Leistung = 0
                    try:
                        KLeistung = el.LookupParameter(ULK_Familien[el_Name]).AsValueString()
                        Cooling.Leistung = KLeistung
                    except:
                        pass

                    Liste_K.Add(Cooling)

                elif el_Name in DeS_Familien.keys():
                    if not Verbinder:
                        K_elemliste_new.append(el)
                        des += 1
                        Name = 'DeS-' + str(des)
                        Cooling.Element = Name
                        Cooling.Leistung = 0
                        try:
                            KLeistung = el.LookupParameter(DeS_Familien[el_Name][1]).AsValueString()
                            Cooling.Leistung = KLeistung
                        except:
                            pass
                        Liste_K.Add(Cooling)

                elif el_Name in DeS_Verbinder_Familien.keys():
                    K_elemliste_new.append(el)
                    des += 1
                    Name = 'DeS_Verbinder-' + str(des)
                    Cooling.Element = Name
                    Cooling.Leistung = 0
                    try:
                        KLeistung = el.LookupParameter(DeS_Verbinder_Familien[el_Name][1]).AsValueString()
                        Cooling.Leistung = KLeistung
                    except:
                        pass
                    Liste_K.Add(Cooling)

    else:
        TaskDialog.Show("fehler","Keine MEP Raum ausgewählt!")

    if any(H_elemliste) or any(K_elemliste):
        HeizWPF= HeizUKuehl('window.xaml',Liste_H,Liste_K,nr,HL,KL,H_elemliste_new,K_elemliste_new)
        HeizWPF.ShowDialog()
    else:
        TaskDialog.Show("fehler","Keine HZK/ULK/DeS im ausgewählten Raum gefunden! Bitte wählen Sie ein neu Raum.")
        PickMEP()

PickMEP()

# total = time.time() - start
# logger.info("total time: {} {}".format(total, 100 * "_"))
