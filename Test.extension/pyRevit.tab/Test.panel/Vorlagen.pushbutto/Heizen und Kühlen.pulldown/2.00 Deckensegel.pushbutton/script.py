# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from IGF_libKopie import get_value
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter


__title__ = "2.00  Daten von DeS in MEP schreiben"
__doc__ = """Temperatur und Volumenstrom bzw. Heiz- und Kühlleistung der Deckensegel in MEP Raum schreiben

DeS Category: HLS-Bauteile
Parameter in DeS: manuell definiert
Parameter in MEP-Räume:
IGF_H_DeS-VL_Win_Temp: VL_Temp Heizung
IGF_H_DeS-RL_Win_Temp: RL_Temp Heizung
IGF_K_DeS-VL_Som_Temp: VL_Temp Kühlung
IGF_K_DeS-RL_Som_Temp: RL_Temp Kühlung
IGF_H_DeS_Winter: Volumenstrom Heizung
IGF_K_DeS_Sommer: Volumenstrom Kühlung
IGF_H_DeS_Leistung: Leistung Heizung
IGF_K_DeS_Leistung: Leistung Kühlung

[Version: 1.1]
[2021.11.16]


"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'DeS-Familie')

try:
    getlog(__title__)
except:
    pass


bauteile_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
bauteile = bauteile_collector.ToElementIds()
bauteile_collector.Dispose()

phase = doc.Phases[0]

HLS_Dict = {}
HLS_Para_Dict = {}
for elid in bauteile:
    elem = doc.GetElement(elid)
    Family = elem.Symbol.FamilyName
    if not Family in HLS_Dict.keys():
        HLS_Dict[Family] = [elem]
        Paraliste = []
        for para in elem.Parameters:
            if not para.Definition.Name in Paraliste:
                try:
                    Paraliste.append(para.Definition.Name)
                except:
                    pass
        HLS_Para_Dict[Family]=Paraliste
    else:
        HLS_Dict[Family].append(elem)


class MechanicalEquipment(object):
    def __init__(self,name):
        self.Name = name
        self.Vol_Id = -1
        self.V_H_Id = -1
        self.R_H_Id = -1
        self.V_K_Id = -1
        self.R_K_Id = -1
        self.KL_Id = -1
        self.HL_Id = -1
        self.Paras = []
        self.checked = False

    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self, value):
        self._Name = value
    
    @property
    def Paras(self):
        return self._Paras
    @Paras.setter
    def Paras(self, value):
        self._Paras = value
    @property
    def Para_dict(self):
        return self._Para_dict
    @Para_dict.setter
    def Para_dict(self, value):
        self._Para_dict = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def Para_id_dict(self):
        return self._Para_id_dict
    @Para_id_dict.setter
    def Para_id_dict(self, value):
        self._Para_id_dict = value
    
    @property
    def Vol_Id(self):
        return self._Vol_Id
    @Vol_Id.setter
    def Vol_Id(self, value):
        self._Vol_Id = value
    @property
    def V_H_Id(self):
        return self._V_H_Id
    @V_H_Id.setter
    def V_H_Id(self, value):
        self._V_H_Id = value
    @property
    def R_H_Id(self):
        return self._R_H_Id
    @R_H_Id.setter
    def R_H_Id(self, value):
        self._R_H_Id = value
    @property
    def R_K_Id(self):
        return self._R_K_Id
    @R_K_Id.setter
    def R_K_Id(self, value):
        self._R_K_Id = value
    @property
    def V_K_Id(self):
        return self._V_K_Id
    @V_K_Id.setter
    def V_K_Id(self, value):
        self._V_K_Id = value
    @property
    def HL_Id(self):
        return self._HL_Id
    @HL_Id.setter
    def HL_Id(self, value):
        self._HL_Id = value
    @property
    def KL_Id(self):
        return self._KL_Id
    @KL_Id.setter
    def KL_Id(self, value):
        self._KL_Id = value

class select(object):
    def __init__(self,n,param):
        self.selectId = n
        self.ParaName = param

    @property
    def ParaName(self):
        return self._ParaName
    @ParaName.setter
    def ParaName(self, value):
        self._ParaName = value
    @property
    def selectId(self):
        return self._selectId
    @selectId.setter
    def selectId(self, value):
        self._selectId = value

Liste_DeS = ObservableCollection[MechanicalEquipment]()

DeSName = HLS_Para_Dict.keys()[:]
DeSName.sort()

for elem in DeSName:
    temp = MechanicalEquipment(elem)
    params = HLS_Para_Dict[elem][:]
    params.sort()
    para_Liste = []
    para_dict = {}
    para_id_dict = {}
    for n,para in enumerate(params):
        temp_para = select(n,para)
        para_dict[n] = para
        para_id_dict[para] = n
        para_Liste.append(temp_para)
    temp.Paras = para_Liste
    temp.Para_dict = para_dict
    temp.Para_id_dict = para_id_dict
    Liste_DeS.Add(temp)

class DeS_Familie(forms.WPFWindow):
    def __init__(self, xaml_file_name,Liste_DeS):
        self.Liste_DeS = Liste_DeS
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.altdatagrid = Liste_DeS
        self.dataGrid.ItemsSource = Liste_DeS
        self.read_config()
    
    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic
    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal
        
    
    def read_config(self):
        try:
            DeS_dict = config.DeS_dict
            for item in self.Liste_DeS:
                if item.Name in DeS_dict.keys():
                    try:
                        item.checked = DeS_dict[item.Name][0]
                    except:
                        pass
                    try:
                        item.Vol_Id = item.Para_id_dict[DeS_dict[item.Name][1]]
                    except:
                        pass
                    try:
                        item.V_H_Id = item.Para_id_dict[DeS_dict[item.Name][2]]
                    except:
                        pass
                    try:
                        item.R_H_Id = item.Para_id_dict[DeS_dict[item.Name][3]]
                    except:
                        pass
                    try:
                        item.V_K_Id = item.Para_id_dict[DeS_dict[item.Name][4]]
                    except:
                        pass
                    try:
                        item.R_K_Id = item.Para_id_dict[DeS_dict[item.Name][5]]
                    except:
                        pass
                    try:
                        item.KL_Id = item.Para_id_dict[DeS_dict[item.Name][6]]
                    except:
                        pass
                    try:
                        item.HL_Id = item.Para_id_dict[DeS_dict[item.Name][7]]
                    except:
                        pass
            self.dataGrid.Items.Refresh()
         
        except:
            pass

        try:
            if config.Hide:
                self.click(self.aus)
                for item in self.dataGrid.Items:
                    if item.checked:
                        self.tempcoll.Add(item)
                self.dataGrid.ItemsSource = self.tempcoll
                self.dataGrid.Items.Refresh()
            else:
                self.click(self.ein)
                   
        except:
            pass
        
        
    def write_config(self):
        dict_ = {}
        for item in self.Liste_DeS:
            if item.checked:
                try:
                    dict_[item.Name] = [item.checked,
                    item.Para_dict[item.Vol_Id],
                    item.Para_dict[item.V_H_Id],
                    item.Para_dict[item.R_H_Id],
                    item.Para_dict[item.V_K_Id],
                    item.Para_dict[item.R_K_Id],
                    item.Para_dict[item.KL_Id],
                    item.Para_dict[item.HL_Id]
                    ]
                except:
                    pass

        config.DeS_dict = dict_
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        self.Close()

    def hide(self,sender,args):
        self.tempcoll.Clear()
        for item in self.dataGrid.Items:
            if item.checked:
                self.tempcoll.Add(item)
        self.dataGrid.ItemsSource = self.tempcoll
        config.Hide = True
        self.click(self.aus)
        self.back(self.ein)

    def show(self,sender,args):
        self.dataGrid.ItemsSource = self.altdatagrid
        config.Hide = False
        self.click(self.ein)
        self.back(self.aus)

    def close(self,sender,args):
        self.Close()
        script.exit()
    
DeS_Familie_Auswahl = DeS_Familie("window.xaml",Liste_DeS)
try:
    DeS_Familie_Auswahl.ShowDialog()
except Exception as e:
    DeS_Familie_Auswahl.Close()
    script.exit()

Familie_Bearbeiten = {}
for el in Liste_DeS:
    if el.checked:
        for elem in HLS_Dict[el.Name]:
            Familie_Bearbeiten[elem] = [
                el.Para_dict[el.Vol_Id],
                el.Para_dict[el.V_H_Id],
                el.Para_dict[el.R_H_Id],
                el.Para_dict[el.V_K_Id],
                el.Para_dict[el.R_K_Id],
                el.Para_dict[el.KL_Id],
                el.Para_dict[el.HL_Id]]

if len(Familie_Bearbeiten.keys()) == 0:
    script.exit()


class Deckensegel:
    def __init__(self, elem, Liste):
        self.element = elem
        self.vol = Liste[0]
        self.VH = Liste[1]
        self.RH = Liste[2]
        self.VK = Liste[3]
        self.RK = Liste[4]
        self.KL = Liste[5]
        self.HL = Liste[6]
        try:
            self.RaumNr = self.element.Space[phase].Number
            self.Raum = self.element.Space[phase].Id.ToString()
        except:
            self.RaumNr = ''
            self.Raum = ''
            logger.error('kein MEP-Raum für Element {} gefunden'.format(self.element.Id.ToString()))
        try:
            self.leistung_H = get_value(self.element.LookupParameter(self.HL))
        except:
            self.leistung_H = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.HL))
        try:
            self.leistung_K = get_value(self.element.LookupParameter(self.KL))
        except:
            self.leistung_K = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.KL))
        try:
            self.VLH = get_value(self.element.LookupParameter(self.VH))
        except:
            self.VLH = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.VH))
        try:
            self.RLH = get_value(self.element.LookupParameter(self.RH))
        except:
            self.RLH = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.RH))
        try:
            self.VLK = get_value(self.element.LookupParameter(self.VK))
        except:
            self.VLK = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.VK))
        try:
            self.RLK = get_value(self.element.LookupParameter(self.RK))
        except:
            self.RLK = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.RK))
        try:
            self.VOL = get_value(self.element.LookupParameter(self.vol))
        except:
            self.VOL = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.vol))
        


MEP_Dict = {}
desfamilieliste = Familie_Bearbeiten.keys()[:]

with forms.ProgressBar(title='{value}/{max_value} Deckensegel', cancellable=True, step=10) as pb:
    for n, des in enumerate(desfamilieliste):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(desfamilieliste))
        deckensegel = Deckensegel(des,Familie_Bearbeiten[des])
        if deckensegel.Raum:
            if deckensegel.Raum not in MEP_Dict.keys():
                MEP_Dict[deckensegel.Raum] = [deckensegel]
            else:
                MEP_Dict[deckensegel.Raum].append(deckensegel)

class MEPRaum:
    def __init__(self, elemid, liste_des):
        self.element = doc.GetElement(DB.ElementId(int(elemid)))
        self.des_liste = liste_des
        self.leistungH = 0
        self.leistungK = 0
        self.VLH_temp = -200
        self.RLH_temp = -200
        self.VLK_temp = -200
        self.RLK_temp = -200
        self.Vol = 0
        self.berechnen()


    def berechnen(self):
        liste_VLH = []
        liste_RLH = []
        liste_VLK = []
        liste_RLK = []
        for item in self.des_liste:
            self.leistungH += item.leistung_H
            self.leistungK += item.leistung_K
            self.Vol += item.VOL
            if item.VLH not in liste_VLH:
                liste_VLH.append(item.VLH)
            if item.RLH not in liste_RLH:
                liste_RLH.append(item.RLH)
            if item.VLK not in liste_VLK:
                liste_VLK.append(item.VLK)
            if item.RLK not in liste_RLK:
                liste_RLK.append(item.RLK)
        if len(liste_VLH) == 1:
            self.VLH_temp = liste_VLH[0]
        else:
            liste_VLH_neu = [float(a) for a in liste_VLH]
            self.VLH_temp = max(liste_VLH_neu)
            logger.error('Vorlauftemperatur Heizen von DeS in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
        if len(liste_RLH) == 1:
            self.RLH_temp = liste_RLH[0]
        else:
            liste_RLH_neu = [float(a) for a in liste_RLH]
            self.RLH_temp = min(liste_RLH_neu)
            logger.error('Rücklauftemperatur Heizen von DeS in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
        if len(liste_VLK) == 1:
            self.VLK_temp = liste_VLK[0]
        else:
            liste_VLK_neu = [float(a) for a in liste_VLK]
            self.VLK_temp = min(liste_VLK_neu)
            logger.error('Vorlauftemperatur Kälte von DeS in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
        if len(liste_RLK) == 1:
            self.RLK_temp = liste_RLK[0]
        else:
            liste_RLK_neu = [float(a) for a in liste_RLK]
            self.RLK_temp = max(liste_RLK_neu)
            logger.error('Rücklauftemperatur Kälte von DeS in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))

    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.element.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)
        
        wert_schreiben("IGF_H_DeS-VL_Win_Temp", self.VLH_temp)
        wert_schreiben("IGF_H_DeS-RL_Win_Temp", self.RLH_temp)
        wert_schreiben("IGF_K_DeS-VL_Som_Temp", self.VLK_temp)
        wert_schreiben("IGF_K_DeS-RL_Som_Temp", self.RLK_temp)
        wert_schreiben("IGF_H_DeS_Winter", self.Vol)
        wert_schreiben("IGF_K_DeS_Sommer", self.Vol)
        wert_schreiben("IGF_H_DeS_Leistung", self.leistungH)
        wert_schreiben("IGF_K_DeS_Leistung", self.leistungK)

mep_liste = []
MEP_Liste = MEP_Dict.keys()[:]
with forms.ProgressBar(title='{value}/{max_value} MEP-Räume',cancellable=True, step=10) as pb1:
    for n, mep in enumerate(MEP_Liste):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(MEP_Liste))
        mepraum = MEPRaum(mep,MEP_Dict[mep])
        mep_liste.append(mepraum)


# Werte zuückschreiben + Abfrage
if forms.alert("Berechnete Werte in MEP-Räume schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Daten von DeS in MEP-Räume schreiben')
        for n,mepraum in enumerate(mep_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mep_liste))
            mepraum.werte_schreiben()
        t.Commit()