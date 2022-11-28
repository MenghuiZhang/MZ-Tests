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

__title__ = "2.02 Daten von ULK in MEP schreiben"
__doc__ = """Temperatur und Leistung von Umluftkühler in MEP Raum schreiben
--------------------------
ULK Category: HLS-Bauteile
--------------------------
Parameter in MEP-Räume: 
--------------------------
IGF_K_ULK-VL_Temp: Vorlauftemperatur, 
ermittelt von System. 
Typparameter in System: 'Temperatur von Medium',  'Abkürzung': ('K VL')

IGF_K_ULK-RL_Temp: Rücklauftemperatur, 
ermittelt von System. 
Typparameter in System: 'Temperatur von Medium',  'Abkürzung': ('K RL')

IGF_K_ULK_Leistung: ULK-Kühlleistung, 
ermittelt von ULK-Bauteile

IGF_K_ULK_Wassermenge: [L/s] 
berechnungsformel: IGF_K_ULK_Leistung/4200/(IGF_K_ULK-RL_Temp-IGF_K_ULK-VL_Temp)
--------------------------

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
config = script.get_config(name+number+'ULK-Familie')

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
        self.Selectedindex = -1
        self.Paras = []
        self.checked = False

    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self, value):
        self._Name = value
    @property
    def Selectedindex(self):
        return self._Selectedindex
    @Selectedindex.setter
    def Selectedindex(self, value):
        self._Selectedindex = value
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

Liste_ULK = ObservableCollection[MechanicalEquipment]()

ULKName = HLS_Para_Dict.keys()[:]
ULKName.sort()

for elem in ULKName:
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
    Liste_ULK.Add(temp)


class ULK_Familie(forms.WPFWindow):
    def __init__(self, xaml_file_name,Liste_ULK):
        self.Liste_ULK = Liste_ULK
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.altdatagrid = Liste_ULK
        self.dataGrid.ItemsSource = Liste_ULK
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
            ULK_dict = config.ULK_dict
            for item in self.Liste_ULK:
                if item.Name in ULK_dict.keys():
                    try:
                        item.checked = ULK_dict[item.Name][0]
                    except:
                        pass
                    try:
                        item.Selectedindex = item.Para_id_dict[ULK_dict[item.Name][1]]
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
        for item in self.Liste_ULK:
            if item.checked:
                try:
                    dict_[item.Name] = [item.checked,item.Para_dict[item.Selectedindex]]
                except:
                    pass

        config.ULK_dict = dict_
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
    
ULK_Familie_Auswahl = ULK_Familie("window.xaml",Liste_ULK)
try:
    ULK_Familie_Auswahl.ShowDialog()
except Exception as e:
    ULK_Familie_Auswahl.Close()
    script.exit()

Familie_Bearbeiten = {}
for el in Liste_ULK:
    if el.checked:
        for elem in HLS_Dict[el.Name]:
            Familie_Bearbeiten[elem] = el.Para_dict[el.Selectedindex]

if len(Familie_Bearbeiten.keys()) == 0:
    script.exit()

class Umluftkuehler:
    def __init__(self, elem, param):
        self.element = elem
        self.param = param
        try:
            self.RaumNr = self.element.Space[phase].Number
            self.Raum = self.element.Space[phase].Id.ToString()
        except:
            self.RaumNr = ''
            self.Raum = ''
            logger.error('kein MEP-Raum für Element {} gefunden'.format(self.element.Id.ToString()))
        try:
            self.leistung = get_value(self.element.LookupParameter(self.param))
        except:
            self.leistung = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.param))
        self.VL_Temp = ''
        self.RL_Temp = ''
        self.Temperaturermittelung()
    
    def Temperaturermittelung(self):
        conns = self.element.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                system  = conn.MEPSystem
                if system:
                    typ = doc.GetElement(system.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsElementId())
                    try:
                        abk = typ.LookupParameter('Abkürzung').AsString()
                        temp = get_value(typ.LookupParameter('Temperatur von Medium'))
                        if abk == 'K VL':
                            self.VL_Temp = temp
                        elif abk == 'K RL':
                            self.RL_Temp = temp
                    except:
                        pass


MEP_Dict = {}
ulkfamilieliste = Familie_Bearbeiten.keys()[:]

with forms.ProgressBar(title='{value}/{max_value} Umluftkühler', cancellable=True, step=10) as pb:
    for n, ulk in enumerate(ulkfamilieliste):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(ulkfamilieliste))
        umluft = Umluftkuehler(ulk,Familie_Bearbeiten[ulk])
        if umluft.Raum:
            if umluft.Raum not in MEP_Dict.keys():
                MEP_Dict[umluft.Raum] = [umluft]
            else:
                MEP_Dict[umluft.Raum].append(umluft)

class MEPRaum:
    def __init__(self, elemid, liste_ulk):
        self.element = doc.GetElement(DB.ElementId(int(elemid)))
        self.ulk_liste = liste_ulk
        self.leistung = 0
        self.VL_temp = -200
        self.RL_temp = -200
        self.Wasser = 0
        self.leistungberechnen()
        self.Tempberechnen()
        self.Wassermengenberechnen()

    def leistungberechnen(self):
        for item in self.ulk_liste:
            self.leistung += item.leistung
    def Tempberechnen(self):
        liste_VL = []
        liste_RL = []
        for item in self.ulk_liste:
            if item.VL_Temp not in liste_VL:
                liste_VL.append(item.VL_Temp)
            if item.RL_Temp not in liste_RL:
                liste_RL.append(item.RL_Temp)
        if len(liste_VL) == 1:
            self.VL_Temp = liste_VL[0]
        else:
            logger.error('Vorlauftemperatur von ULK in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
            for el in self.ulk_liste:
                logger.error('Vorlauftemperatur: {} °C, ULKId: {}'.format(el.VL_Temp,el.element.Id.ToString()))
            liste_VL_neu = [float(a) for a in liste_VL]
            self.VL_Temp = min(liste_VL_neu)
            logger.error('!!! Achtung: ULK-Vorlauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.element.Number,self.element.Id.ToString(),self.VL_Temp))
            

        if len(liste_RL) == 1:
            self.RL_Temp = liste_RL[0]
        else:
            logger.error('Rücklauftemperatur von ULK in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
            for el in self.ulk_liste:
                logger.error('Rücklauftemperatur: {} °C, ULKId: {}'.format(el.RL_Temp,el.element.Id.ToString()))
            liste_RL_neu = [float(a) for a in liste_RL]
            self.RL_Temp = max(liste_RL_neu)
            logger.error('!!! Achtung: ULK-Vorlauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.element.Number,self.element.Id.ToString(),self.RL_Temp))
            
    def Wassermengenberechnen(self):
        if self.VL_temp == -200 or self.RL_Temp == -200:
            self.VL_temp = 0
            self.RL_Temp = 0
            return
        else:
            self.Wasser = self.leistung / 4200.0 / abs(self.RL_Temp - self.VL_Temp)
    
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

        
        wert_schreiben("IGF_K_ULK-VL_Temp", self.VL_Temp)
        wert_schreiben("IGF_K_ULK-RL_Temp", self.RL_Temp)
        wert_schreiben("IGF_K_ULK_Leistung", self.leistung)
        wert_schreiben("IGF_K_ULK_Wassermenge", self.Wasser)
    

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
        t.Start('Daten von ULK in MEP-Räume schreiben')
        for n,mepraum in enumerate(mep_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mep_liste))
            mepraum.werte_schreiben()
        t.Commit()