# coding: utf8
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
from IGF_lib import get_value,Muster_Pruefen
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Text.RegularExpressions import Regex
from System.Collections.Generic import List
from clr import GetClrType

__title__ = "Wirkungsort"
__doc__ = """
Kühlleistung summieren
Kühllast und instalierte Kühlleistung abgleichen

Parameter:
IGF_RLT_ZuluftminRaum: Zuluftmengen
IGF_RLT_ZuluftTemperatur: Zulufttemperatur
LIN_BA_OVERFLOW_SUPPLY_AIR_TEMPERATURE: Zulufttemperatur falls IGF_RLT_ZuluftTemperatur nicht eingegeben wird
LIN_BA_DESIGN_COOLING_TEMPERATURE: Raumtemperatur
IGF_K_KühllastLaborRaum: Kühllast Labor Raum
IGF_S_KühllastLaborPWK: Kühllast für Laboreinrichtung über PKW
LIN_BA_CALCULATED_COOLING_LOAD: Kühllast Gebäude
IGF_K_DeS_Leistung: Kühlleistung DeS
IGF_K_ULK_Leistung: Kühlleistung ULK
IGF_K_KA_Leistung: sonstige Kühlleistung
IGF_RLT_ZuluftKühlleistung: Kühlleistung Zuluft, Zuluftfaktor * (Vol_zu * 1000 * 1.2 * 1.006 * (Temp_Raum - Temp_Zu) / 3600)
IGF_K_KühlleistungRaum: Summe von Zuluft- & DeS- & ULK- & Kältekühlleistung
IGF_K_KühllastGesamt: Summe von Kühllast Gebäude und Kühllast Labor Raum
IGF_K_KühlleistungBilanz: gesamte Kühlleistung - gesamte Kühllast
IGF_K_KühlBilanzProzent: gesamte Kühlleistung / gesamte Kühllast


[Version: 2.0]
[2022.12.09]
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'Kuehlen-Familie')

try:getlog(__title__)
except:pass

class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

class Endverbraucher(TemplateItemBase):
    def __init__(self,name,elems,category):
        TemplateItemBase.__init__(self)
        self.familyname = name
        self.category = category
        self._Selectedort = -1
        self._RV = False
        self._Sechswege = False
        self._VSR = False
        self.VSR_enabled = False
        self.RV_enabled = False
        self.liste_ort = ['Vorlauf','Rücklauf']
        if self.category == 'Luftdurchlässe':
            self.Orts = []
            self.VSR = True
            self.VSR_enabled = True
        else:
            self.VSR = False
            self.RV_enabled = True
            self.Orts = sorted(self.liste_ort)
        self.elems = elems
        if len(self.elems) == 0:
            self.info = 'Typ nicht verwendet'
        else:
            self.info = 'Typ bereits verwendet'
    
    @property
    def RV(self):
        return self._RV
    @RV.setter
    def RV(self,value):
        if value != self._RV:
            self._RV = value
            self.RaisePropertyChanged('RV')
    @property
    def Sechswege(self):
        return self._Sechswege
    @Sechswege.setter
    def Sechswege(self,value):
        if value != self._Sechswege:
            self._Sechswege = value
            self.RaisePropertyChanged('Sechswege')
    @property
    def VSR(self):
        return self._VSR
    @VSR.setter
    def VSR(self,value):
        if value != self._VSR:
            self._VSR = value
            self.RaisePropertyChanged('VSR')
    @property
    def Selectedort(self):
        return self._Selectedort
    @Selectedort.setter
    def Selectedort(self,value):
        if value != self._Selectedort:
            self._Selectedort = value
            self.RaisePropertyChanged('Selectedort')

class Regelkomponent(TemplateItemBase):
    def __init__(self,name,category):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.familyname = name
        self.category = category
        self._Selectedart = -1

        self.liste_art1 = ['Regelventil','6-Wege-Ventil']
        self.liste_art2 = ['VSR']
        if self.category == 'Rohrzubehör':
            self.Arts = sorted(self.liste_art1)
        else:
            self.Arts = self.liste_art2 
   
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    @property
    def Selectedart(self):
        return self._Selectedart
    @Selectedart.setter
    def Selectedart(self,value):
        if value != self._Selectedart:
            self._Selectedart = value
            self.RaisePropertyChanged('Selectedart')

AUSWAHL_EV_IS = ObservableCollection[Endverbraucher]()
AUSWAHL_RK_IS = ObservableCollection[Regelkomponent]()

def get_IS():
    Dict = {}
    filter_EV = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_MechanicalEquipment,DB.BuiltInCategory.OST_DuctTerminal]))
    HLSs = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter_EV).ToElements()
    for el in HLSs:
        category = el.Category.Name
        FamilyName = el.Symbol.FamilyName + ': ' + el.Name
        if category not in Dict.keys():
            Dict[category] = {}   
        if FamilyName not in Dict[category].keys():
            Dict[category][FamilyName] = [el.Id]            
        else:Dict[category][FamilyName].append(el.Id)
    
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Dict_EV = {}
    Dict_RK = {}
    for el in Families:
        if el.FamilyCategoryId.IntegerValue in [-2001140,-2008013]:
            category = el.FamilyCategory.Name
            if category not in Dict_EV.keys():
                Dict_EV[category] = []
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Dict_EV[category]:
                    Dict_EV[category].append(typname)
        elif el.FamilyCategoryId.IntegerValue in [-2008016,-2008055]:
            category = el.FamilyCategory.Name
            famname = el.Name
            if category not in Dict_RK.keys():
                Dict_RK[category] = []
            if famname not in Dict_RK[category]:
                Dict_RK[category].append(famname)         
    
    for category in sorted(Dict_EV.keys()):
        for fam in sorted(Dict_EV[category]):
            if category in Dict.keys():
                if fam in Dict[category].keys():
                    AUSWAHL_EV_IS.Add(Endverbraucher(fam,Dict[category][fam],category))
                else:
                    AUSWAHL_EV_IS.Add(Endverbraucher(fam,[],category))
            else:
                AUSWAHL_EV_IS.Add(Endverbraucher(fam,[],category))

    for category in sorted(Dict_RK.keys()):
        for fam in sorted(Dict_RK[category]):
            AUSWAHL_RK_IS.Add(Regelkomponent(fam,category))
    
get_IS()

class FamilieEinstellen(forms.WPFWindow):
    def __init__(self):
        self.Liste_EV = AUSWAHL_EV_IS
        self.Liste_RK = AUSWAHL_RK_IS
        forms.WPFWindow.__init__(self, "window.xaml",handle_esc=False)
        self.tempcoll_EV = ObservableCollection[Endverbraucher]()
        self.tempcoll_RK = ObservableCollection[Regelkomponent]()
        self.datagrid_EV.ItemsSource = self.Liste_EV
        self.datagrid_RK.ItemsSource = self.Liste_RK
        self.read_config()
        self.start = False
        
    def read_config(self):
        try:
            for item in self.Liste_EV:
                if item.familyname in config.Endverbraucher.keys():
                    configdatei = config.Endverbraucher[item.familyname]
                    item.RV = configdatei[0]
                    item.Sechswege = configdatei[1]
                    item.VSR = configdatei[2]
                    if configdatei[3] not in ['',None]:item.Selectedort = item.Orts.index(configdatei[3])
        except:pass
        try:
            for item in self.Liste_RK:
                if item.familyname in config.Regelkomponent.keys():
                    configdatei = config.Regelkomponent[item.familyname]
                    item.checked = configdatei[0]
                    if configdatei[1] not in ['',None]:item.Selectedart = item.Arts.index(configdatei[1])
        except:pass
           
       
    def write_config(self):
        _dict_EV = {}
        _dict_RK = {}
        try:
            for item in self.Liste_EV:
                if item.Selectedort == -1:
                    _dict_EV[item.familyname] = [item.RV,item.Sechswege,item.VSR,'']
                else:
                    _dict_EV[item.familyname] = [item.RV,item.Sechswege,item.VSR,item.Orts[item.Selectedort]]
                
        except:pass
        try:
            for item in self.Liste_RK:
                if item.Selectedart == -1:
                    _dict_RK[item.familyname] = [item.checked,'']
                else:
                    _dict_RK[item.familyname] = [item.checked,item.Arts[item.Selectedart]]
        except:pass
        config.Regelkomponent = _dict_RK
        config.Endverbraucher = _dict_EV

        script.save_config()

    def checkedchanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.datagrid_RK.SelectedIndex != -1:
            if item in self.datagrid_RK.SelectedItems:
                for el in self.datagrid_RK.SelectedItems:el.checked = checked

    def art_select_changed(self,sender,e):
        item = sender.DataContext
        artindex = sender.SelectedIndex
        if artindex != -1:
            art = item.Arts[artindex]
            if self.datagrid_RK.SelectedIndex != -1:
                if item in self.datagrid_RK.SelectedItems:
                    for el in self.datagrid_RK.SelectedItems:
                        if art in el.Arts:
                            el.Selectedart = el.Arts.index(art)
    
    def VSR_checkedchanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.datagrid_EV.SelectedIndex != -1:
            if item in self.datagrid_EV.SelectedItems:
                for el in self.datagrid_EV.SelectedItems:
                    if el.category == 'Luftdurchlässe':
                        el.VSR = checked
    
    def RV_checkedchanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.datagrid_EV.SelectedIndex != -1:
            if item in self.datagrid_EV.SelectedItems:
                for el in self.datagrid_EV.SelectedItems:
                    if el.category == 'HLS-Bauteile':
                        el.RV = checked
    
    def Sechswege_checkedchanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.datagrid_EV.SelectedIndex != -1:
            if item in self.datagrid_EV.SelectedItems:
                for el in self.datagrid_EV.SelectedItems:
                    if el.category == 'HLS-Bauteile':
                        el.Sechswege = checked

    def ort_select_changed(self,sender,e):
        item = sender.DataContext
        ortindex = sender.SelectedIndex
        if ortindex != -1:
            ort = item.Orts[ortindex]
            if self.datagrid_EV.SelectedIndex != -1:
                if item in self.datagrid_EV.SelectedItems:
                    for el in self.datagrid_EV.SelectedItems:
                        if ort in el.Orts:
                            el.Selectedort = el.Orts.index(ort)

    def EV_suchechanged(self,sender,e):
        text = sender.Text
        if not text:
            self.datagrid_EV.ItemsSource = self.Liste_EV
            return
        self.tempcoll_EV.Clear()
        for item in self.Liste_EV:
            if item.familyname.upper().find(text.upper()) != -1:
                self.tempcoll_EV.Add(item)
        self.datagrid_EV.ItemsSource = self.tempcoll_EV 

    def RK_suchechanged(self,sender,e):
        text = sender.Text
        if not text:
            self.datagrid_RK.ItemsSource = self.Liste_RK
            return
        self.tempcoll_RK.Clear()
        for item in self.Liste_RK:
            if item.familyname.upper().find(text.upper()) != -1:
                self.tempcoll_RK.Add(item)
        self.datagrid_RK.ItemsSource = self.tempcoll_RK 

    def pruefen_EV(self,item):
        if item.RV or item.Sechswege:
            if item.Selectedort == -1:
                return "Entverbraucher: {} - Ort nicht definiert".format(item.familyname)
            else:
                return None
    
    def pruefen_RK(self,item):
        if item.checked:
            if item.Selectedart == -1:
                return "Regelkomponent: {} - Art nicht definiert".format(item.familyname)
            else:
                return None

    def ok(self,sender,args):
        for el in self.Liste_EV:
            text = self.pruefen_EV(el)
            if text:
                UI.TaskDialog.Show('Fehler',text)
                return
        for el in self.Liste_RK:
            text = self.pruefen_RK(el)
            if text:
                UI.TaskDialog.Show('Fehler',text)
                return      
        
        self.write_config()
        self.start = True
        self.Close()

    def close(self,sender,args):
        self.Close()
    
Familie_Auswahl = FamilieEinstellen()
try:
    Familie_Auswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    Familie_Auswahl.Close()
    script.exit()

if Familie_Auswahl.start == False:
    script.exit()

Liste_VSR = [el.familyname for el in AUSWAHL_RK_IS if el.checked and el.Selectedart != -1 and el.Arts[el.Selectedart] == 'VSR']
Liste_RV = [el.familyname for el in AUSWAHL_RK_IS if el.checked and el.Selectedart != -1 and el.Arts[el.Selectedart] == 'Regelventil']
Liste_6_Wege = [el.familyname for el in AUSWAHL_RK_IS if el.checked and el.Selectedart != -1 and el.Arts[el.Selectedart] == '6-Wege-Ventil']
Liste_EV = [el for el in AUSWAHL_EV_IS if el.RV or el.Sechswege or el.VSR]

class EndverbraucherInstance:
    def __init__(self, elemid, family):
        self.family = family
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.ismuster = Muster_Pruefen(self.elem)
        if self.ismuster:return
        try:
            self.raumnummer = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
        except:
            self.raumnummer = ''
        
        self.vsr_geregelt = self.family.VSR
        self.rv_geregelt = self.family.RV
        self.sechswege_geregelt = self.family.Sechswege

        self.list = []
        self.vsr = None
        self.rv = None
        self.wege_6 = None
        if self.family.Selectedort != -1:
            self.ort = self.family.Orts[self.family.Selectedort]
        else:
            self.ort = ''
        
        self.regelkomponent_ermitteln()


    def reglerermitteln_VSR(self, elem):
        '''Ermittlung des VSRs Wirkungsorte'''
        if self.vsr:
            return
        if len(self.list) > 500:return 
        id = elem.Id.IntegerValue
        self.list.append(id)
        cate = elem.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Rohrdämmung', 'Luftkanaldämmung außen','Luftkanaldämmung innen']:
            conns = None
            try:
                conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:
                    conns = elem.ConnectorManager.Connectors
                except:
                    pass

            if conns:
                if conns.Size > 8:
                    self.vsr = None
                    return
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner

                        if not owner.Id.IntegerValue in self.list:
                            if owner.Category.Name == 'Luftkanalzubehör':
                                faminame = owner.Symbol.FamilyName
                                if faminame in Liste_VSR:
                                    self.vsr = owner.Id.IntegerValue
                                    return
                            self.reglerermitteln_VSR(owner)
    
    def reglerermitteln_RV(self, elem):
        '''Ermittlung des RVs Wirkungsorte'''
        if self.rv:
            return
        if len(self.list) > 500:return 
        elemid = elem.Id.IntegerValue
        self.list.append(elemid)
        cate = elem.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Rohrdämmung', 'Luftkanaldämmung außen','Luftkanaldämmung innen']:
            conns = None
            try:
                conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:
                    conns = elem.ConnectorManager.Connectors
                except Exception as e:
                    print(e)

            if conns:
                for conn in conns:
       
                    if elemid == self.elemid.IntegerValue:
                
                        if self.ort == 'Rücklauf':
                            try:
                                if conn.PipeSystemType.ToString() != 'ReturnHydronic':
                                    continue
                            except:
                                continue
                        elif self.ort == 'Vorlauf':
                            try:
                                if conn.PipeSystemType.ToString() != 'SupplyHydronic':
                                    continue
                            except:
                                continue

                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner

                        if not owner.Id.IntegerValue in self.list:
                            if owner.Category.Name == 'HLS-Bauteile':
                                return
                            if owner.Category.Name == 'Rohrzubehör':
                                faminame = owner.Symbol.FamilyName
                                if faminame in Liste_RV:
                                    self.rv = owner.Id.IntegerValue
                                    return

                            self.reglerermitteln_RV(owner)

    def reglerermitteln_6_wege(self, elem):
        '''Ermittlung des RVs Wirkungsorte'''
        if self.wege_6:
            return
        if len(self.list) > 500:return 
        elemid = elem.Id.IntegerValue
        self.list.append(elemid)
        cate = elem.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Rohrdämmung', 'Luftkanaldämmung außen','Luftkanaldämmung innen']:
            conns = None
            try:
                conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:
                    conns = elem.ConnectorManager.Connectors
                except:
                    pass

            if conns:
                for conn in conns:
                    if elemid == self.elemid.IntegerValue:
                        if self.ort == 'Rücklauf':
                            try:
                                if conn.PipeSystemType.ToString() != 'ReturnHydronic':
                                    continue
                            except:
                                continue
                        elif self.ort == 'Vorlauf':
                            try:
                                if conn.PipeSystemType.ToString() != 'SupplyHydronic':
                                    continue
                            except:
                                continue

                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner

                        if not owner.Id.IntegerValue in self.list:
                            if owner.Category.Name == 'HLS-Bauteile':
                                return
                            if owner.Category.Name == 'Rohrzubehör':
                                faminame = owner.Symbol.FamilyName
                                if faminame in Liste_6_Wege:
                                    self.wege_6 = owner.Id.IntegerValue
                                    return

                            self.reglerermitteln_6_wege(owner)

    def regelkomponent_ermitteln(self):
        if self.vsr_geregelt:
            self.reglerermitteln_VSR(self.elem)
        else:
            if self.rv_geregelt:
                self.reglerermitteln_RV(self.elem)
            if self.sechswege_geregelt:
                self.list = []
                self.reglerermitteln_6_wege(self.elem)

class RegelkomponentInstance:
    def __init__(self, elemid,elems):
        self.elems = elems
        self.elemid = elemid
        self.elem = doc.GetElement(DB.ElementId(self.elemid))
        self.ismuster = Muster_Pruefen(self.elem)

        if self.ismuster:return
        try:
            self.raumnummer = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
        except:
            self.raumnummer = ''
        self.wirkungsort = ''
        self.get_wirkungsort()
        
    def get_wirkungsort(self):
        text = ''
        for el in self.elems:
            if el.raumnummer:
                text += el.raumnummer + ', '
        if text:
            self.wirkungsort = text[:-2]
        else:
            self.wirkungsort = ''
        
    def werte_schreiben(self):
        try:self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raumnummer)
        except Exception as e:logger.error(e)
        try:self.elem.LookupParameter('IGF_X_Wirkungsort').Set(self.wirkungsort)
        except Exception as e:logger.error(e)

regel_komponent_dict = {}        
with forms.ProgressBar(title='{value}/{max_value} Kälte', cancellable=True, step=10) as pb:
    for n, familie in enumerate(Liste_EV):
        pb.title='{value}/{max_value} Exemplare von ' + familie.familyname + ' --- ' + str(n+1) + ' / '+ str(len(Liste_EV)) + 'Familien'
        for n1, elemid in enumerate(familie.elems):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n1 + 1, len(familie.elems))
            bauteil = EndverbraucherInstance(elemid,familie)
            if not bauteil.ismuster:
                if bauteil.raumnummer:
                    if bauteil.vsr:
                        if bauteil.vsr not in regel_komponent_dict.keys():
                            regel_komponent_dict[bauteil.vsr] = [bauteil]
                        else:
                            regel_komponent_dict[bauteil.vsr].append(bauteil)
                    if bauteil.rv:
                        if bauteil.rv not in regel_komponent_dict.keys():
                            regel_komponent_dict[bauteil.rv] = [bauteil]
                        else:
                            regel_komponent_dict[bauteil.rv].append(bauteil)
                    if bauteil.wege_6:
                        if bauteil.wege_6 not in regel_komponent_dict.keys():
                            regel_komponent_dict[bauteil.wege_6] = [bauteil]
                        else:
                            regel_komponent_dict[bauteil.wege_6].append(bauteil)

Liste_RV_Final = []
with forms.ProgressBar(title='{value}/{max_value} Regelkomponent',cancellable=True, step=10) as pb1:
    for n, elemid in enumerate(regel_komponent_dict.keys()):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(regel_komponent_dict.keys()))
        rv = RegelkomponentInstance(elemid,regel_komponent_dict[elemid])
        if not rv.ismuster:Liste_RV_Final.append(rv)

# Werte zurückschreiben + Abfrage
if forms.alert("Einbauort & Wirkungsort an Regelkomponenten schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Regelkomponenten",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Regelkomponent')
        for n,bauteil_rv in enumerate(Liste_RV_Final):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(Liste_RV_Final))
            bauteil_rv.werte_schreiben()
        t.Commit()
