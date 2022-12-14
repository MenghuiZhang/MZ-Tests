# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,ExternalEvent,TaskDialogCommonButtons,TaskDialogResult 
from System.Collections.ObjectModel import ObservableCollection
import System
from rpw import revit,DB,UI
from pyrevit import script,forms
from System.Text.RegularExpressions import Regex
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Windows.Media import Brushes
from System.Collections.Generic import List
from clr import GetClrType

RED = Brushes.Red
BLACK = Brushes.Black
GRAY = Brushes.Gray

def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if not param:return ''
    if param.StorageType.ToString() == 'ElementId':
        return param.AsValueString()
    elif param.StorageType.ToString() == 'Integer':
        value = param.AsInteger()
    elif param.StorageType.ToString() == 'Double':
        value = param.AsDouble()
    elif param.StorageType.ToString() == 'String':
        value = param.AsString()
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'HK-Bauteile')

DICT_MEP_NUMMER_NAME = {}

def get_MEP_NUMMER_NAMR():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        DICT_MEP_NUMMER_NAME[el.Number] = el.Number + ' - ' + el.LookupParameter('Name').AsString()
    
get_MEP_NUMMER_NAMR()

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

class Familie(TemplateItemBase):
    def __init__(self,Familie,elems,Symbol):
        TemplateItemBase.__init__(self)
        self.Familiename = Familie
        self._checked = False

        self._Nennleistung = 0
        
        self._Nenntemperatur = 50
        self._Exponent = 1.3

        self._visibility = 1

        self._typ = ''
        self._size = ''

        self.elems = elems
        self.symbol = Symbol
        self._beschreibung = ''
        try:self.get_daten()
        except:pass
        if len(self.elems) == 0:
            self.info = 'Typ nicht verwendet'
        else:
            self.info = 'Typ bereits verwendet'
    
    def werte_schreiben_beschrifen(self):
        try:self.symbol.LookupParameter('Typenkommentare').Set(self.typ + ', '+self.size)
        except:pass
        try:self.symbol.LookupParameter('Beschreibung').Set(self.beschreibung)
        except:pass
    
    def werte_schreiben_Leistung(self):
        try:self.symbol.LookupParameter('IGF_H_HK-Nennleistung').SetValueString(str(self.Nennleistung))
        except:pass
        try:self.symbol.LookupParameter('IGF_H_HK-Nennübertemperatur').SetValueString(str(self.Nenntemperatur))
        except:pass
        try:self.symbol.LookupParameter('IGF_H_HK-Exponent').SetValueString(str(self.Exponent).replace(',','.'))
        except:pass

    def get_daten(self):
        self.Nennleistung = round(get_value(self.symbol.LookupParameter('IGF_H_HK-Nennleistung')))
        self.Nenntemperatur = round(get_value(self.symbol.LookupParameter('IGF_H_HK-Nennübertemperatur')))
        self.Exponent = round(get_value(self.symbol.LookupParameter('IGF_H_HK-Exponent')),4)
        typ = self.symbol.LookupParameter('Typenkommentare').AsString()
        self.beschreibung = self.symbol.LookupParameter('Beschreibung').AsString()
        if typ:
            if typ.find(', ') != -1:
                liste = typ.split(', ')
                self.typ = liste[0]
                self.size = liste[1]
            else:
                try:
                    h = int(get_value(self.symbol.LookupParameter('MC Height')))
                    b = int(get_value(self.symbol.LookupParameter('MC Width')))
                    l = int(get_value(self.symbol.LookupParameter('MC Length')))
                    self.size = "{}x{}x{}lg".format(h,b,l)
                except:pass
        else:
            try:
                h = int(get_value(self.symbol.LookupParameter('MC Height')))
                b = int(get_value(self.symbol.LookupParameter('MC Width')))
                l = int(get_value(self.symbol.LookupParameter('MC Length')))
                self.size = "{}x{}x{}lg".format(h,b,l)
            except:pass

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    
    @property
    def visibility(self):
        return self._visibility
    @visibility.setter
    def visibility(self,value):
        if value != self._visibility:
            self._visibility = value
            self.RaisePropertyChanged('visibility')

    @property
    def Nennleistung(self):
        return self._Nennleistung
    @Nennleistung.setter
    def Nennleistung(self,value):
        if value != self._Nennleistung:
            self._Nennleistung = value
            self.RaisePropertyChanged('Nennleistung')

    @property
    def Nenntemperatur(self):
        return self._Nenntemperatur
    @Nenntemperatur.setter
    def Nenntemperatur(self,value):
        if value != self._Nenntemperatur:
            self._Nenntemperatur = value
            self.RaisePropertyChanged('Nenntemperatur')

    @property
    def Exponent(self):
        return self._Exponent
    @Exponent.setter
    def Exponent(self,value):
        if value != self._Exponent:
            self._Exponent = value
            self.RaisePropertyChanged('Exponent')

    @property
    def typ(self):
        return self._typ
    @typ.setter
    def typ(self,value):
        if value != self._typ:
            self._typ = value
            self.RaisePropertyChanged('typ')
    
    @property
    def size(self):
        return self._size
    @size.setter
    def size(self,value):
        if value != self._size:
            self._size = value
            self.RaisePropertyChanged('size')
    
    
    
    @property
    def beschreibung(self):
        return self._beschreibung
    @beschreibung.setter
    def beschreibung(self,value):
        if value != self._beschreibung:
            self._beschreibung = value
            self.RaisePropertyChanged('beschreibung')

AUSWAHL_HEIZKOERPER_IS = ObservableCollection[Familie]()

def get_Heizkoeper_IS():
    Dict = {}
    HLSs = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
        .WhereElementIsNotElementType().ToElements()
    for el in HLSs:
        FamilyName = el.Symbol.FamilyName + ': ' + el.Name
        if FamilyName not in Dict.keys():
            Dict[FamilyName] = [el.Id.ToString()]            
        else:Dict[FamilyName].append(el.Id.ToString())
    
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Dict1 = {}
    for el in Families:
        if el.FamilyCategoryId.IntegerValue == -2001140:
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Dict1.keys():
                    Dict1[typname] = typ
    
    for fam in sorted(Dict1.keys()):
        if fam in Dict.keys():
            AUSWAHL_HEIZKOERPER_IS.Add(Familie(fam,Dict[fam],Dict1[fam]))
        else:
            AUSWAHL_HEIZKOERPER_IS.Add(Familie(fam,[],Dict1[fam]))
    
get_Heizkoeper_IS()

class Familienauswahl(forms.WPFWindow):
    def __init__(self):
        self.HLS_IS = AUSWAHL_HEIZKOERPER_IS
        self.read_config()
        self.regex1 = Regex("[^0-9.]+")
        self.regex2 = Regex("[^0-9]+")
        self.temp_coll = ObservableCollection[Familie]()
        self.temp_coll1 = ObservableCollection[Familie]()
        forms.WPFWindow.__init__(self, 'GUI_Heizkoerper.xaml',handle_esc=False)
        self.lv_HK.ItemsSource = self.HLS_IS
        self.lv_HK1.ItemsSource = self.HLS_IS
        
        self.result = False
    
    def textinput(self, sender, args):
        try:
            if sender.Text in ['',None]:
                args.Handled = self.regex1.IsMatch(args.Text)
            elif sender.Text.find('.') != -1 and args.Text == '.':
                args.Handled = True
            else:
                args.Handled = self.regex1.IsMatch(args.Text)
        except:
            args.Handled = True

    def textinput1(self, sender, args):
        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True
    
    # def textchanged(self,sender,liste0,liste1):
    #     text = sender.Text
    #     if not text:
    #         liste0.ItemsSource = self.HLS_IS
    #         return 
    #     liste1.Clear()
    #     for el in self.HLS_IS:
    #         if el.Familiename.upper().find(text.upper()) != -1:
    #             liste1.Add(el)
    #     liste0.ItemsSource = liste1
    
    def suchechanged(self,sender):
        text = sender.Text
        self.suche1.Text = text
        self.suche0.Text = text

    def checkboxchanged(self,sender):
        checked = sender.IsChecked
        self.checkbox2.IsChecked = checked
        self.checkbox1.IsChecked = checked

    def checkedboxorsuchechanged(self):
        text = self.suche1.Text
        checked = self.checkbox1.IsChecked
        self.temp_coll.Clear()
        if not text:
            if checked:
                for el in self.HLS_IS:
                    if el.checked:
                        self.temp_coll.Add(el)
                self.lv_HK.ItemsSource = self.temp_coll
                self.lv_HK1.ItemsSource = self.temp_coll
            else:
                self.lv_HK.ItemsSource = self.HLS_IS
                self.lv_HK1.ItemsSource = self.HLS_IS
            return 
        else:
            if checked:
                for el in self.HLS_IS:
                    if el.checked and el.Familiename.upper().find(text.upper()) != -1:
                        self.temp_coll.Add(el)
                
            else:
                for el in self.HLS_IS:
                    if el.Familiename.upper().find(text.upper()) != -1:
                        self.temp_coll.Add(el)
            self.lv_HK.ItemsSource = self.temp_coll
            self.lv_HK1.ItemsSource = self.temp_coll

    def familiecheckedchanged(self,sender,e):
        self.checkboxchanged(sender)
        self.checkedboxorsuchechanged()
        
    def textchanged(self,sender,e):
        self.suchechanged(sender)
        self.checkedboxorsuchechanged()
    
    # def textchanged1(self,sender,e):
    #     self.textchanged(sender,self.lv_HK1,self.temp_coll1)
        
    def read_config(self):
        try:
            Liste = config.HK_Familie
            for el in self.HLS_IS:
                if el.Familiename in Liste:
                    el.checked = True
        except:pass
    
    def write_config(self):
        try:
            Liste = []
            for el in self.HLS_IS:
                if el.checked:
                    Liste.append(el.Familiename)
            config.HK_Familie = Liste
        except:
            pass
        script.save_config()

    def cancel(self,sender,args):
        self.Close()
        self.result = False
    
    def Beschriften(self,sender,args):
        t = DB.Transaction(doc,'Beschriften')
        t.Start()
        for el in self.HLS_IS:
            if el.checked:
                el.werte_schreiben_beschrifen()
        t.Commit()
        t.Dispose()
        TaskDialog.Show('Info','Erledigt!')
      
        self.result = False

    def OK(self,sender,args):
        self.write_config()
        self.result = True
        self.Close()
        t = DB.Transaction(doc,'Leistung')
        t.Start()
        for el in self.HLS_IS:
            if el.checked:
                el.werte_schreiben_Leistung()
        t.Commit()
        t.Dispose()

    def checkedchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv_HK.SelectedIndex != -1:
            if item in self.lv_HK.SelectedItems:
                for el in self.lv_HK.SelectedItems:el.checked = checked
    
    def checkedchanged1(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv_HK1.SelectedIndex != -1:
            if item in self.lv_HK1.SelectedItems:
                for el in self.lv_HK1.SelectedItems:el.checked = checked

    def typchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK1.SelectedItem is not None:
            if Item in self.lv_HK1.SelectedItems:
                for item in self.lv_HK1.SelectedItems:
                    item.typ = text 

    def sizechanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK1.SelectedItem is not None:
            if Item in self.lv_HK1.SelectedItems:
                for item in self.lv_HK1.SelectedItems:
                    item.size = text  
    
    def beschreibungchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK1.SelectedItem is not None:
            if Item in self.lv_HK1.SelectedItems:
                for item in self.lv_HK1.SelectedItems:
                    item.beschreibung = text  

    def nlchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK.SelectedItem is not None:
            if Item in self.lv_HK.SelectedItems:
                for item in self.lv_HK.SelectedItems:
                    item.Nennleistung = text    
    
    def ntchanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK.SelectedItem is not None:
            if Item in self.lv_HK.SelectedItems:
                for item in self.lv_HK.SelectedItems:
                    item.Nenntemperatur = text 

    def nechanged(self,sender,args):
        Item = sender.DataContext
        text = sender.Text
        if self.lv_HK.SelectedItem is not None:
            if Item in self.lv_HK.SelectedItems:
                for item in self.lv_HK.SelectedItems:
                    item.Exponent = text 
   
FamilienDialog = Familienauswahl()
try:FamilienDialog.ShowDialog()
except Exception as e:
    logger.error(e)
    FamilienDialog.Close()
    script.exit()

class HeizkoerperInstance(TemplateItemBase):
    def __init__(self,elemid,Familie):
        TemplateItemBase.__init__(self)
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.Familie = Familie
        self.Familie_Name = self.Familie.Familiename
        self._fore = BLACK
        
        self._sollleistung = 0 

        self.F_Nennleistung = self.Familie.Nennleistung
        self.F_NennTemperatur = self.Familie.Nenntemperatur
        self.F_NennExponent = self.Familie.Exponent

        self.HL_Param = self.elem.get_Parameter(System.Guid('49d7278f-3e2b-4f2e-b619-43e856e15be7'))
        self.SL_Param = self.elem.get_Parameter(System.Guid('ab333295-9f54-40bd-ae7b-1fb1f1fc37cc'))

        self.phase = doc.GetElement(self.elem.CreatedPhaseId)
        self.ismuster =  self.Muster_Pruefen(self.elem)
        self.raum = ''

        self._Nennleistung = 0
        self._Heizleistung = 0

        self.vorlauf = 0
        self.ruecklauf = 0
        self.Raumtemp = 20
        self.uebertemp = 0

        self.get_Temperatur()
        self.GetRaum()

        self.Heizleistung = self.get_value(self.HL_Param)
        self.sollleistung = self.get_value(self.SL_Param)
      
    
    def farbe(self):
        if (self.Heizleistung > self.Nennleistung) or (self.sollleistung != self.Heizleistung):
            self.fore = RED
        else:
            self.fore = BLACK
    
    @property
    def sollleistung(self):
        return self._sollleistung
    @sollleistung.setter
    def sollleistung(self,value):
        if value != self._sollleistung:
            self._sollleistung = value
            self.RaisePropertyChanged('sollleistung')

    
    @property
    def Nennleistung(self):
        return self._Nennleistung
    @Nennleistung.setter
    def Nennleistung(self,value):
        if value != self._Nennleistung:
            self._Nennleistung = value
            self.RaisePropertyChanged('Nennleistung')
    
    @property
    def fore(self):
        return self._fore
    @fore.setter
    def fore(self,value):
        if value != self._fore:
            self._fore = value
            self.RaisePropertyChanged('fore')
    
    @property
    def Heizleistung(self):
        return self._Heizleistung
    @Heizleistung.setter
    def Heizleistung(self,value):
        if value != self._Heizleistung:
            self._Heizleistung = value
            self.RaisePropertyChanged('Heizleistung')
 
    def get_uebertemp(self):
        self.uebertemp = (self.vorlauf+self.ruecklauf)/2.0-self.Raumtemp
    
    def get_Nennleistung(self):
        self.get_uebertemp()
        self.Nennleistung = int(round(float(self.F_Nennleistung) * \
            ((self.uebertemp/float(self.F_NennTemperatur)) ** float(str(self.F_NennExponent).replace(',','.'))) ))

    def get_Temperatur(self):
        for conn in self.elem.MEPModel.ConnectorManager.Connectors:
            if conn.IsConnected:
                MEPSystem = conn.MEPSystem
                if MEPSystem:
                    temp = get_value(doc.GetElement(MEPSystem.GetTypeId()).get_Parameter(DB.BuiltInParameter.RBS_PIPE_FLUID_TEMPERATURE_PARAM))
                    systemtype = MEPSystem.SystemType.ToString()
                    if systemtype == 'SupplyHydronic':
                        self.vorlauf = temp
                    elif systemtype == 'ReturnHydronic':
                        self.ruecklauf = temp
    
    def get_value(self,param):
        return get_value(param)

    def Muster_Pruefen(self,el):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = el.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False

    def GetRaum(self):
        try:
            self.raum = self.elem.Space[self.phase].Number + ' - ' + self.elem.Space[self.phase].LookupParameter('Name').AsString()
        except:
            if not self.ismuster:
                try:
                    param_einbauort = self.get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
                    if param_einbauort not in DICT_MEP_NUMMER_NAME.keys():
                        logger.error('Einbauort konnte nicht ermittelt werden, FamilieName: {}, TypName: {}, ElementId: {}'.format(self.familyname,self.typname,self.elem.Id.ToString()))
                    else:
                        self.raum = DICT_MEP_NUMMER_NAME[param_einbauort][0]
                except:pass
            return
    
    def Wert_schreiben(self):
        try:self.HL_Param.SetValueString(str(round(self.Heizleistung,2)))
        except:pass
        try:self.SL_Param.SetValueString(str(round(self.sollleistung,2)))
        except:pass


class MEP_Raum:
    def __init__(self,elemid,HK_Bauteile):
        self.elemid = elemid
        self.HK_Bauteile = HK_Bauteile
        self.HKLeistungindouble = 0
        self.HKLeistung = '0 W'
        self.auswertung = 'OK'
        self.BLACK = BLACK
        self.RED = RED
        self.foreground = self.BLACK
        self.elem = doc.GetElement(self.elemid)
        self.raumnr = self.elem.Number + ' - ' +self.elem.LookupParameter('Name').AsString()
        self.raumnummer = self.elem.Number
        self.raumtemperatur = self.get_value(self.elem.LookupParameter('LIN_BA_DESIGN_HEATING_TEMPERATURE'))
        self.LeistungInDouble = self.get_value(self.elem.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD'))
        self.Leistung = self.elem.LookupParameter('LIN_BA_CALCULATED_HEATING_LOAD').AsValueString()
        self.HK_Update()
        self.Auswerten()
    
    def HK_Update(self):
        for ele in self.HK_Bauteile:
            try:
                ele.Raumtemp = self.raumtemperatur
                ele.get_Nennleistung()
                ele.farbe()
            except Exception as e:
                logger.error('Elementid {}'.format(ele.elem.Id.ToString()))
    
    
    def Auswerten(self):
        leistung = 0
        for ele in self.HK_Bauteile:
            leistung += ele.Heizleistung
        self.HKLeistungindouble = leistung
        self.HKLeistung = str(int(round(self.HKLeistungindouble))) + ' W'
        if abs(self.HKLeistungindouble - self.LeistungInDouble) < 5:
            self.auswertung = 'OK'
            self.foreground = self.BLACK
        else:
            self.auswertung = 'Passt nicht'
            self.foreground = self.RED
    
    def get_value(self,param):
        return get_value(param)

Dict_MEP = {}

def get_Daten_Liste():
    for HK in AUSWAHL_HEIZKOERPER_IS:
        if HK.checked:
            for elemid in HK.elems:
                hkinstance = HeizkoerperInstance(elemid.ToString(),HK)
                if hkinstance.raum not in Dict_MEP.keys():
                    Dict_MEP[hkinstance.raum] = [hkinstance]
                else:Dict_MEP[hkinstance.raum].append(hkinstance)

get_Daten_Liste()

Dict_MEP_Itemssource = {}

def get_Itemssource():
    if FamilienDialog.result == False:
        return
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).\
        WhereElementIsNotElementType().ToElements()
    for el in spaces:
        try:
            name = el.Number + ' - ' +el.LookupParameter('Name').AsString()
            if name in Dict_MEP.keys():
                mepraum = MEP_Raum(el.Id,ObservableCollection[HeizkoerperInstance](Dict_MEP[name]))
                Dict_MEP_Itemssource[name] = mepraum
        except Exception as e:
            logger.error(e)

get_Itemssource()

class ExternaleventListe(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.ExecuteApp = None
        self.name = ''
    
    def Execute(self,uiapp):
        try:self.ExecuteApp(uiapp)
        except Exception as e:print(e)
    
    def GetName(self):
        return self.name
    
    def Selecteditem(self,uiapp):
        self.name = 'Selecteditem' 
        uidoc = uiapp.ActiveUIDocument
        uidoc.Selection.SetElementIds(List[DB.ElementId]([self.GUI.lv_HK.SelectedItem.elem.Id]))
 
    def DatenAktualisieren(self,uiapp):
        self.name = 'Daten aus Revit' 
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document    
        elems = [doc.GetElement(e) for e in uidoc.Selection.GetElementIds()]
        if len(elems) != 1:
            TaskDialog.Show('Info.','Bitte nur ein MEP-Raum auswählen!')
            return
        el = elems[0]
        if el.Category.Id.ToString() != '-2003600':
            TaskDialog.Show('Info.','Bitte ein MEP-Raum auswählen!')
            return
        Nummer = el.Number + ' - ' +el.LookupParameter('Name').AsString()
        self.GUI.Nummer.SelectedItem = Nummer
        if Nummer not in self.GUI.MEP_IS.keys():
            TaskDialog.Show('Info','Kein Heizkörper in Raum {} gefunden'.format(Nummer))
            return
        self.GUI.mepraum = self.GUI.MEP_IS[Nummer]
        self.GUI.set_up()
    
    def HKLeistungschreiben(self,mep):
        for bauteil in mep.HK_Bauteile:
            bauteil.Wert_schreiben()
        mep.Auswerten()
        if mep.auswertung != 'OK':
            print('Die Heizleistung von Heizkörper in MEP Raum {} passt nicht. '.format(mep.raumnr))
        self.GUI.set_up()
    
    def HKL_Gleich_MEP(self,uiapp):
        self.name = "HK-Leistung-gleichmäßig für Raum {}".format(self.GUI.mepraum.raumnr)
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document  
        summe = sum([x.Nennleistung for x in self.GUI.mepraum.HK_Bauteile])
        if summe == 0:return
        for el in self.GUI.mepraum.HK_Bauteile:
            el.Heizleistung = round(el.Nennleistung * self.GUI.mepraum.LeistungInDouble / summe,1)
            el.sollleistung = el.Heizleistung
            if el.Heizleistung > el.Nennleistung:
                el.fore = RED
                el.Heizleistung = el.Nennleistung

            else:
                el.fore = BLACK
        t = DB.Transaction(doc,self.name)
        t.Start()
        self.HKLeistungschreiben(self.GUI.mepraum)
        t.Commit()
        t.Dispose()
    
    def HKL_Gleich_PRO(self,uiapp):
        self.name = "HK-Leistung-gleichmäßig fürs Projekt"
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document 
        result = TaskDialog.Show('HK-Leistung verteilen','HK-Leistung-gleichmäßig fürs Projekt verteilen?',TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No )
        if result != TaskDialogResult.Yes:
            return

        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.MEP_IS.values():
            summe = sum([x.Nennleistung for x in mep.HK_Bauteile])
            if summe == 0:continue
            for el in mep.HK_Bauteile:
                el.Heizleistung = round(el.Nennleistung * mep.LeistungInDouble / summe,1)
                el.sollleistung = el.Heizleistung
                if el.Heizleistung > el.Nennleistung:
                    el.fore = RED
                    el.Heizleistung = el.Nennleistung
                else:
                    el.fore = BLACK
            self.HKLeistungschreiben(mep)
            
        t.Commit()
        t.Dispose()
    
    def Aederung_MEP(self,uiapp):
        self.name = "HK-Leistung für Raum {}".format(self.GUI.mepraum.raumnr)
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document  
        t = DB.Transaction(doc,self.name)
        t.Start()
        self.HKLeistungschreiben(self.GUI.mepraum)
        t.Commit()
        t.Dispose()
    
    