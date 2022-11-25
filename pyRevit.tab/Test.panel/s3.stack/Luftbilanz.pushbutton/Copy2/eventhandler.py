# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,UIView,RevitCommandId,UIApplication,PostableCommand,TaskDialogResult,TaskDialogCommonButtons,ExternalEvent
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility,GridLength
import System
from rpw import revit,DB,UI
from pyrevit import script,forms
from System.Windows.Input import ModifierKeys,Keyboard,Key
from System.Collections.Generic import List
from System.Windows.Media import Brushes
from clr import GetClrType
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs


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
config = script.get_config(name+number+'Raumluftbilanz')

RED = Brushes.Red
BLACK = Brushes.Black
BLUE = Brushes.Blue
BLUEVIOLET = Brushes.BlueViolet
HIDDEN = Visibility.Hidden
VISIBLE = Visibility.Visible

DICT_MEP_NUMMER_NAME = {} # 100.103 : 100.103 - Büro
DICT_MEP_AUSLASS = {} ## MEPID: OB(Auslässe)
DICT_MEP_VSR = {} ## MEPID: [VSRID]
DICT_MEP_UEBERSTROM = {} ## MEPID: {'Ein': ..., 'Aus':...}

ELEMID_UEBER = [] ## ElememntId Überstrom
LISTE_SCHACHT = []

DICT_VSR_MEP_NUR_NUMMER = {} ## VSRID: [MEPNr,MEPNr]
DICT_VSR_MEP = {} ## VSRID: [MEPNr-Name,MEPNr-Name]
DICT_VSR_VSRLISTE = {} ## VSRID: [VSRID,VSRID]
DICT_VSR_AUSLASS = {} ## VSRID: OB(Auslässe)
DICT_VSR = {} ## VSRID: Class VSR
LISTE_VSR = []
LISTE_IRIS = []

views = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.View3D)).ToElements()
for el in views:
    if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
        uidoc.ActiveView = el
        break

active_view = uidoc.ActiveView
if active_view.Name != 'Berechnungsmodell_L_KG4xx_IGF':
    logger.error('die Ansicht "Berechnungsmodell_L_KG4xx_IGF" nicht gefunden!')
    script.exit()

class ABFRAGE(forms.WPFWindow):
    def __init__(self,maininfo = '',checked=True,height = None,minmax = True):
        forms.WPFWindow.__init__(self,'abfrage.xaml')
        self.maininfo.Text = maininfo
        self.gridlenge = GridLength(0.0)
        self.hoehe = height
        self.minmax = minmax
        if self.minmax == False:
            self.maingrid.RowDefinitions[1].Height = self.gridlenge
        if self.hoehe != None:
            self.Height = self.hoehe
        self.checked = checked
        self.result = False
        if self.checked == True:
            self.bestaetigen.IsChecked = True
            self.maingrid.RowDefinitions[2].Height = self.gridlenge
        else:
            self.bestaetigen.IsChecked = False
            self.ja.IsEnabled = False
    
    def movewindow(self, sender, args):
        self.DragMove()
        
    def checkedchanged(self,sender,e):
        if sender.IsChecked == True:
            self.ja.IsEnabled = True
        else:
            self.ja.IsEnabled = False
    
    def yes(self,sender,e):
        self.result = True
        self.Close()
    
    def no(self,sender,e):
        self.result = False
        self.Close()

    @staticmethod
    def show(maininfo = '',checked=True):
        abfrage = ABFRAGE(maininfo,checked)
        abfrage.ShowDialog()
        return abfrage.result

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

## Familien Einstellungen
class Itemtemplate(TemplateItemBase):
    def __init__(self,name,checked = False):
        TemplateItemBase.__init__(self)
        self.name = name
        self._checked = checked
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

IS_AUSLASS = ObservableCollection[Itemtemplate]()
IS_AUSLASS_D = ObservableCollection[Itemtemplate]()
IS_AUSLASS_LAB = ObservableCollection[Itemtemplate]()
IS_AUSLASS_24H = ObservableCollection[Itemtemplate]()
IS_VSR = ObservableCollection[Itemtemplate]()
IS_DOSSEL = ObservableCollection[Itemtemplate]()
def get_IS():
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Liste_Auslass = []
    Liste_Auslass_Lab = []
    Liste_Zubehoer = []
    for el in Families:
        FamName = el.Name
        if el.FamilyCategoryId.IntegerValue == -2008013:
            if FamName not in Liste_Auslass:
                Liste_Auslass.append(FamName)
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Liste_Auslass_Lab:
                    Liste_Auslass_Lab.append(typname)
        elif el.FamilyCategoryId.IntegerValue == -2008016:
            if FamName not in Liste_Auslass:
                Liste_Zubehoer.append(FamName)

    #     FamName = el.Symbol.FamilyName
    #     famundtyp = FamName + ': ' + el.Name
    #     if FamName not in Liste_Auslass:
    #         Liste_Auslass.append(FamName)
    #     if famundtyp not in Liste_Auslass_Lab:
    #         Liste_Auslass_Lab.append(famundtyp)
        
    # for el in Ductaccessorys:
    #     if el.Symbol.FamilyName not in Liste_Zubehoer:
    #         Liste_Zubehoer.append(el.Symbol.FamilyName)

    for el in sorted(Liste_Auslass):
        IS_AUSLASS_D.Add(Itemtemplate(el))
    for el in sorted(Liste_Zubehoer):
        IS_VSR.Add(Itemtemplate(el))
        IS_DOSSEL.Add(Itemtemplate(el))
    for el in sorted(Liste_Auslass_Lab):
        IS_AUSLASS.Add(Itemtemplate(el,True))
        IS_AUSLASS_LAB.Add(Itemtemplate(el))
        IS_AUSLASS_24H.Add(Itemtemplate(el))

get_IS()

class Familien(forms.WPFWindow):
    def __init__(self,IS_Auslass0,IS_Auslass1,IS_Auslass2,IS_Auslass3,IS_VSR,IS_Drossel):
        forms.WPFWindow.__init__(self,'Familien.xaml')
        self.IS_Auslass0 = IS_Auslass0
        self.IS_Auslass1 = IS_Auslass1
        self.IS_Auslass2 = IS_Auslass2
        self.IS_Auslass3 = IS_Auslass3
        self.IS_VSR = IS_VSR
        self.IS_Drossel = IS_Drossel
        self.read_config()
        self.auslass.ItemsSource = IS_Auslass0
        self.auslass_d.ItemsSource = IS_Auslass1
        self.auslass_lab.ItemsSource = IS_Auslass2
        self.auslass_24h.ItemsSource = IS_Auslass3
        self.vsr.ItemsSource = IS_VSR
        self.klappe.ItemsSource = IS_Drossel
        self.temp0 = ObservableCollection[Itemtemplate]()
        self.temp1 = ObservableCollection[Itemtemplate]()
        self.temp2 = ObservableCollection[Itemtemplate]()
        self.temp3 = ObservableCollection[Itemtemplate]()
        self.temp4 = ObservableCollection[Itemtemplate]()
        self.temp5 = ObservableCollection[Itemtemplate]()
        
    def read_config(self):
        def readconfig_intern(liste,liste1,checked):
            for el in liste:
                if el.name in liste1:
                    el.checked = checked
        try:readconfig_intern(self.IS_Auslass0,config.auslass,False)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Auslass1,config.auslassd,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_VSR,config.vsr,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Drossel,config.drossel,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Auslass2,config.auslasslab,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Auslass3,config.auslass24h,True)
        except Exception as e:print(e)
    
    def write_config(self):
        def write_config_intern(liste0,checked):
            return [el.name for el in liste0 if el.checked == checked]
        config.auslass = write_config_intern(self.IS_Auslass0,False)
        config.auslassd = write_config_intern(self.IS_Auslass1,True)
        config.vsr = write_config_intern(self.IS_VSR,True)
        config.drossel = write_config_intern(self.IS_Drossel,True)
        config.auslasslab = write_config_intern(self.IS_Auslass2,True)
        config.auslass24h = write_config_intern(self.IS_Auslass3,True)
        script.save_config()
    
    def checkedchanged(self,sender,liste):
        Checked = sender.IsChecked
        if liste.SelectedItem is not None:
            try:
                if sender.DataContext in liste.SelectedItems:
                    for item in liste.SelectedItems:
                        try:item.checked = Checked
                        except:pass
                else:pass
            except:pass

    def checkedchanged_auslass(self,sender,e):
        self.checkedchanged(sender,self.auslass)

    def checkedchanged_auslass_lab(self,sender,e):
        self.checkedchanged(sender,self.auslass_lab)
    
    def checkedchanged_auslass_24h(self,sender,e):
        self.checkedchanged(sender,self.auslass_24h)
    
    def checkedchanged_auslass_d(self,sender,e):
        self.checkedchanged(sender,self.auslass_d)
    
    def checkedchanged_vsr(self,sender,e):
        self.checkedchanged(sender,self.vsr)
    
    def checkedchanged_klappe(self,sender,e):
        self.checkedchanged(sender,self.klappe)
    
    def suche_text_changed(self,sender,liste0,liste1,liste2):
        liste1.Clear()
        text = sender.Text
        if not text:liste2.ItemsSource = liste0
        else:
            for el in liste0:
                if el.name.upper().find(text.upper()) != -1:
                    liste1.Add(el)
            liste2.ItemsSource = liste1
        liste2.Items.Refresh()
    
    def suche_textchanged_Lab(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass2,self.temp0,self.auslass_lab)

    def suche_textchanged_24h(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass3,self.temp1,self.auslass_24h)
    
    def suche_textchanged_auslass(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass0,self.temp2,self.auslass)
    
    def suche_textchanged_auslass_d(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass1,self.temp3,self.auslass_d)
    
    def suche_textchanged_vsr(self,sender,e):
        self.suche_text_changed(sender,self.IS_VSR,self.temp4,self.vsr)
    
    def suche_textchanged_klappe(self,sender,e):
        self.suche_text_changed(sender,self.IS_Drossel,self.temp5,self.klappe)
    
    def ok(self,sender,e):
        self.write_config()
        self.Close()
    
    def schliessen(self,sender,e):
        self.Close()
        script.exit()
        
familien = Familien(IS_AUSLASS,IS_AUSLASS_D,IS_AUSLASS_LAB,IS_AUSLASS_24H,IS_VSR,IS_DOSSEL)
try:familien.ShowDialog()
except Exception as e:
    logger.error(e)
    familien.Close()
    script.exit()

VSR_AUSLASS_LISTE = config.auslass
DRO_AUSLASS_LISTE = config.auslassd
LAB_AUSLASS_LISTE = config.auslasslab
H24_AUSLASS_LISTE = config.auslass24h
VSR_FAMILIE_LISTE = config.vsr
DRO_FAMILIE_LISTE = config.drossel

## DICT_MEP_NUMMER_NAME = {} # 100.103 : 100.103 - Büro
def get_MEP_NUMMER_NAMR():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        DICT_MEP_NUMMER_NAME[el.Number] = [el.Number + ' - ' + el.LookupParameter('Name').AsString(),el.Id.ToString()]
        schacht = el.LookupParameter('TGA_RLT_InstallationsSchacht').AsInteger()
        if schacht == 1:
            name = el.LookupParameter('TGA_RLT_InstallationsSchachtName').AsString()
            if name not in LISTE_SCHACHT:
                LISTE_SCHACHT.append(name)

get_MEP_NUMMER_NAMR() 
LISTE_SCHACHT.sort()
## NoMEPSpace
class NoMEPSpace(Exception):
    def __init__(self,elemid,typ,family,art):
        self.elemid = elemid
        self.typ = typ
        self.family = family
        self.art = art
        self.errorinfo = '{}: Einbauort konnte nicht ermittelt werden, FamilieName: {}, TypName: {}, ElementId: {}'.format(self.art,self.family,self.typ,self.elemid.ToString())
    
    def __str__(self):
        return self.errorinfo

## Grundklasse für alle Familienexamplar
class FamilieExemplar(TemplateItemBase):
    def __init__(self,elemid,art):
        TemplateItemBase.__init__(self)
        self.elemid = elemid
        self.logger = logger
        self.art_temp = art
        self._art = ''
        self._Luftmengenmin = 0
        self._Luftmengennacht = 0
        self._Luftmengenmax = 0
        self._Luftmengentnacht = 0
        self._size = ''
        self._Anmerkung = ''
        self.elem = doc.GetElement(self.elemid)
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self._familyandtyp = self.familyname + ': ' + self.typname
        self.ismuster =  self.Muster_Pruefen()
        self.phase = doc.GetElement(self.elem.CreatedPhaseId)
        self.raum = ''
        self.raumnummer = ''
        self.raumid = ''
        self.GetRaum()
    
    @property
    def art(self):
        return self._art
    @art.setter
    def art(self,value):
        if value != self._art:
            self._art = value
            self.RaisePropertyChanged('art')
    
    @property
    def familyandtyp(self):
        return self._familyandtyp
    @familyandtyp.setter
    def familyandtyp(self,value):
        if value != self._familyandtyp:
            self._familyandtyp = value
            self.RaisePropertyChanged('familyandtyp')
    
    @property
    def size(self):
        return self._size
    @size.setter
    def size(self,value):
        if value != self._size:
            self._size = value
            self.RaisePropertyChanged('size')
    
    @property
    def Luftmengenmin(self):
        return self._Luftmengenmin
    @Luftmengenmin.setter
    def Luftmengenmin(self,value):
        if value != self._Luftmengenmin:
            self._Luftmengenmin = value
            self.RaisePropertyChanged('Luftmengenmin')
    
    @property
    def Luftmengenmax(self):
        return self._Luftmengenmax
    @Luftmengenmax.setter
    def Luftmengenmax(self,value):
        if value != self._Luftmengenmax:
            self._Luftmengenmax = value
            self.RaisePropertyChanged('Luftmengenmax')
    
    @property
    def Luftmengennacht(self):
        return self._Luftmengennacht
    @Luftmengennacht.setter
    def Luftmengennacht(self,value):
        if value != self._Luftmengennacht:
            self._Luftmengennacht = value
            self.RaisePropertyChanged('Luftmengennacht')
    
    @property
    def Luftmengentnacht(self):
        return self._Luftmengentnacht
    @Luftmengentnacht.setter
    def Luftmengentnacht(self,value):
        if value != self._Luftmengentnacht:
            self._Luftmengentnacht = value
            self.RaisePropertyChanged('Luftmengentnacht')
    
    @property
    def Anmerkung(self):
        return self._Anmerkung
    @Anmerkung.setter
    def Anmerkung(self,value):
        if value != self._Anmerkung:
            self._Anmerkung = value
            self.RaisePropertyChanged('Anmerkung')
    
    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
    def GetRaum(self):
        try:
            self.raumid = self.elem.Space[self.phase].Id.ToString()
            self.raumnummer = self.elem.Space[self.phase].Number
            self.raum = self.elem.Space[self.phase].Number + ' - ' + self.elem.Space[self.phase].LookupParameter('Name').AsString()
        except:
            if not self.ismuster:
                try:
                    param_einbauort = get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
                    if param_einbauort not in DICT_MEP_NUMMER_NAME.keys():
                        raise NoMEPSpace(self.elemid,self.typname,self.familyname,self.art_temp)
                    else:
                        self.raum = DICT_MEP_NUMMER_NAME[param_einbauort][0]
                        self.raumid = DICT_MEP_NUMMER_NAME[param_einbauort][1]
                except NoMEPSpace as e:
                    self.logger.error(str(e))
            return

    def set_up(self):
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.familyandtyp = self.familyname + ': ' + self.typname
    
    def get_Size(self):
        try:
            conns = self.elem.MEPModel.ConnectorManager.Connectors

            for conn in conns:
                try:
                    self.lufttyp = conn.DuctSystemType.ToString()
                except:pass
                try:
                    return 'DN ' + str(int(round(conn.Radius*609.6)))
                except:
                    Height = str(int(round(conn.Height*304.8)))
                    Width = str(int(round(conn.Width*304.8)))
                    return Width + 'x' + Height
        except:return
    
    def wert_schreiben(self):
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(self.Luftmengenmin).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(self.Luftmengennacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(self.Luftmengenmax).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(self.Luftmengentnacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raumnummer)
        except Exception as e:pass

## Klasse für Überstromsgruppe
class Baugruppe:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        if self.Pruefen():
            self.Volumen = get_value(self.elem.LookupParameter('IGF_RLT_Überströmung'))
            self.Eingang = ''
            self.Ausgang = ''
            try:self.Analyse()
            except Exception as e:logger.error(e)

    @property
    def auslaesse(self):
        auslass_liste = []
        for elemid in self.elem.GetMemberIds():
            elem = doc.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008013' and elem.Symbol.FamilyName.upper().find("ÜBER") != -1:
                auslass_liste.append(elem.Id)
        return auslass_liste
    
    def Pruefen(self):
        return len(self.auslaesse) == 2
   
    def Analyse(self):
        for elemid in self.auslaesse:
            auslass = UeberStromAuslass(elemid,self.Volumen)
            if auslass.typ == 'Aus':
                self.Ausgang = auslass
            elif auslass.typ == 'Ein':
                self.Eingang = auslass
        self.Ausgang.anderesraum = self.Eingang.raum
        self.Eingang.anderesraum = self.Ausgang.raum

## Klasse für Überstrom
class UeberStromAuslass(FamilieExemplar):
    def __init__(self,elemid,vol):
        FamilieExemplar.__init__(self,elemid,'Überstrom')
        self._menge = vol
        self._anderesraum = ""
        self.typ = self.get_typ()
    
    @property
    def menge(self):
        return self._menge
    @menge.setter
    def menge(self,value):
        if value != self._menge:
            self._menge = value
            self.RaisePropertyChanged('menge')
    @property
    def anderesraum(self):
        return self._anderesraum
    @anderesraum.setter
    def anderesraum(self,value):
        if value != self._anderesraum:
            self._anderesraum = value
            self.RaisePropertyChanged('anderesraum')
        
    @property
    def conns(self):
        return list(self.elem.MEPModel.ConnectorManager.Connectors)

    def get_typ(self):
        conn = self.conns[0]
        if conn.Direction.ToString() == 'Out':
            return 'Aus'
        elif conn.Direction.ToString() == 'In':
            return 'Ein'

# ## Klasse für Auslässe
class Luftauslass(FamilieExemplar):
    """
    iris: '', -1, str(elemid)
    vsr = elemid
    vsrliste: []
    vsr kann nicht in vsrliste sein.
    wenn vsr in vsrliste ist, vsr ist falsch platziert.
    vsr:[]
    oder vsr: [vsr1,vsr2]
    """
    def __init__(self,elemid):
        FamilieExemplar.__init__(self,elemid,'Luftauslass')
        self.Liste_LAB_Auslass = LAB_AUSLASS_LISTE
        self.Liste_24h_Auslass = H24_AUSLASS_LISTE
        self.enabledmin = True
        self.enabledmax = True
        self.enablednacht = True
        self.enabledtnacht = True
        self.System = ''
        self.AnlNr = ''
        self.vsr = ''
        self.RoutingListe = []
        self.Einbauteile = []
        self.VSR_Liste = []
        self.lufttyp = ''
        self.iris = ''
        self.VSR_Class = None
        self.Haupt_Class = None
        self.IRIS_Class = None
        self.Luftmengenermitteln()
        self.size = self.get_Size()
        try:self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:self.System = ''
        try:
            self.AnlNr = doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId()).LookupParameter('IGF_X_AnlagenNr').AsValueString()
        except:self.AnlNr = ''
      
        if self.familyandtyp not in VSR_AUSLASS_LISTE:
            self.get_RountingListe(self.elem)   
        self.get_Art()
        if self.familyname in DRO_AUSLASS_LISTE:
            if self.iris in [-1,'']:
                self.Anmerkung = 'Der Auslass ist von Drosselklappe geregelt aber keine ist damit angeschlossen.'

        else:
            if self.iris not in [-1,'']:
                self.Anmerkung = 'Der Auslass ist nicht von Drosselklappe geregelt aber eine ist damit angeschlossen.'

        if self.vsr in self.VSR_Liste:
            self.Anmerkung = 'VSR falsch angeschlossen. Grund kann sein, dass das Eingang des VSRs ist mit ein Zuluftgitter verbunden.'
        
        if not self.vsr:
            self.Anmerkung = 'Kein VSR.'
    
    def set_up(self):
        FamilieExemplar.set_up(self)
        self.Luftmengenermitteln()
        self.size = self.get_Size()
        self.get_Art()

    def Luftmengenermitteln(self):
        try:self.Luftmengenmin = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin'))),1)
        except:pass
        try:self.Luftmengennacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht'))),1)
        except:pass
        try:self.Luftmengenmax = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax'))),1)
        except:pass
        try:self.Luftmengentnacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht'))),1)
        except:pass

    def get_RountingListe(self,element):
        if self.RoutingListe.Count > 500:return
        elemid = element.Id.ToString()
        self.RoutingListe.append(elemid)
        cate = element.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Luftkanaldämmung außen', 'Luftkanaldämmung innen', 'Rohre', 'Rohrformteile', 'Rohrzubehör','Rohrdämmung']:
            conns = None
            try:conns = element.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = element.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and self.iris == '':self.iris = -1
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.RoutingListe:
                            if owner.Category.Name == 'Luftkanalzubehör':
                                faminame = owner.Symbol.FamilyName
                                typname = owner.Name
                                if faminame in VSR_FAMILIE_LISTE:
                                    conns_temp = owner.MEPModel.ConnectorManager.Connectors
                                    for conn_temp in conns_temp:
                                        if self.lufttyp == 'ReturnAir':
                                            if conn_temp.Direction.ToString() == 'In' or conn_temp.Description == 'Haupt':#???
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                        if self.lufttyp == 'SupplyAir':
                                            if conn_temp.Direction.ToString() == 'Out' or conn_temp.Description == 'Haupt':#???
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                    if not self.vsr:
                                        self.vsr = owner.Id.ToString()
                                    self.VSR_Liste.append(owner.Id.ToString())
                                    return
                                elif faminame in DRO_FAMILIE_LISTE:
                                    if not self.iris:self.iris = owner.Id.ToString()
                                    
                            self.get_RountingListe(owner)
    
    def get_Art(self):
        systyp = self.System
        if systyp:
            if systyp.upper().find('TIERHALTUNG') != -1:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB' 
            else:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB'
        else:
            self.art = 'XXX'
        try:
            if self.familyandtyp in self.Liste_LAB_Auslass:
                self.art = 'LAB'
            elif self.familyandtyp in self.Liste_24h_Auslass:
                self.art = '24h'
                self.enabledmax = False
                self.enablednacht = False
                self.enabledtnacht = False
        except:
            pass

## Klasse für Volumenstromregler
class VSR(FamilieExemplar):
    def __init__(self,elemid):
        FamilieExemplar.__init__(self,elemid,'VSR')
        self.checked = False
        self.Auslass = ObservableCollection[Luftauslass]()
        self.liste_Raum = []
        self.liste_Raum_nurNummer = []
        self.size = self.get_Size()
        self.IsHaupt = False
        self.IsIris = False
        self.List_Iris = ObservableCollection[VSR]()
        self.List_VSR = ObservableCollection[VSR]()
        self.List_Haupt = ObservableCollection[VSR]()
    
    def set_up(self):
        FamilieExemplar.set_up(self)
        self.size = self.get_Size()
        self.get_Art()

    def Luftmengenermitteln(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(auslass.Luftmengenmin)
            Luftmengennacht += float(auslass.Luftmengennacht)
            Luftmengenmax += float(auslass.Luftmengenmax)
            Luftmengentnacht += float(auslass.Luftmengentnacht)
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht))
    
    def Luftmengenermitteln_new(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
            Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
            Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
            Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def Luftmengenermitteln_um(self):
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0

        for auslass in self.Auslass:
            if auslass.art == 'UMZU':
                Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
                Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
                Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
                Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def get_Art(self):
        art_liste = []
        for auslass in self.Auslass:
            if auslass.art:
                if auslass.art not in art_liste:
                    art_liste.append(auslass.art)
        
        if art_liste.Count == 0:return
        if art_liste.Count == 1:
            self.art = art_liste[0]
        elif art_liste.Count == 2:
            if 'RZU' in art_liste and 'RAB' in art_liste:
                self.art = 'RUM'
                for auslass in self.Auslass:
                    if auslass.art == 'RZU':
                        auslass.art = 'UMZU'
                    elif auslass.art == 'RAB':
                        auslass.art = 'UMAB'
            elif 'RAB' in art_liste and 'LAB' in art_liste:
                self.art = 'RAB'
            else:
                print(self.elemid.ToString(),art_liste)
        elif art_liste.Count == 3 and '24h' in art_liste and 'LAB' in art_liste and 'RAB' in art_liste:self.art = 'RAB'
        else:print(self.elemid.ToString(),art_liste)
    
    def wert_schreiben(self):
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(self.Luftmengenmin).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(self.Luftmengennacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(self.Luftmengenmax).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(self.Luftmengentnacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raumnummer)
        except Exception as e:pass
        try:
            numm = ''
            for nummer in sorted(self.liste_Raum_nurNummer):
                numm += nummer + ', '
            self.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
        except:pass

def Get_Ueberstrom_Info():
    """
    DICT_MEP_UEBERSTROM:  {[raumid:{'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}}
    ELEMID_UEBER:   [lemid.ToString()]
    """
    Baugruppen = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title = "{value}/{max_value} Überströmungsbaugruppen",cancellable=True,step=10) as pb:
        for n, BGid in enumerate(Baugruppen):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(Baugruppen))
            baugruppe = Baugruppe(BGid)
            if not baugruppe.Pruefen():
                continue
            if not baugruppe.Eingang.raumid in DICT_MEP_UEBERSTROM.keys():
                DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
            if not baugruppe.Ausgang.raumid in DICT_MEP_UEBERSTROM.keys():
                DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
            DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid][baugruppe.Eingang.typ].Add(baugruppe.Eingang)
            DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid][baugruppe.Ausgang.typ].Add(baugruppe.Ausgang)
            ELEMID_UEBER.append(baugruppe.Eingang.elemid.ToString())
            ELEMID_UEBER.append(baugruppe.Ausgang.elemid.ToString())

Get_Ueberstrom_Info()

def Get_Auslass_Info():
    """
    DICT_MEP_AUSLASS: {raumid:{auslass.art:{auslass.familyandtyp:[auslass(Klasse)]}}}
    DICT_VSR_MEP: {auslass.vsr: ["auslass.raumnr-auslass.raumname"],auslass.iris: ["auslass.raumnr-auslass.raumname"]}
    DICT_VSR_MEP_NUR_NUMMER: {auslass.vsr: [auslass.raumnummer],auslass.iris: [auslass.raumnummer]}
    DICT_VSR_VSRLISTE: {auslass.vsr: [auslass.VSR_Liste]} elemid: [elemid]
    DICT_MEP_VSR: {auslass.raumid: [auslass.iris,auslass.vsr]}
    DICT_VSR_AUSLASS: {auslass.vsr:{auslass.art:{auslass.familyandtyp:[auslass]}}}
    DICT_VSR: {vsrid:vsr}
    LISTE_VSR: [elemid]
    LISTE_IRIS: [elemid]

    """
    Ductterminalids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe", cancellable=True,step=10) as pb:
        for n, ductid in enumerate(Ductterminalids):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(Ductterminalids))
            if ductid.ToString() in ELEMID_UEBER:continue
            auslass = Luftauslass(ductid)
            
            if not auslass.raumid in DICT_MEP_AUSLASS.keys():
                 DICT_MEP_AUSLASS[auslass.raumid] = {}
            if not auslass.art in DICT_MEP_AUSLASS[auslass.raumid].keys():
                DICT_MEP_AUSLASS[auslass.raumid][auslass.art] = {}
            if not auslass.familyandtyp in DICT_MEP_AUSLASS[auslass.raumid][auslass.art].keys():
                DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp] = []
            DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp].append(auslass)

            if auslass.familyandtyp in VSR_AUSLASS_LISTE:
                continue
            if not auslass.vsr:
                if auslass.Muster_Pruefen() != True:
                    logger.error('Kein VSR mit Auslass {} angeschlossen. Raum {}, Familie: {}'.format(auslass.elemid,auslass.raumnummer,auslass.familyandtyp))
                continue
            
            if not auslass.vsr in DICT_VSR_MEP.keys():
                DICT_VSR_MEP[auslass.vsr] = [auslass.raum]
                DICT_VSR_MEP_NUR_NUMMER[auslass.vsr] = [auslass.raumnummer]
            else:
                if auslass.raum not in DICT_VSR_MEP[auslass.vsr]:
                    DICT_VSR_MEP[auslass.vsr].append(auslass.raum)
                    DICT_VSR_MEP_NUR_NUMMER[auslass.vsr].append(auslass.raumnummer)
            if auslass.iris not in [-1,'']:

                if not auslass.iris in DICT_VSR_MEP.keys():
                    DICT_VSR_MEP[auslass.iris] = [auslass.raum]
                    DICT_VSR_MEP_NUR_NUMMER[auslass.iris] = [auslass.raumnummer]
                else:
                    if auslass.raum not in DICT_VSR_MEP[auslass.iris]:
                        DICT_VSR_MEP[auslass.iris].append(auslass.raum)
                        DICT_VSR_MEP_NUR_NUMMER[auslass.iris].append(auslass.raumnummer)
            
            if len(auslass.VSR_Liste) != 0 and auslass.vsr not in auslass.VSR_Liste:
                if auslass.vsr not in DICT_VSR_VSRLISTE.keys():
                    DICT_VSR_VSRLISTE[auslass.vsr] = auslass.VSR_Liste
                else:
                    DICT_VSR_VSRLISTE[auslass.vsr].extend(auslass.VSR_Liste)
            else:
                LISTE_VSR.append(auslass.vsr)
            
            if auslass.iris not in [-1,'']:
                LISTE_IRIS.append(auslass.iris)

            if not auslass.raumid in DICT_MEP_VSR.keys():
                DICT_MEP_VSR[auslass.raumid] = []
                DICT_MEP_VSR[auslass.raumid].append(auslass.vsr) 
                if auslass.iris not in [-1,'']:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.iris)
            else:
                if auslass.vsr not in DICT_MEP_VSR[auslass.raumid]:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.vsr)
                if auslass.iris not in [-1,'']:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.iris)

            if not auslass.vsr in DICT_VSR_AUSLASS.keys():
                DICT_VSR_AUSLASS[auslass.vsr] = {}
            if not auslass.art in DICT_VSR_AUSLASS[auslass.vsr].keys():
                DICT_VSR_AUSLASS[auslass.vsr][auslass.art] = {}
            if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.vsr][auslass.art].keys():
                DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp] = []
            DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp].append(auslass)
            if auslass.iris not in [-1,'']:
                if not auslass.iris in DICT_VSR_AUSLASS.keys():
                    DICT_VSR_AUSLASS[auslass.iris] = {}
                if not auslass.art in DICT_VSR_AUSLASS[auslass.iris].keys():
                    DICT_VSR_AUSLASS[auslass.iris][auslass.art] = {}
                if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.iris][auslass.art].keys():
                    DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp] = []
                DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp].append(auslass)  


    liste_temp = DICT_VSR_AUSLASS.keys()[:]
    liste_temp1 = LISTE_VSR[:]
    liste_temp2 = LISTE_IRIS[:]
    liste_temp.extend(liste_temp1)
    liste_temp.extend(liste_temp2)
    liste_temp = set(liste_temp)
    liste_temp = list(liste_temp)
   

    with forms.ProgressBar(title = "{value}/{max_value} Volumenstromregler", cancellable=True, step=10) as pb:
        for n, vsrid in enumerate(liste_temp):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(liste_temp))
            if vsrid in DICT_VSR_VSRLISTE.keys():
                vsrliste = set(DICT_VSR_VSRLISTE[vsrid])
                vsrliste = list(vsrliste)                
                for vsrid_neu in vsrliste:
                    if vsrid != vsrid_neu:
                        if vsrid_neu not in DICT_VSR_AUSLASS.keys():
                            logger.error('Kein Auslass mit VSR {} verbunden.'.format(vsrid_neu))
                            continue
                        for key0 in DICT_VSR_AUSLASS[vsrid_neu].keys():
                            for key1 in DICT_VSR_AUSLASS[vsrid_neu][key0].keys():
                                for auslass in DICT_VSR_AUSLASS[vsrid_neu][key0][key1]:
                                    if key0 not in DICT_VSR_AUSLASS[vsrid].keys():
                                        DICT_VSR_AUSLASS[vsrid][key0] = {}
                                    if key1 not in DICT_VSR_AUSLASS[vsrid][key0].keys():
                                        DICT_VSR_AUSLASS[vsrid][key0][key1] = []
                                    if auslass not in DICT_VSR_AUSLASS[vsrid][key0][key1]:
                                        DICT_VSR_AUSLASS[vsrid][key0][key1].append(auslass)

         
            vsr = VSR(DB.ElementId(int(vsrid)))
            if vsr.elemid.ToString() in LISTE_IRIS:
                vsr.IsIris = True
            elif vsrid in DICT_VSR_VSRLISTE.keys():
                vsr.IsHaupt = True

            vsr.Auslass = ObservableCollection[Luftauslass]()
            for art in sorted(DICT_VSR_AUSLASS[vsrid].keys()):
                for fam in sorted(DICT_VSR_AUSLASS[vsrid][art].keys()):
                    for terminal in DICT_VSR_AUSLASS[vsrid][art][fam]:
                        vsr.Auslass.Add(terminal)
            vsr.get_Art()
            vsr.Luftmengenermitteln()
            vsr.liste_Raum = DICT_VSR_MEP[vsrid]
            vsr.liste_Raum_nurNummer = DICT_VSR_MEP_NUR_NUMMER[vsrid]
            
            DICT_VSR[vsrid] = vsr
            for auslass in vsr.Auslass:
                if vsr.IsIris:
                    auslass.IRIS_Class = vsr
                elif vsr.IsHaupt:
                    auslass.Haupt_Class = vsr
                else:
                    auslass.VSR_Class = vsr

Get_Auslass_Info()

class SchachtGrundinfo(object):
    def __init__(self,name):
        self.name = name

class MEPGrundInfo(TemplateItemBase):
    def __init__(self,name,soll,tooltip):
        TemplateItemBase.__init__(self)
        self.name = name
        self._soll = soll
        self._ist = ''
        self.tooltip = tooltip
    @property
    def soll(self):
        return self._soll
    @soll.setter
    def soll(self,value):
        if value != self._soll:
            self._soll = value
            self.RaisePropertyChanged('soll')
    @property
    def ist(self):
        return self._ist
    @ist.setter
    def ist(self,value):
        if value != self._ist:
            self._ist = value
            self.RaisePropertyChanged('ist')

class MEPSchachtInfo(TemplateItemBase):
    def __init__(self,name,nr,menge,SchachtListe = []):
        TemplateItemBase.__init__(self)
        self.name = name
        self.nr = nr
        self._menge = menge
        self.SchachtListe = SchachtListe
        self._schachtindex = -1
        self.get_index()
    
    @property
    def menge(self):
        return self._menge
    @menge.setter
    def menge(self,value):
        if value != self._menge:
            self._menge = value
            self.RaisePropertyChanged('menge')
    
    @property
    def schachtindex(self):
        return self._schachtindex
    @schachtindex.setter
    def schachtindex(self,value):
        if value != self._schachtindex:
            self._schachtindex = value
            self.RaisePropertyChanged('schachtindex')

    def get_index(self):
        try:
            self.schachtindex = self.SchachtListe.index(self.nr)
        except:
            self.schachtindex = -1

class MEPAnlagenInfo(TemplateItemBase):
    def __init__(self,name,mep_nr,mep_mengen,sys_nr,sys_mengen):
        TemplateItemBase.__init__(self)
        self.name = name
        self._mep_nr = mep_nr
        self._mep_mengen = mep_mengen
        self._sys_nr = sys_nr
        self._sys_mengen = sys_mengen 
    @property
    def mep_nr(self):
        return self._mep_nr
    @mep_nr.setter
    def mep_nr(self,value):
        if value != self._mep_nr:
            self._mep_nr = value
            self.RaisePropertyChanged('mep_nr')
    @property
    def mep_mengen(self):
        return self._mep_mengen
    @mep_mengen.setter
    def mep_mengen(self,value):
        if value != self._mep_mengen:
            self._mep_mengen = value
            self.RaisePropertyChanged('mep_mengen')
    @property
    def sys_nr(self):
        return self._sys_nr
    @sys_nr.setter
    def sys_nr(self,value):
        if value != self._sys_nr:
            self._sys_nr = value
            self.RaisePropertyChanged('sys_nr')
    @property
    def sys_mengen(self):
        return self._sys_mengen
    @sys_mengen.setter
    def sys_mengen(self,value):
        if value != self._sys_mengen:
            self._sys_mengen = value
            self.RaisePropertyChanged('sys_mengen')

spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
Dict_Ueber_Manuell = {}
for ele in spaces_collector:
    raum = get_value(ele.LookupParameter("TGA_RLT_RaumÜberströmungAus"))
    if raum:
        summe2 = get_value(ele.LookupParameter('TGA_RLT_RaumÜberströmungMenge'))
        if raum not in Dict_Ueber_Manuell.keys():
            Dict_Ueber_Manuell[raum] = summe2
        else:Dict_Ueber_Manuell[raum] += summe2

class MEPRaum(object):
    def __init__(self, elemid,list_vsr):
        self.logger = logger
        self.berechnung_nach = {
                        '1': "Fläche",
                        '2': "Luftwechsel",
                        '3': "Person",
                        '4': "manuell",
                        '5': "nurZUMa",
                        '6': "nurABMa",
                        '5.1': "nurZU_Fläche",
                        '5.2': "nurZU_Luftwechsel",
                        '5.3': "nurZU_Person",
                        '6.1': "nurAB_Fläche",
                        '6.2': "nurAB_Luftwechsel",
                        '6.3': "nurAB_Person",
                        '9': "keine",
                    }
        self.einheit_liste = {
                    '1': 'm³/h pro m²',
                    '2': '-1/h',
                    '3': 'm3/h pro P',
                    '4': 'm³/h ',
                    '5': 'm³/h ',
                    '6': 'm³/h' ,
                    '5.1': "m³/h pro m²",
                    '5.2': '-1/h',
                    '5.3': 'm3/h pro P',
                    '6.1': "m³/h pro m²",
                    '6.2': '-1/h',
                    '6.3': 'm3/h pro P',
                    '9': '',
                }
        self.liste_schacht = LISTE_SCHACHT
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.box = self.elem.get_BoundingBox(None)
        self.Raumnr = self.elem.Number + ' - ' + self.elem.LookupParameter('Name').AsString()
        self.list_vsr = list_vsr
        self.list_vsr0 = ObservableCollection[VSR]() # Haupt
        self.list_vsr1 = ObservableCollection[VSR]() # VSR
        self.list_vsr2 = ObservableCollection[VSR]() # Iris
        for el in self.list_vsr:
            if el.IsIris == True:
                self.list_vsr2.Add(el)
            elif el.IsHaupt == True:
                self.list_vsr0.Add(el) 
            else:
                self.list_vsr1.Add(el)

        if self.list_vsr0.Count > 0:
            for el in self.list_vsr0:
                for auslass in el.Auslass:
                    iris = auslass.IRIS_Class
                    vsr = auslass.VSR_Class
                    if iris is not None:
                        if el.List_Iris.Contains(iris) == False:
                            el.List_Iris.Add(iris)
                            iris.List_Haupt.Add(el)
                    if vsr is not None:
                        if el.List_VSR.Contains(vsr) == False:
                            el.List_VSR.Add(vsr)
                            vsr.List_Haupt.Add(el)
                        if iris is not None:
                            if vsr.List_Iris.Contains(iris) == False:
                                vsr.List_Iris.Add(iris)
                                iris.List_VSR.Add(el)

        else:
            if self.list_vsr1.Count > 0:
                for el in self.list_vsr1:
                    for auslass in el.Auslass:
                        iris = auslass.IRIS_Class
                        if iris is not None:
                            if el.List_Iris.Contains(iris) == False:
                                el.List_Iris.Add(iris)
                                iris.List_VSR.Add(el)

        # self.angezuluft = 0
        # self.angeabluft = 0

        self.IsTierRaum = (self.get_element('IGF_RLT_Tierkäfig_raumunabhängig') != 0)
        self.IsSchacht = (self.get_element('TGA_RLT_InstallationsSchacht') != 0)
        self.list_auslass = ObservableCollection[Luftauslass]()
        if self.elemid.ToString() in DICT_MEP_AUSLASS.keys():
             for art in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()].keys()):
                for fam in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()][art].keys()):
                    for terminal in DICT_MEP_AUSLASS[self.elemid.ToString()][art][fam]:
                        self.list_auslass.Add(terminal)

        # try:self.list_auslass = DICT_MEP_AUSLASS[self.elemid.ToString()]
        # except:self.list_auslass = ObservableCollection[Luftauslass]()
        
        try:self.list_ueber = DICT_MEP_UEBERSTROM[self.elemid.ToString()]
        except:self.list_ueber = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        
        # Grundinfo.
        self.bezugsnummer = self.elem.LookupParameter('TGA_RLT_VolumenstromProNummer').AsValueString()
        try:self.bezugsname = self.berechnung_nach[self.bezugsnummer]
        except:self.bezugsname = 'keine'
        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''
        self.flaeche = round(self.get_element('Fläche')+0.5, 0)
        self.volumen = round(self.get_element('Volumen'),1)
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('TGA_RLT_VolumenstromProFaktor')
        try:self.einheit = self.einheit_liste[self.bezugsnummer]
        except:self.einheit = ''
        self.hoehe = int(self.get_element('Lichte Höhe'))
        self.nachtbetrieb = self.get_element('IGF_RLT_Nachtbetrieb')
        self.tiefenachtbetrieb = self.get_element('IGF_RLT_TieferNachtbetrieb')
        self.NB_LW = self.get_element('IGF_RLT_NachtbetriebLW')
        self.T_NB_LW = self.get_element('IGF_RLT_TieferNachtbetriebLW')

        
        # Übersicht
        self.Uebersicht = ObservableCollection[MEPGrundInfo]()
        self.angezuluft = MEPGrundInfo('Ange.Zuluft',self.get_element('Angegebener Zuluftstrom'),'Angegebener Zuluftstrom')
        self.Uebersicht.Add(self.angezuluft)
        self.angeabluft = MEPGrundInfo('Ange.Abluft',self.get_element('Angegebener Abluftluftstrom'),'Angegebener Abluftluftstrom')
        self.Uebersicht.Add(self.angeabluft)

        self.ab_24h = MEPGrundInfo('24h Ab',self.get_element('IGF_RLT_AbluftSumme24h'),'IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_24h)

        self.ab_minsum = MEPGrundInfo('min.AB_SUM',0,'RAB,LAB,24h,Über,Über_M')
        self.Uebersicht.Add(self.ab_minsum)
        self.ab_min = MEPGrundInfo('min.RAB',self.get_element('IGF_RLT_AbluftminRaum'),'IGF_RLT_AbluftminRaum-IGF_RLT_AbluftminSummeLabor-IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_min)
        self.ab_lab_min = MEPGrundInfo('min.LAB',self.get_element('IGF_RLT_AbluftminSummeLabor'),'IGF_RLT_AbluftminSummeLabor')
        self.Uebersicht.Add(self.ab_lab_min)
        self.zu_minsum = MEPGrundInfo('min.ZU_SUM',0,'RZU,Über,Über_M')
        self.Uebersicht.Add(self.zu_minsum)
        self.zu_min = MEPGrundInfo('min.RZU',self.get_element('IGF_RLT_ZuluftminRaum'),'IGF_RLT_ZuluftminRaum')
        self.Uebersicht.Add(self.zu_min)
        
        self.ab_maxsum = MEPGrundInfo('max.AB_SUM',0,'RAB,LAB,24h,Über,Über_M')
        self.Uebersicht.Add(self.ab_maxsum)
        self.ab_max = MEPGrundInfo('max.RAB',self.get_element('IGF_RLT_AbluftmaxRaum'),'IGF_RLT_AbluftmaxRaum-IGF_RLT_AbluftmaxSummeLabor-IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_max)
        self.ab_lab_max = MEPGrundInfo('max.LAB',self.get_element('IGF_RLT_AbluftmaxSummeLabor'),'IGF_RLT_AbluftmaxSummeLabor')
        self.Uebersicht.Add(self.ab_lab_max)
        self.zu_maxsum = MEPGrundInfo('max.ZU_SUM',0,'RZU,Über,Über_M')
        self.Uebersicht.Add(self.zu_maxsum)
        self.zu_max = MEPGrundInfo('max.RZU',self.get_element('IGF_RLT_ZuluftmaxRaum'),'IGF_RLT_ZuluftmaxRaum')
        self.Uebersicht.Add(self.zu_max)

        self.nb_von = MEPGrundInfo('NB von',self.get_element('IGF_RLT_NachtbetriebVon'),'IGF_RLT_NachtbetriebVon')
        self.Uebersicht.Add(self.nb_von)
        self.nb_bis = MEPGrundInfo('NB bis',self.get_element('IGF_RLT_NachtbetriebBis'),'IGF_RLT_NachtbetriebBis')
        self.Uebersicht.Add(self.nb_bis)
        self.nb_dauer = MEPGrundInfo('NB Dauer',self.get_element('IGF_RLT_NachtbetriebDauer'),'IGF_RLT_NachtbetriebDauer')
        self.Uebersicht.Add(self.nb_dauer)
        self.nb_zu = MEPGrundInfo('NB Zu',self.get_element('IGF_RLT_ZuluftNachtRaum'),'IGF_RLT_ZuluftNachtRaum')
        self.Uebersicht.Add(self.nb_zu)
        self.nb_ab = MEPGrundInfo('NB Ab',self.get_element('IGF_RLT_AbluftNachtRaum'),'IGF_RLT_AbluftNachtRaum')
        self.Uebersicht.Add(self.nb_ab)

        self.tnb_von = MEPGrundInfo('TNB von',self.get_element('IGF_RLT_TieferNachtbetriebVon'),'IGF_RLT_TieferNachtbetriebVon')
        self.Uebersicht.Add(self.tnb_von)
        self.tnb_bis = MEPGrundInfo('TNB bis',self.get_element('IGF_RLT_TieferNachtbetriebBis'),'IGF_RLT_TieferNachtbetriebBis')
        self.Uebersicht.Add(self.tnb_bis)
        self.tnb_dauer = MEPGrundInfo('TNB Dauer',self.get_element('IGF_RLT_TieferNachtbetriebDauer'),'IGF_RLT_TieferNachtbetriebDauer')
        self.Uebersicht.Add(self.tnb_dauer)
        self.tnb_zu = MEPGrundInfo('TNB Zu',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'),'IGF_RLT_ZuluftTieferNachtRaum')
        self.Uebersicht.Add(self.tnb_zu)
        self.tnb_ab = MEPGrundInfo('TNB Ab',self.get_element('IGF_RLT_AbluftTieferNachtRaum'),'IGF_RLT_AbluftTieferNachtRaum')
        self.Uebersicht.Add(self.tnb_ab)

        if self.IsTierRaum != 0:
            self.tier_zu_min = MEPGrundInfo('min.TZU',self.get_element('IGF_RLT_Luftmenge_min_TZU'),'IGF_RLT_Luftmenge_min_TZU')
            self.Uebersicht.Add(self.tier_zu_min)
            self.tier_ab_min = MEPGrundInfo('min.TAB',self.get_element('IGF_RLT_Luftmenge_min_TAB'),'IGF_RLT_Luftmenge_min_TAB')
            self.Uebersicht.Add(self.tier_ab_min)
            self.tier_zu_max = MEPGrundInfo('max.TZU',self.get_element('IGF_RLT_Luftmenge_max_TZU'),'IGF_RLT_Luftmenge_max_TZU')
            self.Uebersicht.Add(self.tier_zu_max)
            self.tier_ab_max = MEPGrundInfo('max.TAB',self.get_element('IGF_RLT_Luftmenge_max_TAB'),'IGF_RLT_Luftmenge_max_TAB')
            self.Uebersicht.Add(self.tier_ab_max)
        else:
            self.tier_zu_min = MEPGrundInfo('min.TZU',0,'IGF_RLT_Luftmenge_min_TZU')
            self.Uebersicht.Add(self.tier_zu_min)
            self.tier_ab_min = MEPGrundInfo('min.TAB',0,'IGF_RLT_Luftmenge_min_TAB')
            self.Uebersicht.Add(self.tier_ab_min)
            self.tier_zu_max = MEPGrundInfo('max.TZU',0,'IGF_RLT_Luftmenge_max_TZU')
            self.Uebersicht.Add(self.tier_zu_max)
            self.tier_ab_max = MEPGrundInfo('max.TAB',0,'IGF_RLT_Luftmenge_max_TAB')
            self.Uebersicht.Add(self.tier_ab_max)
        

        self.ueber_sum = MEPGrundInfo('Überstrom',self.get_element('IGF_RLT_ÜberströmungRaum'),'IGF_RLT_ÜberströmungRaum')
        self.Uebersicht.Add(self.ueber_sum)
        self.ueber_in = MEPGrundInfo('Über. Ein.',self.get_element('IGF_RLT_ÜberströmungSummeIn'),'IGF_RLT_ÜberströmungSummeIn')
        self.Uebersicht.Add(self.ueber_in)
        self.ueber_aus = MEPGrundInfo('Über. Aus.',self.get_element('IGF_RLT_ÜberströmungSummeAus'),'IGF_RLT_ÜberströmungSummeAus')
        self.Uebersicht.Add(self.ueber_aus)
        self.ueber_in_manuell = MEPGrundInfo('Über.Ein.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        self.Uebersicht.Add(self.ueber_in_manuell)
        self.ueber_aus_manuell = MEPGrundInfo('Über.Aus.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        try:self.ueber_aus_manuell.soll = Dict_Ueber_Manuell[self.elem.Number]
        except:self.ueber_aus_manuell.soll = 0
        self.Uebersicht.Add(self.ueber_aus_manuell)
        self.ueber_aus_manuell.ist =self.ueber_aus_manuell.soll
        self.ueber_in_manuell.ist = self.ueber_in_manuell.soll

        self.ab_min.soll = self.ab_min.soll - self.ab_24h.soll - self.ab_lab_min.soll
        self.ab_max.soll = self.ab_max.soll - self.ab_24h.soll - self.ab_lab_max.soll
        self.sum_update()

        
        
        self.IGF_Legende = ''     
            
        self.Druckstufe = MEPGrundInfo('Druckstufe',self.get_element('IGF_RLT_RaumDruckstufeEingabe'),'IGF_RLT_RaumDruckstufeEingabe')
        self.Uebersicht.Add(self.Druckstufe)
        # Anlagen
        self.Anlagen_info = ObservableCollection[MEPAnlagenInfo]()

        # Schacht
        self.Schacht_info = ObservableCollection[MEPSchachtInfo]()
    
        self.Analyse()
        self.get_Anlagen_info()
        self.get_Schacht_info()
    
    def update(self):        
        self.ab_24h.soll = self.get_element('IGF_RLT_AbluftSumme24h')
        self.ab_lab_min.soll = self.get_element('IGF_RLT_AbluftminSummeLabor') 
        self.Druckstufe.soll = self.get_element('IGF_RLT_RaumDruckstufeEingabe')
    
    def sum_update(self):
        self.ab_minsum.soll = str(int(round(float(self.ab_min.soll)+float(self.ab_lab_min.soll)+float(self.ab_24h.soll)+float(self.ueber_aus.soll)+float(self.ueber_aus_manuell.soll)))) + \
            ' (' + str(int(self.ab_min.soll)) + ', ' + str(int(self.ab_lab_min.soll)) + ', ' + str(int(self.ab_24h.soll)) + ', ' + str(int(self.ueber_aus.soll)) + ', ' + str(int(self.ueber_aus_manuell.soll))+ ')'
        
        self.ab_maxsum.soll = str(int(round(float(self.ab_max.soll)+float(self.ab_lab_max.soll)+float(self.ab_24h.soll)+float(self.ueber_aus.soll)+float(self.ueber_aus_manuell.soll)))) + \
            ' (' + str(int(self.ab_max.soll)) + ', ' + str(int(self.ab_lab_max.soll)) + ', ' + str(int(self.ab_24h.soll)) + ', ' + str(int(self.ueber_aus.soll)) + ', ' + str(int(self.ueber_aus_manuell.soll))+ ')'
        
        self.zu_minsum.soll = str(int(round(float(self.zu_min.soll)+float(self.ueber_in.soll)+float(self.ueber_in_manuell.soll)))) + \
            ' (' + str(int(self.zu_min.soll)) + ', ' + str(int(self.ueber_in.soll)) + ', ' + str(int(self.ueber_in_manuell.soll))+ ')'
        
        self.zu_maxsum.soll = str(int(round(float(self.zu_max.soll)+float(self.ueber_in.soll)+float(self.ueber_in_manuell.soll)))) + \
            ' (' + str(int(self.zu_max.soll)) + ', ' + str(int(self.ueber_in.soll)) + ', ' + str(int(self.ueber_in_manuell.soll)) + ')'
        
        # self.ab_minsum.ist = str(int(round(float(self.ab_min.ist)+float(self.ab_lab_min.ist)+float(self.ab_24h.ist)+float(self.ueber_aus.ist)+float(self.ueber_aus_manuell.ist)))) + \
        #     ' (' + str(int(self.ab_min.ist)) + ', ' + str(int(self.ab_lab_min.ist)) + ', ' + str(int(self.ab_24h.ist)) + ', ' + str(int(self.ueber_aus.ist)) + ', ' + str(int(self.ueber_aus_manuell.ist))+ ')'
        
        # self.ab_maxsum.ist = str(int(round(float(self.ab_max.ist)+float(self.ab_lab_max.ist)+float(self.ab_24h.ist)+float(self.ueber_aus.ist)+float(self.ueber_aus_manuell.ist)))) + \
        #     ' (' + str(int(self.ab_max.ist)) + ', ' + str(int(self.ab_lab_max.ist)) + ', ' + str(int(self.ab_24h.ist)) + ', ' + str(int(self.ueber_aus.ist)) + ', ' + str(int(self.ueber_aus_manuell.ist))+ ')'
        
        # self.zu_minsum.ist = str(int(round(float(self.zu_min.ist)+float(self.ueber_in.ist)+float(self.ueber_in_manuell.ist)))) + \
        #     ' (' + str(int(self.zu_min.ist)) + ', ' + str(int(self.ueber_in.ist)) + ', ' + str(int(self.ueber_in_manuell.ist))+ ')'
        
        # self.zu_maxsum.ist = str(int(round(float(self.zu_max.ist)+float(self.ueber_in.ist)+float(self.ueber_in_manuell.ist)))) + \
        #     ' (' + str(int(self.zu_max.ist)) + ', ' + str(int(self.ueber_in.ist)) + ', ' + str(int(self.ueber_in_manuell.ist)) + ')'

    def luft_round(self,luft):
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Druckstufe_Berechnen(self):
        n = abs(int(self.Druckstufe.soll/10)) if abs(int(self.Druckstufe.soll/10)) < 6 else 5
        if self.Druckstufe.soll > 0:self.IGF_Legende = n*'+'
        else:self.IGF_Legende = n*'-'
    
    def Nachtbetrieb_Berechnen(self):
        self.TiefeNachtbetrieb_Berechnen()
        if self.nachtbetrieb:
            if self.tiefenachtbetrieb == 0:
                if self.nb_bis.soll - self.nb_von.soll > 0:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll
                else:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll + 24.00
            else:
                if self.nb_bis.soll - self.nb_von.soll > 0:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll
                else:
                    self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll + 24.00 - self.tnb_dauer.soll
            if self.bezugsname in ["Fläche","Luftwechsel","Person","manuell"]:
                self.nb_zu.soll = self.luft_round(float(self.NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.nb_zu.soll -= abweichung
                    self.nb_ab.soll = 0
                else:
                    self.nb_ab.soll = abweichung
            elif self.bezugsname in ["nurZU_Fläche","nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
                self.nb_zu.soll = self.luft_round(float(self.NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.nb_ab.soll = 0
                else:
                    self.nb_ab.soll = 0
                self.nb_zu.soll -= abweichung
            
            elif self.bezugsname in ["nurAB_Fläche","nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
                self.nb_zu.soll = 0
                abweichung = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll + self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.nb_zu.soll -= abweichung
                    self.nb_ab.soll = 0
                else:
                    self.nb_ab.soll = abweichung 
        else:
            self.nb_zu.soll = 0
            self.nb_ab.soll = 0
    def TiefeNachtbetrieb_Berechnen(self):
        if self.tiefenachtbetrieb and self.nachtbetrieb:
            if self.tnb_bis.soll - self.tnb_von.soll > 0:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
            else:
                self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll + 24.00
            if self.bezugsname in ["Fläche","Luftwechsel","Person","manuell"]:
                self.tnb_zu.soll = self.luft_round(float(self.T_NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.tnb_zu.soll -= abweichung
                    self.tnb_ab.soll = 0
                else:
                    self.tnb_ab.soll = abweichung
            elif self.bezugsname in ["nurZU_Fläche","nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
                self.tnb_zu.soll = self.luft_round(float(self.T_NB_LW) * self.volumen) + self.Druckstufe.soll
                abweichung = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll - self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.tnb_ab.soll = 0
                else:
                    self.tnb_ab.soll = 0
                self.tnb_zu.soll -= abweichung
            
            elif self.bezugsname in ["nurAB_Fläche","nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
                self.tnb_zu.soll = 0
                abweichung = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - \
                self.ab_lab_min.soll + self.Druckstufe.soll - self.ab_24h.soll
                if abweichung < 0:
                    self.tnb_zu.soll -= abweichung
                    self.tnb_ab.soll = 0
                else:
                    self.tnb_ab.soll = abweichung 
        else:
            self.tnb_zu.soll = 0
            self.tnb_ab.soll = 0

    def Tagesbetrieb_Berechnen(self):
        if self.flaeche == 0:
            return
        if self.bezugsname == 'nurZU_Fläche':
            self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            if abweichung < 0:
                self.logger.error('Bitte Überstrom-Aus erhöhen in Raum {}. min. Zuluft {} m³/h.'.format(self.Raumnr,self.zu_min.soll))
            else:
                self.zu_min.soll += abweichung
            self.ab_min.soll = 0
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            
            if abweichung >= 0:self.zu_max.soll += abweichung
            self.ab_max.soll = 0

        elif self.bezugsname == 'nurZU_Luftwechsel':
            self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            if abweichung < 0:self.logger.error('Bitte Überstrom-Aus erhöhen in Raum {}. min. Zuluft {} m³/h.'.format(self.Raumnr,self.zu_min.soll))
            else:self.zu_min.soll += abweichung
            self.ab_min.soll = 0
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            if abweichung >= 0:self.zu_max.soll += abweichung
            self.ab_max.soll = 0
        
        elif self.bezugsname == 'nurZU_Person':
            self.zu_min.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            if abweichung < 0:self.logger.error('Bitte Überstrom-Aus erhöhen in Raum {}. min. Zuluft {} m³/h.'.format(self.Raumnr,self.zu_min.soll))
            else:self.zu_min.soll += abweichung
            self.ab_min.soll = 0
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            if abweichung >= 0:self.zu_max.soll += abweichung
            self.ab_max.soll = 0
        
        elif self.bezugsname == 'nurAB_Fläche':
            self.angezuluft.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll > self.angezuluft.soll:
                    self.angezuluft.soll -= abweichung2
                    self.ab_min.soll = 0
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll
                else:
                    self.ab_min.soll = 0
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll

            else:
                if self.ueber_in.soll + self.ueber_in_manuell.soll < self.angezuluft.soll:
                    logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                self.ab_min.soll = abweichung2
                self.zu_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
                self.zu_max.soll = 0
            else:
                self.ab_max.soll = abweichung2
                self.zu_max.soll = 0
                
        elif self.bezugsname == 'nurAB_Luftwechsel':
            self.angezuluft.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll > self.angezuluft.soll:
                    self.angezuluft.soll -= abweichung2
                    self.ab_min.soll = 0
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll
                else:
                    self.ab_min.soll = 0-abweichung2
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll

            else:
                if self.ueber_in.soll + self.ueber_in_manuell.soll < self.angezuluft.soll:
                    logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                self.ab_min.soll = abweichung2
                self.zu_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
                self.zu_max.soll = 0


            else:
                self.ab_max.soll = abweichung2
                self.zu_max.soll = 0
       
        elif self.bezugsname == 'nurAB_Person':
            self.angezuluft.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll > self.angezuluft.soll:
                    self.angezuluft.soll -= abweichung2
                    self.ab_min.soll = 0
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll
                else:
                    self.ab_min.soll = 0-abweichung2
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll

            else:
                if self.ueber_in.soll + self.ueber_in_manuell.soll < self.angezuluft.soll:
                    logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                self.ab_min.soll = abweichung2
                self.zu_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
                self.ab_min.soll = abweichung2
                self.zu_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
                self.zu_max.soll = 0
            else:
                self.ab_max.soll = abweichung2
                self.zu_max.soll = 0
                
        elif self.bezugsname == 'Fläche':
            self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.flaeche * float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung
           
        elif self.bezugsname == 'Luftwechsel':
            self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.volumen * float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung

        elif self.bezugsname == 'Person':
            self.zu_min.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(self.personen * float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
            else:
                self.ab_max.soll = abweichung
  
        elif self.bezugsname == 'manuell':
            self.zu_min.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            abweichung = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_min.soll -= abweichung
                self.angezuluft.soll -= abweichung
                self.ab_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
            else:
                self.ab_min.soll = abweichung
                self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            abweichung = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung < 0:
                self.zu_max.soll -= abweichung
                self.ab_max.soll = 0
                
            else:
                self.ab_max.soll = abweichung
        
        elif self.bezugsname == 'nurZUMa':
            self.zu_min.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.zu_min.soll
            if self.ab_lab_min.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll
            if abweichung < 0:self.logger.error('Bitte Überstrom-Aus erhöhen in Raum {}. min. Zuluft {} m³/h.'.format(self.Raumnr,self.zu_min.soll))
            else:self.zu_min.soll += abweichung
            self.ab_min.soll = 0
            self.angezuluft.soll = self.zu_min.soll
            self.angeabluft.soll = self.angezuluft.soll

            self.zu_max.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            if self.ab_lab_max.soll + self.ab_24h.soll> 0:
                self.logger.error('Berechnungsprinzip von Raum {} ist Falsch, da Laborabluft/24h Abluft vorhanden.'.format(self.Raumnr))
            abweichung = self.ueber_aus.soll - self.zu_max.soll - self.ueber_in.soll - self.ueber_in_manuell.soll + \
                self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll
            if abweichung >= 0:self.zu_max.soll += abweichung
            self.ab_max.soll = 0
                
        elif self.bezugsname == 'nurABMa':
            self.ab_min.soll = self.luft_round(float(self.faktor)) + self.Druckstufe.soll
            self.angezuluft.soll = self.ab_min.soll
            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_min.soll + self.ab_24h.soll > self.angezuluft.soll:
                    self.angezuluft.soll -= abweichung2
                    self.ab_min.soll = 0
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll
                else:
                    self.ab_min.soll = 0-abweichung2
                    self.zu_min.soll = 0
                    self.logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                    self.angeabluft.soll = self.angezuluft.soll

            else:
                if self.ueber_in.soll + self.ueber_in_manuell.soll < self.angezuluft.soll:
                    logger.error('Bitte Überstrom-Ein erhöhen in Raum {}. min. Abluft {} m³/h.'.format(self.Raumnr,self.angezuluft.soll))
                self.ab_min.soll = abweichung2
                self.zu_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll
                
                self.ab_min.soll = abweichung2
                self.zu_min.soll = 0
                self.angeabluft.soll = self.angezuluft.soll

            abweichung2 = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll- self.Druckstufe.soll
            if abweichung2 < 0:
                self.ab_max.soll = 0
                self.zu_max.soll = 0
            else:
                self.ab_max.soll = abweichung2
                self.zu_max.soll = 0

        elif self.bezugsname == 'keine':
            self.zu_min.soll = 0
            self.angezuluft.soll = 0
            self.angeabluft.soll = 0
            self.ab_min.soll = 0
            self.zu_max.soll = 0
            self.ab_max.soll = 0
            
        self.get_Schacht_info()
        self.get_Anlagen_info()
        self.sum_update()

    def Analyse(self):
        min_zu = 0
        min_ab = 0
        max_zu = 0
        max_ab = 0
        ab24h = 0
        lab_min = 0
        lab_max = 0
        nb_zu = 0
        nb_ab = 0
        tnb_zu = 0
        tnb_ab = 0
        uber_in = 0
        uber_aus = 0

        for uber in self.list_ueber["Ein"]:
            uber_in += uber.menge
        for uber in self.list_ueber["Aus"]:
            uber_aus += uber.menge

        for auslass in self.list_auslass:
            if auslass.art == '24h':
                ab24h += auslass.Luftmengenmin
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'LAB':
                lab_min += auslass.Luftmengenmin
                lab_max += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'RZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentnacht
            elif auslass.art == 'RAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'TZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentnacht
            elif auslass.art == 'TAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
        
        self.zu_min.ist = int(round(min_zu))
        self.ab_min.ist = int(round(min_ab))
        self.zu_max.ist = int(round(max_zu))
        self.ab_max.ist = int(round(max_ab))
        self.ab_24h.ist = int(round(ab24h))
        self.ab_lab_min.ist = int(round(lab_min))
        self.ab_lab_max.ist = int(round(lab_max))
        self.nb_zu.ist = int(round(nb_zu))
        self.nb_ab.ist = int(round(nb_ab))
        self.tnb_zu.ist = int(round(tnb_zu))
        self.tnb_ab.ist = int(round(tnb_ab))
        self.ueber_in.ist = int(round(uber_in))
        self.ueber_aus.ist = int(round(uber_aus))
        self.ueber_sum.ist = int(round(uber_in-uber_aus))
        
    def get_Anlagen_info(self):
        self.Anlagen_info.Clear()
        Dict = {}
        for el in self.list_auslass:
            if not el.art in Dict.keys():
                Dict[el.art] = {}
            if not el.AnlNr in Dict[el.art].keys():
                Dict[el.art][el.AnlNr] = float(el.Luftmengenmax)
            else:
                Dict[el.art][el.AnlNr] += float(el.Luftmengenmax)
        
        if 'RZU' in Dict.keys():
            for anl in sorted(Dict['RZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                    self.zu_max.soll,anl,Dict['RZU'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_RZU') or self.get_element('IGF_RLT_Luftmenge_RZU'):
            self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                self.zu_max.soll,'',''))
        if 'RAB' in Dict.keys():
            for anl in sorted(Dict['RAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                    self.ab_max.soll,anl,Dict['RAB'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_RAB') or self.get_element('IGF_RLT_Luftmenge_RAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                self.ab_max.soll,'',''))
        if 'TZU' in Dict.keys():
            for anl in sorted(Dict['TZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                    self.tier_zu_max.soll,anl,Dict['TZU'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_TZU') or self.get_element('IGF_RLT_Luftmenge_TZU'):
            self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                self.tier_zu_max.soll,'',''))
        
        if 'TAB' in Dict.keys():
            for anl in sorted(Dict['TAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                    self.tier_ab_max.soll,anl,Dict['TAB'][anl]))
        else:
            # if self.get_element('IGF_RLT_AnlagenNr_TAB') or self.get_element('IGF_RLT_Luftmenge_TAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                self.tier_ab_max.soll,'',''))
        
        if '24h' in Dict.keys():
            for anl in sorted(Dict['24h'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                    self.ab_24h.soll,anl,Dict['24h'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_24h') or self.get_element('IGF_RLT_Luftmenge_24h'):
            self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                self.ab_24h.soll,'',''))
           
        if 'LAB' in Dict.keys():
            for anl in sorted(Dict['LAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                    self.ab_lab_max.soll,anl,Dict['LAB'][anl]))  
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_LAB') or self.get_element('IGF_RLT_Luftmenge_LAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                self.ab_lab_max.soll,'',''))
        
    def get_Schacht_info(self):
        self.Schacht_info.Clear()
        self.rzu_Schacht = MEPSchachtInfo('RZU',self.get_element('TGA_RLT_SchachtZuluft'),self.zu_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.rzu_Schacht)
        self.rab_Schacht = MEPSchachtInfo('RAB',self.get_element('TGA_RLT_SchachtAbluft'),self.ab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.rab_Schacht)
        
        self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.tier_zu_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.tzu_Schacht)
        self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.tier_ab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.tab_Schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('TGA_RLT_Schacht24hAbluft'),self.ab_24h.soll,self.liste_schacht)
        self.Schacht_info.Add(self._24h_Schacht)
        self.lab_Schacht = MEPSchachtInfo('LAB',self.get_element('IGF_RLT_Schacht_LAB'),self.ab_lab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.lab_Schacht)
        
    def get_element(self, param_name):
        param = self.elem.LookupParameter(param_name)
        if not param:
            self.logger.info("Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return get_value(param)
    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            try:
                if not wert is None:
                #  logger.info(
                #      "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                    if self.elem.LookupParameter(param_name):
                        if self.elem.LookupParameter(param_name).IsReadOnly is True:
                            self.logger.error(self.elemid)
                            self.logger.error(param_name)
                            return
                        self.elem.LookupParameter(param_name).SetValueString(str(wert))
                    else:
                        print(param_name)
            except Exception as e:print(e,0)
        def wert_schreiben2(param_name, wert):
            try:
                if self.elem.LookupParameter(param_name):
                    if self.elem.LookupParameter(param_name).IsReadOnly is True:
                        self.logger.error(self.elemid)
                        self.logger.error(param_name)
                        return
                    self.elem.LookupParameter(param_name).Set(wert)
                else:
                    print(param_name)
            except Exception as e:print(e,1)  
        def wert_schreiben3(param_name, wert):
            '''für Schacht'''
            try:
                if self.elem.LookupParameter(param_name):
                    if wert.schachtindex != -1:
                        self.elem.LookupParameter(param_name).Set(wert.SchachtListe[wert.schachtindex])
                else:
                    print(param_name)
                  
            except Exception as e:print(e)

        wert_schreiben("Angegebener Zuluftstrom", self.angezuluft.soll)
        wert_schreiben("Angegebener Abluftluftstrom", self.angeabluft.soll)
        wert_schreiben("IGF_RLT_AbluftminRaum", self.ab_min.soll+self.ab_lab_min.soll+self.ab_24h.soll+self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_AbluftmaxRaum", self.ab_max.soll+self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zu_min.soll)
        wert_schreiben("IGF_RLT_ZuluftmaxRaum", self.zu_max.soll)

        wert_schreiben2("TGA_RLT_VolumenstromProName", self.bezugsname)
        wert_schreiben("TGA_RLT_VolumenstromProEinheit", self.einheit)
        wert_schreiben("TGA_RLT_VolumenstromProNummer", self.bezugsnummer)
        wert_schreiben("TGA_RLT_VolumenstromProFaktor", float(self.faktor))

        wert_schreiben2("IGF_RLT_Nachtbetrieb", self.nachtbetrieb)
        wert_schreiben("IGF_RLT_NachtbetriebLW", self.NB_LW)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.nb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.nb_ab.soll)
        wert_schreiben("IGF_RLT_NachtbetriebVon", self.nb_von.soll)
        wert_schreiben("IGF_RLT_NachtbetriebBis", self.nb_bis.soll)

        wert_schreiben2("IGF_RLT_TieferNachtbetrieb", self.tiefenachtbetrieb)
        wert_schreiben("IGF_RLT_TieferNachtbetriebLW", self.T_NB_LW)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer", self.tnb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tnb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tnb_ab.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebVon", self.tnb_von.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebBis", self.tnb_bis.soll)

        wert_schreiben("IGF_RLT_AbluftSumme24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_AbluftminRaumL24h", self.ab_24h.soll)    
        wert_schreiben("IGF_RLT_AbluftminSummeLabor", self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftminSummeLabor24h", self.ab_24h.soll + self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor24h", self.ab_24h.soll + self.ab_lab_max.soll)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor", self.ab_lab_max.soll)

        wert_schreiben("IGF_RLT_RaumDruckstufeEingabe", self.Druckstufe.soll)
        wert_schreiben2("IGF_RLT_RaumDruckstufeLegende", self.IGF_Legende)
        
        # wert_schreiben("IGF_RLT_AnlagenRaumAbluft", self.ab_min.soll+self.ab_lab_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaumZuluft", self.zu_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaum24hAbluft", self.ab_24h.soll)
        
        wert_schreiben("IGF_RLT_Luftmenge_RAB", self.ab_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_RZU", self.zu_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_Luftmenge_LAB", self.ab_lab_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TAB", self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TZU", self.tier_zu_min.soll)
        
        wert_schreiben3('TGA_RLT_SchachtZuluft',self.rzu_Schacht)
        wert_schreiben3('TGA_RLT_SchachtAbluft',self.rab_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TZU',self.tzu_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TAB',self.tab_Schacht)
        wert_schreiben3('TGA_RLT_Schacht24hAbluft',self._24h_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_LAB',self.lab_Schacht)

        for el in self.Anlagen_info:
            if el.name == 'RZU':
                wert_schreiben2('IGF_RLT_AnlagenNr_RZU',int(el.mep_nr))
            elif el.name == 'RAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_RAB',int(el.mep_nr))
            elif el.name == 'TZU':
                wert_schreiben2('IGF_RLT_AnlagenNr_TZU',int(el.mep_nr))
            elif el.name == 'TAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_TAB',int(el.mep_nr))
            elif el.name == 'LAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_LAB',int(el.mep_nr))
            elif el.name == '24h':
                wert_schreiben2('IGF_RLT_AnlagenNr_24h',int(el.mep_nr))
       
DICT_MEP_ITEMSSOIRCE = {}

def Get_MEP_Info():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()
    
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb:
        for n,space_id in enumerate(spaces):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n+1, len(spaces))
            
            if space_id.ToString() in DICT_MEP_VSR.keys():
                list_vsr = ObservableCollection[VSR]()
                dict_vsr = {}
                for e in DICT_MEP_VSR[space_id.ToString()]:
                    temp = DICT_VSR[e]
                    if temp.art not in dict_vsr.keys():
                        dict_vsr[temp.art] = {}
                    if temp.familyandtyp not in dict_vsr[temp.art].keys():
                        dict_vsr[temp.art][temp.familyandtyp] = []
                    dict_vsr[temp.art][temp.familyandtyp].append(temp)
                for art in sorted(dict_vsr.keys()):
                    for fam in sorted(dict_vsr[art].keys()):
                        for kla in dict_vsr[art][fam]:
                            try:list_vsr.Add(kla)
                            except:pass
            else:list_vsr = ObservableCollection[VSR]()

            mepraum = MEPRaum(space_id,list_vsr)
            # if not mepraum.IsSchacht:
            DICT_MEP_ITEMSSOIRCE[mepraum.Raumnr] = mepraum

Get_MEP_Info()

class ExtenalEventListe(IExternalEventHandler):
    def __init__(self):
        self.name = ''
        self.XYZAnpassen = DB.XYZ(1/3.048,1/3.048,1/3.048)
        self.GUI = None
        self.Executeapp = None
        self.Liste = []
        
    def Execute(self,uiapp):
        try:
            self.Executeapp(uiapp)
        except Exception as e:
            TaskDialog.Show('Fehler',e.ToString())

    def GetName(self):
        return self.name
    
    def ChangeTo3D(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        views = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.View3D)).ToElements()
        for el in views:
            if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
                uidoc.ActiveView = el
                break
        
    def SelectedElements(self,uiapp,Liste):
        uiapp.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId](Liste))
    
    def SelectElements(self,uiapp):
        self.name = 'Element auswählen'
        uiapp.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId](self.Liste))
    
    def RaumAnzeigen(self,uiapp):
        self.name = 'Raum anzeigen'
        self.ChangeTo3D(uiapp)
        self.SelectedElements(uiapp,[self.GUI.mepraum.elemid])
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        t = DB.Transaction(doc,self.name)
        t.Start()
        view.IsSectionBoxActive = True
        box = view.GetSectionBox()
        Max = box.Transform.Inverse.OfPoint(self.GUI.mepraum.box.Max + self.XYZAnpassen)
        Min = box.Transform.Inverse.OfPoint(self.GUI.mepraum.box.Min - self.XYZAnpassen)
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()
        t.Dispose()
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                return

    def RaumundBauteileanzeigen(self,uiapp):
        self.name = 'Raum und alle Bauteile anzeigen'
        self.ChangeTo3D(uiapp)
        self.SelectedElements(uiapp,[self.GUI.mepraum.elemid])
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        t = DB.Transaction(doc,self.name)
        t.Start()
        view.IsSectionBoxActive = True
        box = view.GetSectionBox()
        max_x,max_y,max_z = self.GUI.mepraum.box.Max.X,self.GUI.mepraum.box.Max.Y,self.GUI.mepraum.box.Max.Z
        min_x,min_y,min_z = self.GUI.mepraum.box.Min.X,self.GUI.mepraum.box.Min.Y,self.GUI.mepraum.box.Min.Z
        for el in self.GUI.mepraum.list_vsr:
            box_temp = el.elem.get_BoundingBox(None)
            max_x = max(max_x,box_temp.Max.X)
            max_y = max(max_y,box_temp.Max.Y)
            max_z = max(max_z,box_temp.Max.Z)
            min_x = min(min_x,box_temp.Min.X)
            min_y = min(min_y,box_temp.Min.Y)
            min_z = min(min_z,box_temp.Min.Z)

        Max = box.Transform.Inverse.OfPoint(DB.XYZ(max_x,max_y,max_z) + self.XYZAnpassen)
        Min = box.Transform.Inverse.OfPoint(DB.XYZ(min_x,min_y,min_z) - self.XYZAnpassen)
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()
        t.Dispose()
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                return
    
    def RaumluftBerechnen(self,uiapp):
        self.name = 'Raumluftberechnen'
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        t.Start()
        for el in self.GUI.mepraum_liste.values():          
            try:el.Nachtbetrieb_Berechnen()
            except Exception as e:print(e)
            try:el.Tagesbetrieb_Berechnen()
            except Exception as e:print(e)
            try:el.Druckstufe_Berechnen()
            except Exception as e:print(e)
            try:el.werte_schreiben()
            except Exception as e:print(e)

        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
    
    def LuftmengeverteilenMEP(self,uiapp):
        self.name = 'Luftmengeverteilen ' + self.GUI.mepraum.Raumnr
        task = ABFRAGE('Luftmenge in akt. Raum gleichmäßig verteilen?',True,130)
        task.ShowDialog()
        if task.result == False:
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        self.Luftmengeverteilen(self.GUI.mepraum,task)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
    
    def LuftmengeverteilenProjekt(self,uiapp):
        self.name = 'Luftmengeverteilen Projekt' 
        task = ABFRAGE('Luftmenge für das Projekt gleichmäßig verteilen?',False,160)
        task.ShowDialog()
        if task.result == False:
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.mepraum_liste.values():  
            self.Luftmengeverteilen(mep,task)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
    
    def Luftmengeverteilen(self,mep,task):
        zu = {}
        ab = {}
        h24 = {}
        lab = {}

        for auslass in mep.list_auslass:
            if auslass.art == 'RZU':
                if auslass.familyandtyp not in zu.keys():
                    zu[auslass.familyandtyp] = [auslass]
                else:
                    zu[auslass.familyandtyp].append(auslass)
            if auslass.art == 'RAB':
                if auslass.familyandtyp not in ab.keys():
                    ab[auslass.familyandtyp] = [auslass]
                else:
                    ab[auslass.familyandtyp].append(auslass)
            if auslass.art == '24h':
                if auslass.familyandtyp not in h24.keys():
                    h24[auslass.familyandtyp] = [auslass]
                else:
                    h24[auslass.familyandtyp].append(auslass)
            if auslass.art == 'LAB':
                if auslass.familyandtyp not in lab.keys():
                    lab[auslass.familyandtyp] = [auslass]
                else:
                    lab[auslass.familyandtyp].append(auslass)
        
        if int(mep.ab_24h.soll) != int(mep.ab_24h.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist):
            print(30*'-')
            print('24h-Abluft oder Laborabluft in MEP-Raum {} stimmt nicht übereinandern.'.format(mep.Raumnr))
            print('24h-Abluft-soll: {} m³/h, 24h-Abluft-ist: {} m³/h'.format(mep.ab_24h.soll,mep.ab_24h.ist))
            print('Laborabluftmin-soll: {} m³/h, Laborabluftmin-ist: {} m³/h'.format(mep.ab_lab_min.soll,mep.ab_lab_min.ist))
            print('Laborabluftmax-soll: {} m³/h, Laborabluftmax-ist: {} m³/h'.format(mep.ab_lab_max.soll,mep.ab_lab_max.ist))
            print('Bitte manuell anpassen. ')

        if (int(mep.zu_min.soll) > 0 or int(mep.zu_max.soll) > 0) and len(zu.keys()) == 0:
            print(30*'-')
            print('Es fehlt Zuluftauslass in MEP-Raum {}.'.format(mep.Raumnr))
            print('Zuluft min : {} m³/h, Zuluft max: {} m³/h'.format(mep.zu_min.soll,mep.zu_max.soll))
            print('Bitte manuell anpassen. ')

        if (int(mep.ab_min.soll) > 0 or int(mep.ab_max.soll) > 0) and len(ab.keys()) == 0:
            print(30*'-')
            print('Es fehlt Abluftauslass in MEP-Raum {}.'.format(mep.Raumnr))
            print('Abluft min : {} m³/h, Abluft max: {} m³/h'.format(mep.ab_min.soll,mep.ab_max.soll))
            print('Bitte manuell anpassen. ')
   
        if len(h24.keys()) > 0:
            for key in h24.keys():
                for auslass in h24[key]:
                    auslass.Luftmengennacht = auslass.Luftmengenmin
                    auslass.Luftmengenmax = auslass.Luftmengenmin
                    auslass.Luftmengentnacht = auslass.Luftmengenmin
                    
        if len(zu.keys()) == 1:
            for key in zu.keys():
                for auslass in zu[key]:
                    if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll * 1.0 / len(zu[key]),1)
                    if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll * 1.0 / len(zu[key]),1)
                    if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll * 1.0 / len(zu[key]),1)
                    if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll * 1.0 / len(zu[key]),1)
        
        elif len(zu.keys()) > 1:
            sum_luft = 0
            for key in zu.keys():
                for auslass in zu[key]:
                    if auslass.Luftmengenmin > 0:sum_luft += auslass.Luftmengenmin
                    else:sum_luft += 0.01
            for key in zu.keys():
                for auslass in zu[key]:
                    if auslass.Luftmengenmin > 0:
                        if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                    else:
                        if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll *1.0 * 0.01 / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll *1.0 * 0.01 / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll *1.0 * 0.01 / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll *1.0 * 0.01 / sum_luft,1)

        if len(ab.keys()) == 1:
            for key1 in ab.keys():
                for auslass in ab[key1]:
                    if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 / len(ab[key1]),1)
                    if task.nacht:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 / len(ab[key1]),1)
                    if task.max:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 / len(ab[key1]),1)
                    if task.tnacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 / len(ab[key1]),1)
        elif len(ab.keys()) > 1:
            sum_luft = 0
            for key in ab.keys():
                for auslass in ab[key]:
                    if auslass.Luftmengenmin > 0:sum_luft += auslass.Luftmengenmin
                    else:sum_luft += 0.01
            for key in ab.keys():
                for auslass in ab[key]:
                    if auslass.Luftmengenmin > 0:
                        if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                    else:
                        if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 * 0.01 / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 * 0.01 / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 * 0.01 / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 * 0.01 / sum_luft,1)
        
        for vsr in mep.list_vsr:
            if vsr.art in ['RZU','RAB','LAB','24h','RUM']:
                vsr.Luftmengenermitteln_new()
        self.AederungUebernehmen(mep)
    
    def AederungUebernehmen(self,mep):
        mep.werte_schreiben()
        for auslass in mep.list_auslass:auslass.wert_schreiben()
        for vsr in mep.list_vsr:vsr.wert_schreiben()
    
    def AederungUebernehmenMEP(self,uiapp):
        self.name = 'Änderung übernehmen ' + self.GUI.mepraum.Raumnr
        task = ABFRAGE('Luftmenge akt. Raum übernehmen?',True,70,False)
        task.ShowDialog()
        if task.result == False:
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        self.AederungUebernehmen(self.GUI.mepraum)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
    
    def AederungUebernehmenProjekt(self,uiapp):
        self.name = 'Änderung übernehmen Projekt' 
        task = ABFRAGE('Alle Änderung übernehmen?',False,100,False)
        task.ShowDialog()
        if task.result == False:
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.mepraum_liste.values():  
            self.AederungUebernehmen(mep)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()