# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from Autodesk.Revit.UI.Selection import ISelectionFilter
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
import clr
import math
from IGF_lib import get_value
from System.Collections.Generic import List
doc = __revit__.ActiveUIDocument.Document

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

class Filter(ISelectionFilter):
    def AllowElement(self,elem):
        try:
            if elem.Category.Name == 'MEP-Räume':
                return True
            else:
                return False
        except:return False
    def AllowReference(self,ref,p):
        return False

class ReduzierFaktor(TemplateItemBase):
    def __init__(self,nummer,faktor = 1):
        TemplateItemBase.__init__(self)
        self._nummer = nummer
        self._faktor = faktor
    @property
    def nummer(self):
        return self._nummer
    @nummer.setter
    def nummer(self,value):
        if value != self._nummer:
            self._nummer = value
            self.RaisePropertyChanged('nummer')
    @property
    def faktor(self):
        return self._faktor
    @faktor.setter
    def faktor(self,value):
        if value != self._faktor:
            self._faktor = value
            self.RaisePropertyChanged('faktor')
    
class Schacht(TemplateItemBase):
    def __init__(self,elem,schacht):
        TemplateItemBase.__init__(self)
        self.elem = elem
        self._checked = False
        self.schacht = schacht
        self.reduzierfaktor = ObservableCollection[ReduzierFaktor]()
        self.value = self.elem.LookupParameter('IGF_RLT_Schacht-LuftmengenReduzierung').AsString()
        self.daten = {}
        self.get_detail_daten()

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    
    def get_wert(self):
        wert = ''
        for el in self.reduzierfaktor:
            wert += 'Anl_' + str(el.nummer) + '=' + str(el.faktor) + ', '
        if wert:
            wert = wert[:-2]
        self.value = wert
    
    def get_detail_daten(self):
        self.daten = {}
        if not self.value:
            return
        Liste = self.value.split(', ')
        for el in Liste:
            if el:
                if el.find('Anl_') != -1 and el.find('=') != -1:
                    try:
                        anlnummer = el[4:el.find('=')]
                        faktor = el[el.find('=')+1:]
                        self.daten[anlnummer] = faktor
                    except:
                        pass
    
    def get_items(self):
        for el in self.reduzierfaktor:
            if el.nummer in self.daten.keys():
                el.faktor = self.daten[el.nummer]
    
   
    
    def wert_schreiben(self):
        self.elem.LookupParameter('IGF_RLT_Schacht-LuftmengenReduzierung').Set(self.value)

def get_AnlagenInSchacht(doc,schacht):
    outlin = DB.Outline(schacht.get_BoundingBox(None).Min,schacht.get_BoundingBox(None).Max)
    fil = DB.BoundingBoxIntersectsFilter(outlin,2)
    coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().WherePasses(fil).ContainedInDesignOption(DB.ElementId(-1)).ToElementIds()
    Liste = []
    for elid in coll:
        system = doc.GetElement(elid).LookupParameter('Systemtyp').AsElementId()
        if system.IntegerValue != -1:
            param = doc.GetElement(system).LookupParameter('IGF_X_AnlagenNr')
            if param:
                anlnr = param.AsValueString()
                if anlnr not in Liste:Liste.append(anlnr)
    Items = ObservableCollection[ReduzierFaktor]()
    for anlnr in sorted(Liste):
        Items.Add(ReduzierFaktor(anlnr))
    outlin.Dispose()
    fil.Dispose()
    del(Liste)
    del(coll)
    return Items

class MEPRaum(object):
    def __init__(self,elem):
        self.elem = elem
        self.nummer = self.elem.Number
        self.Name = self.elem.LookupParameter('Name').AsString()
        self.IsSchacht = self.elem.LookupParameter('TGA_RLT_InstallationsSchacht').AsInteger()
        if self.IsSchacht == 0:
            self.schacht = False
        else:
            self.schacht = True
        self.Schachtname = self.elem.LookupParameter('TGA_RLT_InstallationsSchachtName').AsString()

    def werteschreiben(self):
        if self.schacht:
            self.elem.LookupParameter('TGA_RLT_InstallationsSchacht').Set(1)
        else:
            self.elem.LookupParameter('TGA_RLT_InstallationsSchacht').Set(0)
        self.elem.LookupParameter('TGA_RLT_InstallationsSchachtName').Set(self.Schachtname)

class Laboranschlusss(TemplateItemBase):
    def __init__(self,name,art):
        TemplateItemBase.__init__(self)
        self.art = art
        self.Name = name
        self.typname = ''
        if self.art == '24h':
            self.typname = self.Name + ' (24h)'
        else:
            self.typname = self.Name
        self.min = 0
        self.max = 0
        self.Druckverluft = '0'
        self.get_Luftmenge()
        self._Anzahl = ''
        self._ganzahl = 0
        self.Paramwert = ''

        if self.art == '24h':
            self.parametername = 'IGF_RLT_Laboranschluss_24h_'+self.Name
        else:
            self.parametername = 'IGF_RLT_Laboranschluss_LAB_'+self.Name

    @property
    def Anzahl(self):
        return self._Anzahl
    @Anzahl.setter
    def Anzahl(self,value):
        if value != self._Anzahl:
            self._Anzahl = value
            self.RaisePropertyChanged('Anzahl')

    @property
    def ganzahl(self):
        return self._ganzahl
    @ganzahl.setter
    def ganzahl(self,value):
        if value != self._ganzahl:
            self._ganzahl = value
            self.RaisePropertyChanged('ganzahl')

    def get_Luftmenge(self):
        if self.Name.find('Pa') != -1:
            stelle1 = self.Name.find('Pa')
            stelle = self.Name.find('Pa')
            while(self.Name[stelle] != '_'):
                stelle -=1
                if stelle == 0:
                    self.Druckverluft = '0'
                    break
            try:self.Druckverluft = int(self.Name[stelle+1:stelle1])
            except:self.Druckverluft = '0'

        if self.art == '24h':
            self.min = int(self.Name[self.Name.find('_kon')+4:])
            self.max  =self.min
        else:
            self.max = int(self.Name[self.Name.find('_max')+4:])
            self.min = int(self.Name[self.Name.find('_min')+4:self.Name.find('_max')])

    def get_ganzahl(self):
        summe = 0
        try:
            if self.Anzahl.find(',') == -1:
                Liste = [self.Anzahl]
            else:
                Liste = self.Anzahl.split(',')
            for i in Liste:
                if not i:
                    continue
                while(i[0] == ' '):
                    i = i[1:]
                    if not i:break
                if not i:
                    continue
                anzahl = int(i[:i.find('xZeile')])
                summe += anzahl
            self.ganzahl = summe
        except:pass

    def get_Paramwert(self):
        if self.ganzahl == 0:
            self.Paramwert = ''
            return
        try:
            self.Paramwert = str(self.ganzahl) + '_(' + self.Anzahl + ')'
        except:
            pass

class ListeExternalEvent(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.Name = ''
        self.ExcuseApp = ''

    def Execute(self,uiapp):
        if self.ExcuseApp:
            try:
                self.ExcuseApp(uiapp)
            except Exception as e:
                TaskDialog.Show('Fehler',e.ToString())

    def GetName(self):
        return self.Name

    def wert_Schreiben_sp(self,elem):
        """Nur für VolumenstromProNummer, VolumenstromProEinheit"""

        para = elem.LookupParameter('TGA_RLT_VolumenstromProNummer')
        if para:
            try:
                para.SetValueString(self.GUI.berechnung_nach[self.GUI.bezugsname.SelectedItem.ToString()])
            except:
                para.SetValueString('9')

        para1 = elem.LookupParameter('TGA_RLT_VolumenstromProEinheit')
        if para1:
            try:
                para1.Set(self.GUI.einheit_liste[self.GUI.bezugsname.SelectedItem.ToString()])
            except:para1.Set('')

    def wert_schreiben0(self,elem,param,wert):
        """Parametertype String, Wert Combobox"""
        para = elem.LookupParameter(param)
        if wert.SelectedIndex != -1:wert_s = wert.SelectedItem.ToString()
        if para:
            try:
                para.Set(wert_s)
            except:para.Set('')

    def wert_schreiben1(self,elem,param,wert):
        """Parametertype Double/Int, Wert Textbox"""
        para = elem.LookupParameter(param)
        if para:
            try:
                para.Set(int(wert.Text))
            except:para.Set(0)

    def wert_schreiben2(self,elem,param,wert):
        """Parametertype Double, Wert Textbox"""
        para = elem.LookupParameter(param)
        if para:
            try:
                para.SetValueString(wert.Text.replace(',', '.'))
            except:para.SetValueString('')

    def wert_schreiben3(self,elem,param,wert):
        """Parametertype Int, Wert Checkbox"""
        para = elem.LookupParameter(param)
        if para:
            try:
                if wert.IsChecked:
                    para.Set(1)
                else:
                    para.Set(0)
            except:para.Set(0)

    def wert_schreiben4(self,elem,param,wert):
        """Parametertype String, wert String"""
        para = elem.LookupParameter(param)
        if para:
            try:para.Set(wert)

            except:para.Set('')

    def SelectRaum(self,uiapp):
        self.Name = 'MEP-Räume auswählen'
        uidoc = uiapp.ActiveUIDocument
        self.GUI.schachtanlagentabitem.IsEnabled = False
        self.GUI.Raumlufttabitem.IsEnabled = False
        self.GUI.labortabitem.IsEnabled = False
        self.GUI.Schachttabitem.IsEnabled = False
        try:
            filters = Filter()
            elems = uidoc.Selection.PickElementsByRectangle(filters,'MEP-Räume auswählen')
            if elems.Count == 0:
                TaskDialog.Show('Info.','Kein MEP-Raum ausgewählt')
                self.GUI.schachtanlagentabitem.IsEnabled = True
                self.GUI.Raumlufttabitem.IsEnabled = True
                self.GUI.labortabitem.IsEnabled = True
                self.GUI.Schachttabitem.IsEnabled = True
                return
            elemids = uidoc.Selection.GetElementIds()
            elemids.Clear()
            for el in elems:
                elemids.Add(el.Id)
            uidoc.Selection.SetElementIds(elemids)
            self.GUI.schachtanlagentabitem.IsEnabled = True
            self.GUI.Raumlufttabitem.IsEnabled = True
            self.GUI.labortabitem.IsEnabled = True
            self.GUI.Schachttabitem.IsEnabled = True
            return
        except:
            self.GUI.schachtanlagentabitem.IsEnabled = True
            self.GUI.Raumlufttabitem.IsEnabled = True
            self.GUI.labortabitem.IsEnabled = True
            self.GUI.Schachttabitem.IsEnabled = True
    
    def SelectRaum1(self,uiapp):
        def get_parameter(elem,param):
            para = elem.LookupParameter(param)
            if para:
                return para
            else:
                print('Parameter {} nicht vorhanden'.format(param))
        
        self.Name = 'MEP-Räume auswählen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = [doc.GetElement(e) for e in uidoc.Selection.GetElementIds() if doc.GetElement(e).Category.Id.ToString() == '-2003600']
        if len(elems) == 0:
            TaskDialog.Show('Info.','Kein MEP Raum ausgewählt!')
            return
        el = elems[0]
        try:
            self.GUI.bezugsname.SelectedItem = get_parameter(el,'TGA_RLT_VolumenstromProName').AsString()
        except:
            try:
                nummer = get_parameter(el,'TGA_RLT_VolumenstromProNummer').AsValueString()
                if nummer in self.GUI.berechnung_nach.values():
                    for name in self.GUI.berechnung_nach.keys():
                        if nummer == self.GUI.berechnung_nach[name]:
                            self.GUI.bezugsname.SelectedItem = name
                            break
                self.GUI.bezugsname.SelectedIndex = -1
            except:
                self.GUI.bezugsname.SelectedIndex = -1
        try:
            self.GUI.faktor.Text = get_parameter(el,'TGA_RLT_VolumenstromProFaktor').AsValueString()
        except:self.GUI.faktor.Text = ''
        try:
            self.GUI.druckeingabe.Text = str(get_value(get_parameter(el,'IGF_RLT_RaumDruckstufeEingabe')))
        except:self.GUI.druckeingabe.Text = '0'
        try:
            self.GUI.reduRaumeingabe.Text = str(get_value(get_parameter(el,'IGF_RLT_Raum-ReduziertFaktor')))
        except:self.GUI.reduRaumeingabe.Text = '1'
        try:
            self.GUI.isnachtbetrieb.IsChecked = (get_parameter(el,'IGF_RLT_Nachtbetrieb').AsInteger() == 1)
        except:self.GUI.isnachtbetrieb.IsChecked = False
        try:
            self.GUI.LW_nacht.Text = str(get_value(get_parameter(el,'IGF_RLT_NachtbetriebLW')))
        except:self.GUI.LW_nacht.Text = '0'
        try:
            self.GUI.von_nacht.Text = str(get_value(get_parameter(el,'IGF_RLT_NachtbetriebVon')))
        except:self.GUI.von_nacht.Text = '0'
        try:
            self.GUI.bis_nacht.Text = str(get_value(get_parameter(el,'IGF_RLT_NachtbetriebBis')))
        except:self.GUI.bis_nacht.Text = '0'#

        try:
            self.GUI.istiefenachtbetrieb.IsChecked = (get_parameter(el,'IGF_RLT_TieferNachtbetrieb').AsInteger() == 1)
        except:self.GUI.istiefenachtbetrieb.IsChecked = False
        try:
            self.GUI.LW_tnacht.Text = str(get_value(get_parameter(el,'IGF_RLT_TieferNachtbetriebLW')))
        except:self.GUI.LW_tnacht.Text = '0'
        try:
            self.GUI.von_tnacht.Text = str(get_value(get_parameter(el,'IGF_RLT_TieferNachtbetriebVon')))
        except:self.GUI.von_tnacht.Text = '0'
        try:
            self.GUI.bis_tnacht.Text = str(get_value(get_parameter(el,'IGF_RLT_TieferNachtbetriebBis')))
        except:self.GUI.bis_tnacht.Text = '0'

        uidoc.Dispose()
        doc.Dispose()
    
    def get_AnlagenInSchacht(self,doc,schacht):
        return get_AnlagenInSchacht(doc,schacht)
    
    def SelectRaum2(self,uiapp):
        def get_parameter(elem,param):
            para = elem.LookupParameter(param)
            if para:
                return para
            else:
                print('Parameter {} nicht vorhanden'.format(param))
        
        self.Name = 'MEP-Räume auswählen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = [doc.GetElement(e) for e in uidoc.Selection.GetElementIds() if doc.GetElement(e).Category.Id.ToString() == '-2003600']
        if len(elems) == 0:
            TaskDialog.Show('Info.','Kein MEP Raum ausgewählt!')
            return
        el = elems[0]

        try:
            self.GUI.SchachtRZU.SelectedItem = get_parameter(el,'TGA_RLT_SchachtZuluft').AsString()
        except:self.GUI.SchachtRZU.SelectedIndex = -1
        try:
            self.GUI.SchachtRAB.SelectedItem = get_parameter(el,'TGA_RLT_SchachtAbluft').AsString()
        except:self.GUI.SchachtRAB.SelectedIndex = -1
        try:
            self.GUI.SchachtTZU.SelectedItem = get_parameter(el,'IGF_RLT_Schacht_TZU').AsString()
        except:self.GUI.SchachtTZU.SelectedIndex = -1
        try:
            self.GUI.SchachtTAB.SelectedItem = get_parameter(el,'IGF_RLT_Schacht_TAB').AsString()
        except:self.GUI.SchachtTAB.SelectedIndex = -1
        try:
            self.GUI.Schacht24h.SelectedItem = get_parameter(el,'TGA_RLT_Schacht24hAbluft').AsString()
        except:self.GUI.Schacht24h.SelectedIndex = -1
        try:
            self.GUI.SchachtLAB.SelectedItem = get_parameter(el,'IGF_RLT_Schacht_LAB').AsString()
        except:self.GUI.SchachtLAB.SelectedIndex = -1

        try:
            self.GUI.AnlagenRZU.Text = str(get_value(get_parameter(el,'IGF_RLT_AnlagenNr_RZU')))
        except:self.GUI.AnlagenRZU.Text = '0'
        try:
            self.GUI.AnlagenRAB.Text = str(get_value(get_parameter(el,'IGF_RLT_AnlagenNr_RAB')))
        except:self.GUI.AnlagenRAB.Text = '0'
        try:
            self.GUI.AnlagenTZU.Text = str(get_value(get_parameter(el,'IGF_RLT_AnlagenNr_TZU')))
        except:self.GUI.AnlagenTZU.Text = '0'

        try:
            self.GUI.AnlagenTAB.Text = str(get_value(get_parameter(el,'IGF_RLT_AnlagenNr_TAB')))
        except:self.GUI.AnlagenTAB.Text = '0'
        try:
            self.GUI.Anlagen24h.Text = str(get_value(get_parameter(el,'IGF_RLT_AnlagenNr_24h')))
        except:self.GUI.Anlagen24h.Text = '0'
        try:
            self.GUI.AnlagenLAB.Text = str(get_value(get_parameter(el,'IGF_RLT_AnlagenNr_LAB')))
        except:self.GUI.AnlagenLAB.Text = '0'

        uidoc.Dispose()
        doc.Dispose()


    def SchreibenSchachtAnlagen(self,uiapp):
        self.Name = 'Schacht&Anlagen schreiebn'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        for el in uidoc.Selection.GetElementIds():
            elem = doc.GetElement(el)
            if elem.Category.Id.ToString() == '-2003600':
                elems.append(elem)
        if len(elems) == 0:
            TaskDialog.Show('Info.','Bitte MEP-Raum auswählen!')
            return

        t = DB.Transaction(doc,'Schacht&Anlagen schreiebn')
        t.Start()
        for el in elems:
            if self.GUI.cbSzu.IsChecked:self.wert_schreiben0(el,'TGA_RLT_SchachtZuluft',self.GUI.SchachtRZU)
            if self.GUI.cbSab.IsChecked:self.wert_schreiben0(el,'TGA_RLT_SchachtAbluft',self.GUI.SchachtRAB)
            if self.GUI.cbStzu.IsChecked:self.wert_schreiben0(el,'IGF_RLT_Schacht_TZU',self.GUI.SchachtTZU)
            if self.GUI.cbStab.IsChecked:self.wert_schreiben0(el,'IGF_RLT_Schacht_TAB',self.GUI.SchachtTAB)
            if self.GUI.cbS24h.IsChecked:self.wert_schreiben0(el,'TGA_RLT_Schacht24hAbluft',self.GUI.Schacht24h)
            if self.GUI.cbSlab.IsChecked:self.wert_schreiben0(el,'IGF_RLT_Schacht_LAB',self.GUI.SchachtLAB)
            if self.GUI.cbAzu.IsChecked:self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_RZU',self.GUI.AnlagenRZU)
            if self.GUI.cbAab.IsChecked:self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_RAB',self.GUI.AnlagenRAB)
            if self.GUI.cbAtzu.IsChecked:self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_TZU',self.GUI.AnlagenTZU)
            if self.GUI.cbAtab.IsChecked:self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_TAB',self.GUI.AnlagenTAB)
            if self.GUI.cbA24h.IsChecked:self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_24h',self.GUI.Anlagen24h)
            if self.GUI.cbAlab.IsChecked:self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_LAB',self.GUI.AnlagenLAB)

        t.Commit()

    def SchreibenRaumluft(self,uiapp):
        self.Name = 'Raumluftberechnung'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        for el in uidoc.Selection.GetElementIds():
            elem = doc.GetElement(el)
            if elem.Category.Id.ToString() == '-2003600':
                elems.append(elem)
        if len(elems) == 0:
            TaskDialog.Show('Info.','Bitte MEP-Raum auswählen!')
            return


        t = DB.Transaction(doc,'Raumluftberechnung schreiebn')
        t.Start()
        for el in elems:
            self.wert_Schreiben_sp(el)
            if self.GUI.cbbn.IsChecked:
                self.wert_schreiben0(el,'TGA_RLT_VolumenstromProName',self.GUI.bezugsname)
                if self.GUI.bezugsname.SelectedItem.ToString() in self.GUI.berechnung_nach.keys():
                    el.LookupParameter('TGA_RLT_VolumenstromProNummer').SetValueString(str(self.GUI.berechnung_nach[self.GUI.bezugsname.SelectedItem.ToString()]))
            if self.GUI.cbfk.IsChecked:self.wert_schreiben2(el,'TGA_RLT_VolumenstromProFaktor',self.GUI.faktor)
            if self.GUI.cbds.IsChecked:self.wert_schreiben2(el,'IGF_RLT_RaumDruckstufeEingabe',self.GUI.druckeingabe)
            if self.GUI.cbrr.IsChecked:self.wert_schreiben2(el,'IGF_RLT_Raum-ReduziertFaktor',self.GUI.reduRaumeingabe)
            if self.GUI.cbsr.IsChecked:self.wert_schreiben2(el,'IGF_RLT_Schacht-ReduziertFaktor',self.GUI.reduschachteingabe)

            if self.GUI.cbnb.IsChecked:
                self.wert_schreiben3(el,'IGF_RLT_Nachtbetrieb',self.GUI.isnachtbetrieb)
                self.wert_schreiben2(el,'IGF_RLT_NachtbetriebLW',self.GUI.LW_nacht)
                self.wert_schreiben2(el,'IGF_RLT_NachtbetriebVon',self.GUI.von_nacht)
                self.wert_schreiben2(el,'IGF_RLT_NachtbetriebBis',self.GUI.bis_nacht)

            if self.GUI.cbtnb.IsChecked:
                self.wert_schreiben3(el,'IGF_RLT_TieferNachtbetrieb',self.GUI.istiefenachtbetrieb)
                self.wert_schreiben2(el,'IGF_RLT_TieferNachtbetriebLW',self.GUI.LW_tnacht)
                self.wert_schreiben2(el,'IGF_RLT_TieferNachtbetriebVon',self.GUI.von_tnacht)
                self.wert_schreiben2(el,'IGF_RLT_TieferNachtbetriebBis',self.GUI.bis_tnacht)

        t.Commit()

    def SchreibenLabor(self,uiapp):
        self.Name = 'Laborabluft'
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

        t = DB.Transaction(doc,'Laborabluft schreiebn')
        t.Start()

        self.wert_schreiben2(el,'IGF_RLT_AbluftminSummeLabor',self.GUI.labmineingabe)
        self.wert_schreiben2(el,'IGF_RLT_AbluftmaxSummeLabor',self.GUI.labmaxeingabe)
        self.wert_schreiben2(el,'IGF_RLT_AbluftSumme24h',self.GUI.ab24heingabe)
        t.Commit()
        t.Dispose()

    def Laboranschlussanpassen(self,uiapp):
        self.Name = 'Laboranschluss'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        t = DB.Transaction(doc,'Laboranschluss anpassen')
        t.Start()

        for el in self.GUI.IS_Laboranschluss:
            try:
                if el.typname in self.GUI.LaboranschluesseDict.keys():
                    typ = doc.GetElement(self.GUI.LaboranschluesseDict[el.typname])
                    typ.LookupParameter('TGA_RLT_AuslassVolumenstromMin').SetValueString(str(el.min))
                    typ.LookupParameter('TGA_RLT_AuslassVolumenstromMax').SetValueString(str(el.max))
                    typ.LookupParameter('TGA_RLT_AuslassDruckverlust').SetValueString(str(el.Druckverluft))

                else:
                    typ = doc.GetElement(self.GUI.LaboranschluesseDict.values()[0]).Duplicate(el.typname)
                    doc.Regenerate()
                    typ.LookupParameter('TGA_RLT_AuslassVolumenstromMin').SetValueString(str(el.min))
                    typ.LookupParameter('TGA_RLT_AuslassVolumenstromMax').SetValueString(str(el.max))
                    typ.LookupParameter('TGA_RLT_AuslassDruckverlust').SetValueString(str(el.Druckverluft))
            except Exception as e:
                print(e)
        t.Commit()
        t.Dispose()

    def get_Laboranschluss(self,uiapp):
        self.Name = 'Laboranschluss'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        self.GUI.IsEnabled = False
        Familie = self.GUI.labanschluss_Dict[self.GUI.labanschluss.SelectedItem.ToString()]
        Symbols = Familie.GetFamilySymbolIds()
        self.GUI.LaboranschluesseDict = {}
        for el in Symbols:
            name = doc.GetElement(el).LookupParameter('Typname').AsString()
            self.GUI.LaboranschluesseDict[name] = el
        self.GUI.IsEnabled = True

    def UpdateRaumInfo(self,uiapp):
        self.Name = 'GUI Aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        self.GUI.ItemtempMEP.Clear()
        elems = []
        for elid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(elid)
            if el.Category.Id.IntegerValue == -2003600:
                elems.append(el)

        for el in elems:
            self.GUI.ItemtempMEP.Add(MEPRaum(el))
        self.GUI.LVSchacht0.Items.Refresh()

    def UpdateRaumInfo_Labor(self,uiapp):
        self.Name = 'GUI Aktualisieren'
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
        name = el.Number + ' - ' + el.LookupParameter('Name').AsString()
        if name in self.GUI.Dict_MEP.keys():
            self.GUI.Raumdatensource.SelectedItem = name
        self._UpdateRaumInfo_Labor(el)
    
    def MEP_Labor(self,uiapp):
        self.Name = 'GUI Aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        if self.GUI.Raumdatensource.SelectedIndex == -1:
            TaskDialog.Show('Info.','Bitte ein MEP-Raum auswählen!')
            return
        elid = self.GUI.Dict_MEP[self.GUI.Raumdatensource.SelectedItem.ToString()]
        uidoc.Selection.SetElementIds(List[DB.ElementId]([elid]))
        el = doc.GetElement(elid)
        self._UpdateRaumInfo_Labor(el)

        uidoc.Dispose()
        doc.Dispose()
    
    def _UpdateRaumInfo_Labor(self,mep):
        for Param in self.GUI.IS_Laboranschluss:
            value = mep.LookupParameter(Param.parametername).AsString()
            if value:
                try:
                    detail = value[value.find('(')+1:value.find(')')]
                except:
                    detail = ''
                Param.Anzahl = detail
            else:
                Param.Anzahl = ''
            Param.get_ganzahl()
        self.GUI.Getluftmenge()

    def SchreibenLaborDetail(self,uiapp):
        self.Name = 'Laborabluft'
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

        t = DB.Transaction(doc,'Laborabluft schreiebn')
        t.Start()
        self.wert_schreiben2(el,'IGF_RLT_AbluftminSummeLabor',self.GUI.labminresult)
        self.wert_schreiben2(el,'IGF_RLT_AbluftmaxSummeLabor',self.GUI.labmaxresult)
        self.wert_schreiben2(el,'IGF_RLT_AbluftSumme24h',self.GUI.ab24hresult)
        for param in self.GUI.IS_Laboranschluss:
            param.get_Paramwert()
            self.wert_schreiben4(el,param.parametername,param.Paramwert)
        t.Commit()
        t.Dispose()

    def SchreibenLaborDetail_Anschluss(self,uiapp):
        self.Name = 'Laboranschluss'
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

        if self.GUI.labanschluss.SelectedIndex == -1:
            TaskDialog.Show('Info.','Bitte Laboranschluss-Familie auswählen!')
            return

        if self.GUI.ebene.SelectedIndex == -1:
            TaskDialog.Show('Info.','Bitte Ebene auswählen!')
            return

        param_equality = DB.FilterStringEquals()
        fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        fam_prov=DB.ParameterValueProvider(fam_id)
        fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,self.GUI.labanschluss.SelectedItem.ToString(),True)
        fam_filter = DB.ElementParameterFilter(fam_value_rule)
        colls = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().WherePasses(fam_filter).ToElements()
        colls_ = {}

        for elem in colls:
            space = elem.Space[doc.GetElement(elem.CreatedPhaseId)]
            if space:
                if space.Id.ToString() == el.Id.ToString():
                    if elem.Name not in colls_.keys():
                        colls_[elem.Name] = []
                    colls_[elem.Name].append(elem)

        t = DB.Transaction(doc,'Laboranschluss')
        t.Start()

        Liste = colls_.keys()[:]
        Text = ''

        level = self.GUI.Levels[self.GUI.ebene.SelectedItem.ToString()]
        height = level.get_Parameter('DB.BuiltInParameter.LEVEL_ELEV').AsDouble()
        try:
            height0 = float(self.GUI.heightvonebene.Text.replace(',','.'))/304.8
        except:
            height0 = 0

        p0 = el.Location.Point
        z = height + height0

        Liste_Neu = []

        for param in self.GUI.Anschlussdetail.Items:
            anzahl = param.ganzahl
            name = param.typname
            if name in Liste:
                Liste.remove(name)
                anzahl_vorhanden = len(colls_[name])
            else:
                anzahl_vorhanden = 0

            if anzahl > anzahl_vorhanden:
                for n in range(anzahl-anzahl_vorhanden):
                    Liste_Neu.append(doc.GetElement(self.GUI.LaboranschluesseDict[name]))
            elif anzahl < anzahl_vorhanden:
                Text += '{} {} gebraucht, aber {} gezeichnet'.format(anzahl,name,anzahl_vorhanden)

        ganzahl = len(Liste_Neu)
        if ganzahl > 0:
            col = math.ceil(math.sqrt(ganzahl))
            row = col
            while (col*row >= ganzahl):
                row -= 1
            row += 1
            n_Liste_Neu = 0
            if row % 2 != 0 and col % 2 != 0:
                for row0 in range(-row/2,row/2+1):
                    for col0 in range(-col/2,col/2+1):
                        p = DB.XYZ(p0.X+300/304.8*row0,p0.Y+300/304.8*col0,z)
                        try:
                            doc.Create.NewFamilyInstance(p,Liste_Neu[n_Liste_Neu],level,DB.Structure.StructuralType.NonStructural)
                        except Exception as e:
                            print(e)
                        n_Liste_Neu+=1
                        if n_Liste_Neu+1 > ganzahl:
                            break
            elif row % 2 == 0 and col % 2 != 0:
                for row0 in range(-row/2,row/2):
                    for col0 in range(-col/2,col/2+1):
                        p = DB.XYZ(p0.X+300/304.8*(row0+0.5),p0.Y+300/304.8*col0,z)
                        try:
                            doc.Create.NewFamilyInstance(p,Liste_Neu[n_Liste_Neu],level,DB.Structure.StructuralType.NonStructural)
                        except Exception as e:
                            print(e)
                        n_Liste_Neu+=1
                        if n_Liste_Neu+1 > ganzahl:
                            break
            elif row % 2 == 0 and col % 2 == 0:
                for row0 in range(-row/2,row/2):
                    for col0 in range(-col/2,col/2):
                        p = DB.XYZ(p0.X+300/304.8*(row0+0.5),p0.Y+300/304.8*(col0+0.5),z)
                        try:
                            doc.Create.NewFamilyInstance(p,Liste_Neu[n_Liste_Neu],level,DB.Structure.StructuralType.NonStructural)
                        except Exception as e:
                            print(e)
                        n_Liste_Neu+=1
                        if n_Liste_Neu+1 > ganzahl:
                            break
            elif row % 2 != 0 and col % 2 == 0:
                for row0 in range(-row/2,row/2+1):
                    for col0 in range(-col/2,col/2):
                        p = DB.XYZ(p0.X+300/304.8*row0,p0.Y+300/304.8*(col0+0.5),z)
                        try:
                            doc.Create.NewFamilyInstance(p,Liste_Neu[n_Liste_Neu],level,DB.Structure.StructuralType.NonStructural)
                        except Exception as e:
                            print(e)
                        n_Liste_Neu+=1
                        if n_Liste_Neu+1 > ganzahl:
                            break

        t.Commit()
        t.Dispose()
        for eleme in Liste:
            Text += '0 {} gebraucht, aber {} gezeichnet'.format(eleme,len(colls_[eleme]))
        if Text:
            Task = TaskDialog('Fehler')
            Task.MainInstruction = 'Folgende Laboranschlüsse:'
            Task.MainContent = Text
            Task.Show()

    def SchreibenSchacht0(self,uiapp):
        self.Name = 'Schacht aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _Dict = {}
        for schacht in self.GUI.IS_Schacht_Neu:
             _Dict[schacht.schacht] = schacht
        t = DB.Transaction(doc,'Schacht aktualisieren')
        t.Start()
        for el in self.GUI.LVSchacht0.Items:
            el.werteschreiben()
            if el.schacht:
                if el.elem.Id.ToString() not in self.GUI.schachtelemids:
                    self.GUI.schachtelemids.append(el.elem.Id.ToString())
                    self.GUI.IS_Schacht.Add(el)
                else:
                    for elem in self.GUI.LVSchacht.Items:
                        if el.elem.Id.ToString() == elem.elem.Id.ToString():
                            elem.schacht = True
                            elem.Schachtname = el.Schachtname
                if el.Schachtname not in self.GUI.LISTE_SCHACHT:
                    self.GUI.LISTE_SCHACHT.append(el.Schachtname)

                if el.Schachtname not in _Dict.keys():
                    temp_klasse = Schacht(el.elem,el.Schachtname)
                    temp_klasse.reduzierfaktor = self.get_AnlagenInSchacht(doc,el.elem)
                    temp_klasse.get_items()
                    self.GUI.IS_Schacht_Neu.Add(temp_klasse)

            else:
                if el.elem.Id.ToString() in self.GUI.schachtelemids:
                    if el.elem.Id.ToString() == elem.elem.Id.ToString():
                        elem.schacht = False
                        elem.Schachtname = el.Schachtname
                if el.Schachtname in self.GUI.LISTE_SCHACHT:
                    self.GUI.LISTE_SCHACHT.remove(el.Schachtname)
        t.Commit()
        self.GUI.LVSchacht.Item.Refresh()
        self.GUI.set_up_is()

    def SchreibenSchacht1(self,uiapp):
        self.Name = 'Schacht aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        t = DB.Transaction(doc,'Schacht')
        t.Start()
        _Dict = {}
        for schacht in self.GUI.IS_Schacht_Neu:
             _Dict[schacht.schacht] = schacht
        for el in self.GUI.LVSchacht.Items:
            el.werteschreiben()
            if not el.schacht:
                if el.Schachtname in self.GUI.LISTE_SCHACHT:
                    self.GUI.LISTE_SCHACHT.remove(el.Schachtname)
                if el.Schachtname in _Dict.keys():
                    self.GUI.IS_Schacht_Neu.Remove(_Dict[el.Schachtname])
                    del _Dict[schacht.schacht] 
          
        t.Commit()
        self.GUI.set_up_is()
    
    def schachtfaktorschreiben(self,uiapp):
        self.Name = 'Schacht aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        try:
            t = DB.Transaction(doc,'Schacht')
            t.Start()
            for schacht in self.GUI.IS_Schacht_Neu:
                print( schacht.checked)
                if schacht.checked:
                    schacht.get_wert()
                    schacht.wert_schreiben()
            
            t.Commit()
        except Exception as e:print(e)


LISTE_SCHACHT = []
IS_Schacht = ObservableCollection[MEPRaum]()
IS_Laboranschluss = ObservableCollection[Laboranschlusss]()
labanschluss_Dict = {}
LEVEL_Dict = {}
Dict_MEPRaum = {}
IS_Schacht_Neu = ObservableCollection[Schacht]()
def get_AlleInfos():
    levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
    for el in levels:
        name = el.Name
        LEVEL_Dict[name] = el

    Liste = []
    Dict = {'Lab':[],'24h':[]}
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        schacht = el.LookupParameter('TGA_RLT_InstallationsSchacht').AsInteger()
        if schacht == 1:
            Liste.append(el)
        else:
            name = el.Number + ' - ' + el.LookupParameter('Name').AsString()
            Dict_MEPRaum[name] = el.Id
    for el in Liste:
        temp = MEPRaum(el)
        IS_Schacht.Add(temp)
        schacht = Schacht(el,temp.Schachtname)
        schacht.reduzierfaktor = get_AnlagenInSchacht(doc,el)
        schacht.get_items()
        IS_Schacht_Neu.Add(schacht)
        if temp.Schachtname not in LISTE_SCHACHT:
            LISTE_SCHACHT.append(temp.Schachtname)

    for el in spaces:
        params = el.Parameters
        for param in params:
            name = param.Definition.Name
            if name.find('IGF_RLT_Laboranschluss_LAB') != -1:
                Dict['Lab'].append(name[27:])
            elif name.find('IGF_RLT_Laboranschluss_24h') != -1:
                Dict['24h'].append(name[27:])
        break
    for el in Dict.keys():
        for param in sorted(Dict[el]):
            IS_Laboranschluss.Add(Laboranschlusss(param,el))

    coll = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family)).ToElements()
    for el in coll:
        if el.FamilyCategory.Name == 'Luftdurchlässe':
            labanschluss_Dict[el.Name] = el

get_AlleInfos()
