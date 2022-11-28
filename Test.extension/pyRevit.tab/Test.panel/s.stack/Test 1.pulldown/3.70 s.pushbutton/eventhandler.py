# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection

doc = __revit__.ActiveUIDocument.Document

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

class Laboranschlusss(object):
    def __init__(self,name,art):
        self.art = art
        self.Name = name	
        self.min = 0
        self.max = 0
        self.get_Luftmenge()
        self.Anzahl = ''
        self.ganzahl = 0
        self.Paramwert = ''
        if self.art == '24h':
            self.parametername = 'IGF_RLT_Laboranschluss_24h_'+self.Name
        else:
            self.parametername = 'IGF_RLT_Laboranschluss_LAB_'+self.Name
    
    def get_Luftmenge(self):
        if self.art == '24h':
            self.min = int(self.Name[self.Name.find('_kon')+4:])
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
                print(e)
    
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
            self.wert_schreiben0(el,'TGA_RLT_SchachtZuluft',self.GUI.SchachtRZU)
            self.wert_schreiben0(el,'TGA_RLT_SchachtAbluft',self.GUI.SchachtRAB)
            self.wert_schreiben0(el,'IGF_RLT_Schacht_TZU',self.GUI.SchachtTZU)
            self.wert_schreiben0(el,'IGF_RLT_Schacht_TAB',self.GUI.SchachtTAB)
            self.wert_schreiben0(el,'TGA_RLT_Schacht24hAbluft',self.GUI.Schacht24h)
            self.wert_schreiben0(el,'IGF_RLT_Schacht_LAB',self.GUI.SchachtLAB)
            self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_RZU',self.GUI.AnlagenRZU)
            self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_RAB',self.GUI.AnlagenRAB)
            self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_TZU',self.GUI.AnlagenTZU)
            self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_TAB',self.GUI.AnlagenTAB)
            self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_24h',self.GUI.Anlagen24h)
            self.wert_schreiben1(el,'IGF_RLT_AnlagenNr_LAB',self.GUI.AnlagenLAB)
        
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
            self.wert_schreiben0(el,'TGA_RLT_VolumenstromProName',self.GUI.bezugsname)
            self.wert_schreiben2(el,'TGA_RLT_VolumenstromProFaktor',self.GUI.faktor)
            self.wert_schreiben2(el,'IGF_RLT_RaumDruckstufeEingabe',self.GUI.druckeingabe)

            self.wert_schreiben3(el,'IGF_RLT_Nachtbetrieb',self.GUI.isnachtbetrieb)
            self.wert_schreiben2(el,'IGF_RLT_NachtbetriebLW',self.GUI.LW_nacht)
            self.wert_schreiben2(el,'IGF_RLT_NachtbetriebVon',self.GUI.von_nacht)
            self.wert_schreiben2(el,'IGF_RLT_NachtbetriebBis',self.GUI.bis_nacht)

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

    def SchreibenSchacht0(self,uiapp):
        self.Name = 'Schacht aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
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
        for el in self.GUI.LVSchacht.Items:
            el.werteschreiben()
            if not el.schacht:
                if el.Schachtname in self.GUI.LISTE_SCHACHT:
                    self.GUI.LISTE_SCHACHT.remove(el.Schachtname)
        t.Commit()
        self.GUI.set_up_is()

LISTE_SCHACHT = []
IS_Schacht = ObservableCollection[MEPRaum]()
IS_Laboranschluss = ObservableCollection[MEPRaum]()

def get_AlleInfos():
    Liste = []
    Dict = {'Lab':[],'24h':[]}
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        schacht = el.LookupParameter('TGA_RLT_InstallationsSchacht').AsInteger()
        if schacht == 1:
            Liste.append(el)
    for el in Liste:
        temp = MEPRaum(el)
        IS_Schacht.Add(temp)
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

get_AlleInfos()

