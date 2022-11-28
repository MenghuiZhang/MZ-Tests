# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')

from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from eventhandler import SystemTyp,DICT_BERECHNUNG,DICT_BERECHNUNG_UMGEKEHRT,\
    DICT_TEXT_INDEX,ExternalEvent,KeineEinstellen,Zuruecksetzen,ObservableCollection,\
    DICT_INDEX_TEXT,AltDaten,AltBerechnung,IS_BERECHNUNG,\
    SystemTypforalt,Einstellen,Meldung
    #TaskDialog,TaskDialogCommonButtons,
from time import strftime, localtime,time


__title__ = "Berechnungsmodus"
__doc__ = """

[2022.04.07]
Version: 1.1
"""
__authors__ = "Menghui Zhang"



logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+' - Systemberechnung')

try:
    getlog(__title__)
except:
    pass




SYSTEMDATEN_LUFT = {}
SYSTEMDATEN_ROHR = {}
ITEMSSOURCE_LUFT = ObservableCollection[SystemTyp]()
ITEMSSOURCE_ROHR = ObservableCollection[SystemTyp]()

def get_systemdaten():
    luftsystyp = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType().ToElements()
    rohrsystyp = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType().ToElements()
    for el in luftsystyp:
        elemid = el.GetTypeId()
        typ = doc.GetElement(elemid)
        name = typ.LookupParameter('Typname').AsString()
        Klasse = typ.LookupParameter('Systemklassifizierung').AsString()
        berechnung = typ.CalculationLevel.ToString()
        if name not in SYSTEMDATEN_LUFT.keys():
            SYSTEMDATEN_LUFT[name] = SystemTyp(typ,name,DICT_BERECHNUNG_UMGEKEHRT[berechnung],IS_BERECHNUNG[Klasse],DICT_TEXT_INDEX[DICT_BERECHNUNG_UMGEKEHRT[berechnung]])
    for el in rohrsystyp:
        elemid = el.GetTypeId()
        typ = doc.GetElement(elemid)
        name = typ.LookupParameter('Typname').AsString()
        Klasse = typ.LookupParameter('Systemklassifizierung').AsString()
        berechnung = typ.CalculationLevel.ToString()
        if name not in SYSTEMDATEN_ROHR.keys():
            SYSTEMDATEN_ROHR[name] = SystemTyp(typ,name,DICT_BERECHNUNG_UMGEKEHRT[berechnung],IS_BERECHNUNG[Klasse],DICT_TEXT_INDEX[DICT_BERECHNUNG_UMGEKEHRT[berechnung]])

get_systemdaten()

def get_Itemssource():
    for key in sorted(SYSTEMDATEN_LUFT.keys()):ITEMSSOURCE_LUFT.Add(SYSTEMDATEN_LUFT[key])
    for key in sorted(SYSTEMDATEN_ROHR.keys()):ITEMSSOURCE_ROHR.Add(SYSTEMDATEN_ROHR[key])
        
get_Itemssource()

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self,Liste_luft,Liste_rohr):
        self.classSystemtyp = SystemTyp
        self.classSystemtypforalt = SystemTypforalt
        self.classAltDaten = AltDaten
        self.classAltBerechnung = AltBerechnung
        self.DICT_INDEX_TEXT = DICT_INDEX_TEXT
        self.DICT_TEXT_INDEX = DICT_TEXT_INDEX
        self.classMeldung = Meldung
        self.DICT_BERECHNUNG = DICT_BERECHNUNG
        self.script = script
        self.strftime = strftime
        self.localtime = localtime
        self.time = time
        self.zeit = None
        self.Liste_luft = Liste_luft
        self.Liste_rohr = Liste_rohr
        self.ObservableCollection = ObservableCollection
        self.keineEinstellen = KeineEinstellen()
        self.zuruecksetzen = Zuruecksetzen()
        self.einstellenhand = Einstellen()
        self.einstellenhandEvent = ExternalEvent.Create(self.einstellenhand)
        self.keineEinstellenEvent = ExternalEvent.Create(self.keineEinstellen)
        self.zuruecksetzenEvent = ExternalEvent.Create(self.zuruecksetzen)
        forms.WPFWindow.__init__(self,'window.xaml')
        self.listview.ItemsSource = self.Liste_luft
        self.altdatagrid = self.Liste_luft
        self.tempcoll = self.ObservableCollection[self.classSystemtyp]()
        self.config = config
        self.logger = logger
        self.altendaten = self.ObservableCollection[self.classAltDaten]()
        self.read_config()
        self.datum_lv.ItemsSource = self.altendaten
        self.temp_luftsystemtyp = True

        self.alttime = None
        self.warten = None
        
        
    
    def get_zeit(self):
        return self.strftime("%d.%m.%Y-%H:%M", self.localtime())

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.listview.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.name.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.listview.ItemsSource = self.tempcoll
        self.listview.Items.Refresh()

    def read_config(self):
        try:altendaten1 = self.config.alt1
        except:altendaten1 = None
        try:altendaten2 = self.config.alt2
        except:altendaten2 = None
        try:altendaten3 = self.config.alt3
        except:altendaten3 = None
        try:altendaten4 = self.config.alt4
        except:altendaten4 = None
        try:altendaten5 = self.config.alt5
        except:altendaten5 = None

        try:altendatum1 = self.config.altzeit1
        except:altendatum1 = None
        try:altendatum2 = self.config.altzeit2
        except:altendatum2 = None
        try:altendatum3 = self.config.altzeit3
        except:altendatum3 = None
        try:altendatum4 = self.config.altzeit4
        except:altendatum4 = None
        try:altendatum5 = self.config.altzeit5
        except:altendatum5 = None

        try:
            if altendaten1 and altendatum1:
                altdaten1 = self.ObservableCollection[self.classSystemtypforalt]()
                for el in self.Liste_luft:
                    if el.name in altendaten1.keys():
                        altdaten1.Add(self.classSystemtypforalt(el.elem,el.name,altendaten1[el.name]))
                for el in self.Liste_rohr:
                    if el.name in altendaten1.keys():
                        altdaten1.Add(self.classSystemtypforalt(el.elem,el.name,altendaten1[el.name]))
            self.altendaten.Add(self.classAltDaten(altendatum1,altdaten1))    
        except:pass

        try:
            if altendaten2 and altendatum2:
                altdaten2 = self.ObservableCollection[self.classSystemtypforalt]()
                for el in self.Liste_luft:
                    if el.name in altendaten2.keys():
                        altdaten2.Add(self.classSystemtypforalt(el.elem,el.name,altendaten2[el.name]))
                for el in self.Liste_rohr:
                    if el.name in altendaten2.keys():
                        altdaten2.Add(self.classSystemtypforalt(el.elem,el.name,altendaten2[el.name]))
            self.altendaten.Add(self.classAltDaten(altendatum2,altdaten2))    
        except:pass

        try:
            if altendaten3 and altendatum3:
                altdaten3 = self.ObservableCollection[self.classSystemtypforalt]()
                for el in self.Liste_luft:
                    if el.name in altendaten3.keys():
                        altdaten3.Add(self.classSystemtypforalt(el.elem,el.name,altendaten3[el.name]))
                for el in self.Liste_rohr:
                    if el.name in altendaten3.keys():
                        altdaten3.Add(self.classSystemtypforalt(el.elem,el.name,altendaten3[el.name]))
            self.altendaten.Add(self.classAltDaten(altendatum3,altdaten3))    
        except:pass

        try:
            if altendaten4 and altendatum4:
                altdaten4 = self.ObservableCollection[self.classSystemtypforalt]()
                for el in self.Liste_luft:
                    if el.name in altendaten4.keys():
                        altdaten4.Add(self.classSystemtypforalt(el.elem,el.name,altendaten4[el.name]))
                for el in self.Liste_rohr:
                    if el.name in altendaten4.keys():
                        altdaten4.Add(self.classSystemtypforalt(el.elem,el.name,altendaten4[el.name]))
            self.altendaten.Add(self.classAltDaten(altendatum4,altdaten4))    
        except:pass

        try:
            if altendaten5 and altendatum5:
                altdaten5 = self.ObservableCollection[self.classSystemtypforalt]()
                for el in self.Liste_luft:
                    if el.name in altendaten5.keys():
                        altdaten5.Add(self.classSystemtypforalt(el.elem,el.name,altendaten5[el.name]))
                for el in self.Liste_rohr:
                    if el.name in altendaten5.keys():
                        altdaten5.Add(self.classSystemtypforalt(el.elem,el.name,altendaten5[el.name]))
            self.altendaten.Add(self.classAltDaten(altendatum5,altdaten5))    
        except:pass
    
    def write_config(self):
        aktuell = {}
        try:self.config.alt5 = self.config.alt4
        except:self.config.alt5 = {}
        try:self.config.alt4 = self.config.alt3
        except:self.config.alt4 = {}
        try:self.config.alt3 = self.config.alt2
        except:self.config.alt3 = {}
        try:self.config.alt2 = self.config.alt1
        except:self.config.alt2 = {}
        

        try:self.config.altzeit5 = self.config.altzeit4
        except:self.config.altzeit5 = ''
        try:self.config.altzeit4 = self.config.altzeit3
        except:self.config.altzeit4 = ''
        try:self.config.altzeit3 = self.config.altzeit2
        except:self.config.altzeit3 = ''
        try:self.config.altzeit2 = self.config.altzeit1
        except:self.config.altzeit2 = ''
        try:self.config.altzeit1 = self.zeit
        except:self.config.altzeit1 = ''


        try:
            
            for el in self.Liste_luft:
                aktuell[el.name] = el.berechnung
            for el in self.Liste_rohr:
                aktuell[el.name] = el.berechnung
        except:pass

        try:self.config.alt1 = aktuell
        except:self.config.alt1 = {}
        self.script.save_config()
        
    def speichern(self, sender, args):
        self.zeit = self.get_zeit()
        self.write_config()
        alt = self.ObservableCollection[self.classSystemtypforalt]()
        for el in self.Liste_luft:
            alt.Add(self.classSystemtypforalt(el.elem,el.name,self.DICT_INDEX_TEXT[el.berechnungsindex]))
        for el in self.Liste_rohr:
            alt.Add(self.classSystemtypforalt(el.elem,el.name,self.DICT_INDEX_TEXT[el.berechnungsindex]))

        self.altendaten.Insert(0, self.classAltDaten(self.zeit,alt))
        self.datum_lv.Items.Refresh()

    def luft(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.Liste_luft
        self.altdatagrid = self.Liste_luft


    def rohr(self, sender, args):
        self.suche.Text = ''
        self.listview.ItemsSource = self.Liste_rohr
        self.altdatagrid = self.Liste_rohr

    def start(self, sender, args):        
        abfrage = self.classMeldung(nachricht='Haben Sie die alten Daten gespeichert?',yes=True,no_Text='No')
        abfrage.ShowDialog()
        if abfrage.gespeichert == False:
            self.classMeldung(nachricht='Bitte speichern Sie zunächst die alten Daten!').ShowDialog()
            return
        for el in self.Liste_luft:
            if el.checked:
                el.berechnung = 'Keine'
                el.berechnungsindex = 0
        for el in self.Liste_rohr:
            if el.checked:
                el.berechnung = 'Keine'
                el.berechnungsindex = 0
        self.listview.Items.Refresh()
        self.keineEinstellen.Item_Luft = self.Liste_luft
        self.keineEinstellen.Item_Rohr = self.Liste_rohr
        self.keineEinstellenEvent.Raise()
       
        
    def rueck(self, sender, args):
        if not self.datum_lv.SelectedItem:
            self.classMeldung(nachricht='Bitte wählen Sie zunächst Daten aus!').ShowDialog()
            return
            


        for el in self.Liste_luft:
            if el.name in self.config.alt1.keys():
                el.berechnung = self.config.alt1[el.name]
                el.berechnungsindex = self.DICT_TEXT_INDEX[el.berechnung]
        for el in self.Liste_rohr:
            if el.name in self.config.alt1.keys():
                el.berechnung = self.config.alt1[el.name]
                el.berechnungsindex = self.DICT_TEXT_INDEX[el.berechnung]
        self.listview.Items.Refresh()
        
        self.zuruecksetzen.altdaten = self.datum_lv.SelectedItem
        self.zuruecksetzenEvent.Raise()
        
    
    def doubleclick(self, sender, args):
        if not self.datum_lv.SelectedItem:
            self.classMeldung(nachricht='Bitte wählen Sie zunächst Daten aus!').ShowDialog()
            return
        item = self.datum_lv.SelectedItem

        altberechnung = self.classAltBerechnung(item.daten)
        altberechnung.Title = item.datum
        altberechnung.ShowDialog()

    def einstellen(self, sender, args):

        self.einstellenhand.Item_Luft = self.Liste_luft
        self.einstellenhand.Item_Rohr = self.Liste_rohr
        self.einstellenhandEvent.Raise()
        

    def selchanged(self, sender, args):
        SelectedIndex = sender.SelectedIndex
        if SelectedIndex == -1:return
        if not (self.alttime and self.warten):
            temp = self.time()
            try:
                if temp-self.alttime < 1:return
            except:pass
            abfrage = self.classMeldung(nachricht='Haben Sie die alte Daten gespeichert?',erinnern=True,yes=True,no_Text='Nein')
            abfrage.ShowDialog()
            if abfrage.gespeichert == False:
                self.alttime = self.time()
                sender.SelectedIndex = self.DICT_TEXT_INDEX[sender.DataContext.berechnung]
                self.warten = None
                return
            else:
                self.alttime = self.time()
                if abfrage.Zeit.Text == '5 min':
                    self.warten = 300
                elif abfrage.Zeit.Text == '10 min':
                    self.warten = 600
                elif abfrage.Zeit.Text == '15 min':
                    self.warten = 900
                elif abfrage.Zeit.Text == '20 min':
                    self.warten = 1200
                elif abfrage.Zeit.Text == '30 min':
                    self.warten = 1800
                elif abfrage.Zeit.Text == '45 min':
                    self.warten = 2700
                elif abfrage.Zeit.Text == '60 min':
                    self.warten = 3600
        else:
            temp = self.time()
            if temp - self.alttime > self.warten:
                temp = self.time()
                abfrage = self.classMeldung(nachricht='Haben Sie die alte Daten gespeichert?',erinnern=True,yes=True,no_Text='Nein')
                abfrage.ShowDialog()
                if abfrage.gespeichert == False:
                    self.alttime = self.time()
                    sender.SelectedIndex = self.DICT_TEXT_INDEX[sender.DataContext.berechnung]
                    self.warten = None
                    return
                else:
                    self.alttime = self.time()
                    if abfrage.Zeit.Text == '5 min':
                        self.warten = 300
                    elif abfrage.Zeit.Text == '10 min':
                        self.warten = 600
                    elif abfrage.Zeit.Text == '15 min':
                        self.warten = 900
                    elif abfrage.Zeit.Text == '20 min':
                        self.warten = 1200
                    elif abfrage.Zeit.Text == '30 min':
                        self.warten = 1800
                    elif abfrage.Zeit.Text == '45 min':
                        self.warten = 2700
                    elif abfrage.Zeit.Text == '60 min':
                        self.warten = 3600

        try:
            if sender.DataContext in self.listview.SelectedItems:
                for item in self.listview.SelectedItems:
                    try:
                        if SelectedIndex + 1 <= item.IsBerechnung.Count:
                            item.berechnungsindex = SelectedIndex
                            item.berechnung = self.DICT_INDEX_TEXT[SelectedIndex]
                            
                    except:
                        pass                 
                
                self.listview.Items.Refresh()                       
            else:
                pass
        except:
            pass            

    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.listview.SelectedItem is not None:
            try:
                if sender.DataContext in self.listview.SelectedItems:
                    for item in self.listview.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass                 
                    
                    self.listview.Items.Refresh()                       
                else:
                    pass
            except:
                pass
    def namechanged(self, sender, args):
        text = sender.Text
        context = sender.DataContext
        i = self.datum_lv.Items.IndexOf(context)
        if i == 0:
            self.config.altzeit1 = text
            self.script.save_config()
            return
        if i == 1:
            self.config.altzeit2 = text
            self.script.save_config()
            return
        if i == 2:
            self.config.altzeit3 = text
            self.script.save_config()
            return
        if i == 3:
            self.config.altzeit4 = text
            self.script.save_config()
            return
        if i == 4:
            self.config.altzeit5 = text
            self.script.save_config()
            return
        return


wind = AktuelleBerechnung(ITEMSSOURCE_LUFT,ITEMSSOURCE_ROHR)

wind.Show()