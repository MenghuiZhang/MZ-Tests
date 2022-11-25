# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,TaskDialogResult,TaskDialogCommonButtons,ExternalEvent
import Autodesk.Revit.DB as DB
from System.Collections.ObjectModel import ObservableCollection
from pyrevit import forms
from System.Windows import Visibility

class template(object):
    def __init__(self,index,name):
        self.index = index
        self.name = name

class SystemTyp(object):
    def __init__(self,elem,name,berechnung,ISBerechnung,Index):
        self.elem = elem
        self.name = name
        self.berechnung = berechnung
        self.checked = False
        self.IsBerechnung = ISBerechnung
        self.berechnungsindex = Index

class SystemTypforalt(object):
    def __init__(self,elem,name,berechnung):
        self.elem = elem
        self.name = name
        self.berechnung = berechnung
        self.checked = False

DICT_TEXT_INDEX = {'Alle':3,'Nur Volumenstrom':2,'Leistung':1,'Keine':0}
DICT_INDEX_TEXT = {3:'Alle',2:'Nur Volumenstrom',1:'Leistung',0:'Keine'}
ALLE = ObservableCollection[template]()

ALLE.Add(template(0,'Keine'))
ALLE.Add(template(1,'Leistung'))
ALLE.Add(template(2,'Nur Volumenstrom'))
ALLE.Add(template(3,'Alle'))


KEINELEISTUNG = ObservableCollection[template]()
KEINELEISTUNG.Add(template(0,'Keine'))
KEINELEISTUNG.Add(template(1,'Leistung'))


KEINELEISTUNGVOLUMEN = ObservableCollection[template]()
KEINELEISTUNGVOLUMEN.Add(template(0,'Keine'))
KEINELEISTUNGVOLUMEN.Add(template(1,'Leistung'))
KEINELEISTUNGVOLUMEN.Add(template(2,'Nur Volumenstrom'))

IS_BERECHNUNG = {
    'Zuluft' : ALLE,
    'Abluft': ALLE,
    'Fortluft': ALLE,
    'Andere Luft': KEINELEISTUNG,
    'Vorlauf': ALLE,
    'Rücklauf': ALLE,
    'Belüftung': KEINELEISTUNG,
    'Abwasser': KEINELEISTUNGVOLUMEN,
    'Warmwasser': ALLE,
    'Kaltwasser': ALLE,
    'Brandschutz - Nass': KEINELEISTUNG,
    'Brandschutz - Trocken': KEINELEISTUNG,
    'Brandschutz - Vorgesteuert': KEINELEISTUNG,
    'Brandschutz - Andere': KEINELEISTUNG,
    'Sonstige': KEINELEISTUNG,
    'Rohrformteil': ALLE,
    'Global': ALLE,
}

DICT_BERECHNUNG = {
    'Alle': DB.Mechanical.SystemCalculationLevel.All,
    'Nur Volumenstrom': DB.Mechanical.SystemCalculationLevel.Flow,
    'Leistung': DB.Mechanical.SystemCalculationLevel.Performance,
    'Keine': DB.Mechanical.SystemCalculationLevel.None,
}

class AltDaten(object):
    def __init__(self,datum,daten):
        self.daten = daten
        self.datum = datum

class AltBerechnung(forms.WPFWindow):
    def __init__(self,Liste):
        self.Liste = Liste
        forms.WPFWindow.__init__(self,'altdaten.xaml')
        self.listview.ItemsSource = self.Liste

class Meldung(forms.WPFWindow):
    def __init__(self,nachricht = None, erinnern = False, yes = False, no_Text = 'Close'):
        self.Visibility_Visible = Visibility.Visible
        self.Visibility_Hidden = Visibility.Hidden
        self._nachricht = nachricht
        self._erinnern = erinnern
        self._yes = yes
        self._no_Text = no_Text
        self.gespeichert = False
        forms.WPFWindow.__init__(self,'abfrage.xaml')
        if self._erinnern:
            self.erinnern.Visibility = self.Visibility_Visible
            
        else:
            self.erinnern.Visibility = self.Visibility_Hidden
            self.erinnern.Height = 0
            self.Height = 120
        
        if self._nachricht:
            self.nachricht.Text = self._nachricht
        
        if self._yes:
            self.yes.Visibility = self.Visibility_Visible
        else:
            self.yes.Visibility = self.Visibility_Hidden
        
        if self._no_Text:
            self.no.Content = self._no_Text
        self.Zeit.ItemsSource = ['5 min','10 min', '15 min', '20 min', '30 min', '45 min','60 min']
        self.Zeit.Text = '30 min'
        

    def ja(self, sender, args):
        self.gespeichert = True
        self.Close()
    def nein(self, sender, args):
        self.Close()
        return
        
        
DICT_BERECHNUNG_UMGEKEHRT = {
    'All':'Alle',
    'Flow': 'Nur Volumenstrom',
    'Performance': 'Leistung',
    'None': 'Keine',
}


class KeineEinstellen(IExternalEventHandler):
    def __init__(self):
        self.Item_Luft = None
        self.Item_Rohr = None
        
    def Execute(self,app):        
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Systemberechnung--Keine')
        t.Start()

        try:
            for el in self.Item_Luft:
                if el.checked:
                    try:
                        el.elem.CalculationLevel = DB.Mechanical.SystemCalculationLevel.None
                        
                    except:pass
            for el in self.Item_Rohr:
                if el.checked:
                    try:
                        el.elem.CalculationLevel = DB.Mechanical.SystemCalculationLevel.None
                        
                    except:pass
        except:pass

        t.Commit()
        t.Dispose()
    def GetName(self):
        return "setzt den Berechnungsmodus auf keine"
    
class Einstellen(IExternalEventHandler):
    def __init__(self):
        self.Item_Luft = None
        self.Item_Rohr = None
        
    def Execute(self,app):        
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Systemberechnung--Aktualisieren')
        t.Start()

        try:
            for el in self.Item_Luft:
                if el.checked:
                    try:
                        el.elem.CalculationLevel = DICT_BERECHNUNG[DICT_INDEX_TEXT[el.berechnungsindex]]
                    except:pass
            for el in self.Item_Rohr:
                if el.checked:
                    try:
                        el.elem.CalculationLevel = DICT_BERECHNUNG[DICT_INDEX_TEXT[el.berechnungsindex]]
                    except:pass
        except:pass

        t.Commit()
        t.Dispose()
    def GetName(self):
        return "setzt den Berechnungsmodus auf beliebige Auswahl"

class Zuruecksetzen(IExternalEventHandler):
    def __init__(self):
        self.altdaten = None
        self.Luft = None
        self.Rohr = None
        
    def Execute(self,app):        
        doc = app.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Systemberechnung--Zurück')
        t.Start()

        try:
            for el in self.altdaten.daten:
                try:
                    el.elem.CalculationLevel = DICT_BERECHNUNG[el.berechnung]
                except:pass

        except:pass

        t.Commit()
        t.Dispose()
    def GetName(self):
        return "setzt den Berechnungsmodus zurück"