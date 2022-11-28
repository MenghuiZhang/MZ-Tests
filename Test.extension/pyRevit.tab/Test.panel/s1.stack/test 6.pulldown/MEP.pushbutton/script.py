# coding: utf8
from webbrowser import get
from pyrevit import script, forms
from rpw import revit,DB
from IGF_log import getlog,getloglocal
from IGF_lib import get_value
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent


__title__ = "Nachtsbetrieb"
__doc__ = """Luftmengenberechnung für Projekt MFC"""
__author__ = "Menghui Zhang"

try:getlog(__title__)
except:pass

try:getloglocal(__title__)
except:pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))

if not spaces:
    logger.error("Keine MEP Räume in aktueller Ansicht gefunden")
    script.exit()

berechnung_nach = {
    "Fläche": '1',
    "Luftwechsel": '2',
    "Person": '3',
    "manuell": '4',
    "nurZUMa": '5',
    "nurABMa": '6',
    "nurZU_Fläche": '5.1',
    "nurZU_Luftwechsel": '5.2',
    "nurZU_Person": '5.3',
    "nurAB_Fläche": '6.1',
    "nurAB_Luftwechsel": '6.2',
    "nurAB_Person": '6.3',
    "keine": '9'
}


Dict_Ueber_Manuell = {}
for ele in spaces_collector:
    raum = get_value(ele.LookupParameter("TGA_RLT_RaumÜberströmungAus"))
    if raum:
        summe2 = get_value(ele.LookupParameter('TGA_RLT_RaumÜberströmungMenge'))
        if raum not in Dict_Ueber_Manuell.keys():
            Dict_Ueber_Manuell[raum] = summe2
        else:Dict_Ueber_Manuell[raum] += summe2

class MEPRaum(object):
    def __init__(self, element_id):
        """
        Definiert MEP Raum Klasse mit allen object properties für die
        Luftmengen Berechnung.
        """
        self.checked = False
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.name = get_value(self.element.LookupParameter('Name'))
        self.nummer = self.element.Number
        self.isnachtbetrieb = get_value(self.element.LookupParameter('IGF_RLT_Nachtbetrieb'))
        self.von = get_value(self.element.LookupParameter('IGF_RLT_NachtbetriebVon'))
        self.bis = get_value(self.element.LookupParameter('IGF_RLT_NachtbetriebBis'))
        self.LW = get_value(self.element.LookupParameter('IGF_RLT_NachtbetriebLW'))
        self.istiefernachtbetrieb = get_value(self.element.LookupParameter('IGF_RLT_TieferNachtbetrieb'))
        self.tvon = get_value(self.element.LookupParameter('IGF_RLT_TieferNachtbetriebVon'))
        self.tbis = get_value(self.element.LookupParameter('IGF_RLT_TieferNachtbetriebBis'))
        self.tLW = get_value(self.element.LookupParameter('IGF_RLT_TieferNachtbetriebLW'))
        self.volumen = get_value(self.element.LookupParameter('Volumen'))
        self.raum_druckstufe = get_value(self.element.LookupParameter('IGF_RLT_RaumDruckstufeEingabe'))
        self.ueberstroemungIn = get_value(self.element.LookupParameter('IGF_RLT_ÜberströmungSummeIn'))
        self.ueberstroemungAus = get_value(self.element.LookupParameter('IGF_RLT_ÜberströmungSummeAus'))
        self.ueberstroemung2 = get_value(self.element.LookupParameter('TGA_RLT_RaumÜberströmungMenge'))
        self.ABL_24h = get_value(self.element.LookupParameter('IGF_RLT_AbluftminRaumL24h'))
        self.ABL_Labor = get_value(self.element.LookupParameter('IGF_RLT_AbluftminSummeLabor'))
        self.kez = self.element.LookupParameter("TGA_RLT_VolumenstromProNummer").AsValueString()

        try:self.ueberAusMa = Dict_Ueber_Manuell[self.nummer]
        except:self.ueberAusMa = 0

        self.abluft_labor_24h = self.ABL_Labor + self.ABL_24h

        self.nb_dauer = 0
        self.zu_nacht = 0
        self.ab_nacht = 0
        self.tiefer_nb_dauer = 0
        self.tiefer_zu_nacht = 0
        self.tiefer_ab_nacht = 0

    def luft_round(self,luft):
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
   
    def Nachtbetrieb_Berechnen(self):

        if self.isnachtbetrieb:
            self.nb_dauer = self.bis - self.von + 24.00
            if self.kez == berechnung_nach["Fläche"] or self.kez == berechnung_nach["Luftwechsel"] or self.kez == berechnung_nach["Person"] or self.kez == berechnung_nach["manuell"]:
                self.zu_nacht = self.luft_round(self.LW * self.volumen) + self.raum_druckstufe
                abweichung = self.zu_nacht + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.ueberAusMa - self.abluft_labor_24h - self.raum_druckstufe
                if abweichung < 0:
                    self.zu_nacht -= abweichung
                    self.ab_nacht = 0
                else:
                    self.ab_nacht = abweichung
            elif self.kez == berechnung_nach["nurZU_Fläche"] or self.kez == berechnung_nach["nurZU_Luftwechsel"] or self.kez == berechnung_nach["nurZU_Person"] or self.kez == berechnung_nach["nurZUMa"]:
                self.zu_nacht = self.luft_round(self.LW * self.volumen) + self.raum_druckstufe
                abweichung = self.zu_nacht + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.ueberAusMa - self.abluft_labor_24h - self.raum_druckstufe
                if abweichung < 0:
                    
                    self.ab_nacht = 0
                else:
                    self.ab_nacht = 0
                self.zu_nacht -= abweichung
            
            elif self.kez == berechnung_nach["nurAB_Fläche"] or self.kez == berechnung_nach["nurAB_Luftwechsel"] or self.kez == berechnung_nach["nurAB_Person"] or self.kez == berechnung_nach["nurABMa"]:
                self.zu_nacht = 0
                abweichung = self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.ueberAusMa - self.abluft_labor_24h + self.raum_druckstufe
                if abweichung < 0:
                    self.zu_nacht -= abweichung
                    self.ab_nacht = 0
                else:
                    self.ab_nacht = abweichung



        if self.istiefernachtbetrieb:
            self.tiefer_nb_dauer = self.tbis - self.tvon + 24.00
            self.nb_dauer -= self.tiefer_nb_dauer
            self.tiefer_zu_nacht = self.luft_round(self.tLW * self.volumen) + self.raum_druckstufe
            abweichung = self.tiefer_zu_nacht + self.ueberstroemungIn + self.ueberstroemung2 - self.ueberstroemungAus - self.ueberAusMa - self.abluft_labor_24h - self.raum_druckstufe
            if abweichung < 0:
                self.tiefer_zu_nacht -= abweichung
                self.tiefer_ab_nacht = self.abluft_labor_24h
            else:
                self.tiefer_ab_nacht = self.abluft_labor_24h + abweichung

    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                if self.element.LookupParameter(param_name):
                    if self.element.LookupParameter(param_name).IsReadOnly is True:
                        logger.error(self.element_id)
                        logger.error(param_name)
                        return
                    self.element.LookupParameter(param_name).SetValueString(str(wert))

        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.zu_nacht)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.ab_nacht)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer",self.tiefer_nb_dauer)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tiefer_zu_nacht)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tiefer_ab_nacht)
        # wert_schreiben("IGF_RLT_NachtbetriebVon", self.von)
        # wert_schreiben("IGF_RLT_NachtbetriebBis", self.bis)
        # wert_schreiben("IGF_RLT_NachtbetriebLW",self.LW)
        # wert_schreiben("IGF_RLT_TieferNachtbetriebVon", self.tvon)
        # wert_schreiben("IGF_RLT_TieferNachtbetriebBis", self.tbis)
        # wert_schreiben("IGF_RLT_TieferNachtbetriebLW",self.tLW)
        # if self.isnachtbetrieb:
        #     self.element.LookupParameter('IGF_RLT_Nachtbetrieb').Set(1)
        # else:
        #     self.element.LookupParameter('IGF_RLT_Nachtbetrieb').Set(0)
        # if self.istiefernachtbetrieb:
        #     self.element.LookupParameter('IGF_RLT_TieferNachtbetrieb').Set(1)
        # else:
        #     self.element.LookupParameter('IGF_RLT_TieferNachtbetrieb').Set(0)
dict_temp = []
with forms.ProgressBar(title="{value}/{max_value} MEP Räume",cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)
        if mepraum.nummer == '100.224b':continue
        if not mepraum.isnachtbetrieb:continue
        mepraum.Nachtbetrieb_Berechnen()
        
        dict_temp.append(mepraum) 

t =DB.Transaction(doc,'nachtbetrieb')
t.Start()
for el in dict_temp:
    el.werte_schreiben()

t.Commit()
# itemssource = ObservableCollection[MEPRaum]()
# for e in sorted(dict_temp.keys()):
#     itemssource.Add(dict_temp[e])

# class UPDATEMODEL(IExternalEventHandler):
#     def __init__(self):
#         self.GUI_MEPRaum = None

#     def Execute(self,app):
#         doc = app.ActiveUIDocument.Document
#         t = DB.Transaction(doc,'Wert schreiben')
#         t.Start()
#         for raum in self.GUI_MEPRaum.rooms:
#             if raum.checked == False:
#                 continue
#             try:
#                 raum.Nachtbetrieb_Berechnen()
#                 raum.werte_schreiben()
#             except Exception as e:print(e)

#         t.Commit()
#         t.Dispose()

#     def GetName(self):
#         return "wert schreiben"

# class _GUI(forms.WPFWindow):
#     def __init__(self):
#         self.rooms = itemssource
#         self.updateModel = UPDATEMODEL()
#         self.updateModelEvent = ExternalEvent.Create(self.updateModel)
#         forms.WPFWindow.__init__(self,'window.xaml')
#         self.lv.ItemsSource = self.rooms
#         self.temp = ObservableCollection[MEPRaum]()
    
#     def suchchanged(self,sender,a):
#         self.temp.Clear()
#         text = self.suche.Text.upper()
#         if text in [None,'']:
#             self.suche.Text = ''
#             self.lv.ItemsSource = self.rooms
#             return
#         for el in self.rooms:
#             temmpname = el.nummer+'-'+el.name
#             if temmpname.upper().find(text) != -1:
#                 self.temp.Add(el)
#         self.lv.ItemsSource = self.temp
#         self.lv.Items.Refresh()
    
#     def nachtbetriebchanged(self, sender, args):
#         Checked = sender.IsChecked
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.isnachtbetrieb = Checked
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
#     def nachtbetriebvonchanged(self, sender, args):
#         text = sender.Text
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.von = text
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
    
#     def nachtbetriebbischanged(self, sender, args):
#         text = sender.Text
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.bis = text
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
#     def tiefernachtbetriebchanged(self, sender, args):
#         Checked = sender.IsChecked
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.istiefernachtbetrieb = Checked
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
#     def tnachtbetriebvonchanged(self, sender, args):
#         text = sender.Text
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.tvon = text
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
    
#     def tnachtbetriebbischanged(self, sender, args):
#         text = sender.Text
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.tbis = text
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
    
#     def tnachtbetriebLWchanged(self, sender, args):
#         text = sender.Text
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.tLW = text
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 
#     def nachtbetriebLWchanged(self, sender, args):
#         text = sender.Text
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.LW = text
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass    
    
#     def checkedchanged(self, sender, args):
#         Checked = sender.IsChecked
#         if self.lv.SelectedItem is not None:
#             try:
#                 if sender.DataContext in self.lv.SelectedItems:
#                     for item in self.lv.SelectedItems:
#                         try:item.checked = Checked
#                         except:pass
#                     self.lv.Items.Refresh()                       
#                 else:
#                     pass
#             except:
#                 pass 

#     def OK(self, sender, args):
#         # self.updateModelEvent.Raise()
#         t = DB.Transaction(doc,'Wert schreiben')
#         t.Start()
#         for raum in self.rooms:
#             if raum.checked == False:
#                 continue
#             try:
#                 raum.Nachtbetrieb_Berechnen()
#                 raum.werte_schreiben()
#             except Exception as e:print(e)
#         t.Commit()
#         self.Close()

# mepWPF = _GUI()
# mepWPF.updateModel.GUI_MEPRaum = mepWPF
# mepWPF.ShowDialog()