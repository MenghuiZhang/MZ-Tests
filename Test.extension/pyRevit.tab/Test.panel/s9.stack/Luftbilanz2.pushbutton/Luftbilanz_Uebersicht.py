# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.UI import ExternalEvent
from pyrevit import forms
import os
from System.Text.RegularExpressions import Regex
from Luftbilanz_Klasse import Luftauslass,VSR
from Luftbilanz_EventHandler import ExtenalEventListe


XAML_FILES_DIR = os.path.dirname(__file__)
            
class MEPRaum_Uebersicht(forms.WPFWindow):
    def __init__(self,RED,BLACK,LISTE_MEP,VISIBLE,HIDDEN,DICT_MEP_ITEMSSOIRCE,schachte):
        self.red = RED
        self.black = BLACK
        self.schachte = schachte
        self.regex1 = Regex("[^0-9.]+")
        self.regex2 = Regex("[^0-9]+")
        self.regex3 = Regex("[^0-9-]+")
        self.externaleventliste = ExtenalEventListe()
        self.ListeMEP = LISTE_MEP
        self.alt_MEP = None
        self.path = ''
        self.externalevent = ExternalEvent.Create(self.externaleventliste)
        self.visible = VISIBLE
        self.hidden = HIDDEN
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR,'MEP.xaml'),handle_esc=False)
        self.mepraum_liste = DICT_MEP_ITEMSSOIRCE
        self.Raumnr.ItemsSource = sorted(self.mepraum_liste.keys())
        self.mepraum = None
        self.temp_luftauslass = ObservableCollection[Luftauslass]()
        self.temp_vsr = ObservableCollection[VSR]()
        self.temp_Iris = ObservableCollection[VSR]()
        self.warnung.Visibility= self.hidden
        self.Closed += self.ReleaseResources
    
    def TextBox_GotFocus(self,sender,e):
        tb = sender
        tb.SelectAll()
    
    def ReleaseResources(self,sender,e):
        self.externalevent.Dispose()
        self.externaleventliste = None

    def vsr_selectedfabrikat_changed(self,sender,e):
        if self.lv_vsr1_getrennt.SelectedItem:
            try:
                item = sender.DataContext
                item.vsrauswaelten()
                item.vsrueberpruefen()
            except Exception as e:
                pass
    
    def vsr_selectedtyp_changed(self,sender,e):
        if self.lv_vsr1_getrennt.SelectedItem:
            try:
                item = sender.DataContext
                item.VSR_Hersteller = item.DICT_DatenBank[item.Liste_Herstellertyp[item.Herstellertypindex]]
                item.vsrueberpruefen()
            except Exception as e:
                pass

    def set_up(self):
        self.GUI_nachtbetrieb()
        self.bezugsname.IsEnabled = True
        self.bezugsname.ItemsSource = sorted(self.mepraum.berechnung_nach.values())
        self.faktor.IsEnabled = True
        self.istReduziert.IsEnabled = True
        self.isnachtbetrieb.IsEnabled = True
        self.labmineingabe.IsEnabled = True
        self.labmaxeingabe.IsEnabled = True
        self.ab24heingabe.IsEnabled = True
        self.druckeingabe.IsEnabled = True

        self.isnachtbetrieb.IsChecked = True if self.mepraum.nachtbetrieb  else False
        self.istReduziert.IsChecked = True if self.mepraum.istreduziert  else False
        self.istiefenachtbetrieb.IsChecked = True if self.mepraum.tiefenachtbetrieb  else False

        if self.istReduziert.IsChecked:
            self.redufaktor.IsEnabled = True
        else:
            self.redufaktor.IsEnabled = True

        self.LW_nacht.Text = str(self.mepraum.NB_LW)
        self.von_nacht.Text = str(self.mepraum.nb_von.soll)
        self.bis_nacht.Text = str(self.mepraum.nb_bis.soll)
        self.redufaktor.Text = str(self.mepraum.vol_faktorreduzierung)

        self.LW_tnacht.Text = str(self.mepraum.T_NB_LW)
        self.von_tnacht.Text = str(self.mepraum.tnb_von.soll)
        self.bis_tnacht.Text = str(self.mepraum.tnb_bis.soll)

        self.labmineingabe.Text = str(self.mepraum.ab_lab_min.soll)      
        self.labmaxeingabe.Text = str(self.mepraum.ab_lab_max.soll)
        self.ab24heingabe.Text = str(self.mepraum.ab_24h.soll)
        self.druckeingabe.Text = str(self.mepraum.Druckstufe.soll)

        self.ebene.Text = str(self.mepraum.ebene)
        self.bezugsname.Text = str(self.mepraum.bezugsname)
        self.faktor.Text = str(self.mepraum.faktor)
        self.einheit.Text = str(self.mepraum.einheit)
        self.flaeche.Text = str(self.mepraum.flaeche)
        self.volumen.Text = str(self.mepraum.volumen)
        self.personen.Text = str(self.mepraum.personen)
        self.hoehe.Text = str(self.mepraum.hoehe)
        self.grundinfo.ItemsSource = self.mepraum.Uebersicht
        self.grundinfo.Items.Refresh()
        self.anlageninfo.ItemsSource = self.mepraum.Anlagen_info
        self.anlageninfo.Items.Refresh()
        self.schachtinfo.ItemsSource = self.mepraum.Schacht_info
        self.schachtinfo.Items.Refresh()

        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr1.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
        self.lv_vsr2.ItemsSource = self.mepraum.list_vsr2


        self.lv_auslass.ItemsSource = self.mepraum.list_auslass
        self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass

        self.lv_ueber_aus.ItemsSource = self.mepraum.list_ueber['Aus']
        self.lv_ueber_in.ItemsSource = self.mepraum.list_ueber['Ein']

        self.detail_min.ItemsSource = self.mepraum.Detail_Min
        self.detail_max.ItemsSource = self.mepraum.Detail_Max
        self.detail_nacht.ItemsSource = self.mepraum.Detail_Nacht
        self.detail_tnacht.ItemsSource = self.mepraum.Detail_Tnacht
        self.lv_auslass_unrelevant.ItemsSource = self.mepraum.Liste_RaumluftUnrelevant

        self.alt_MEP = self.mepraum

    def GUI_nachtbetrieb(self):
        if self.mepraum.nachtbetrieb:
            self.LW_nacht.IsEnabled = True
            self.von_nacht.IsEnabled = True
            self.bis_nacht.IsEnabled = True
            self.istiefenachtbetrieb.IsEnabled = True
            if self.mepraum.tiefenachtbetrieb:
                self.LW_tnacht.IsEnabled = True
                self.von_tnacht.IsEnabled = True
                self.bis_tnacht.IsEnabled = True
            else:
                self.LW_tnacht.IsEnabled = False
                self.von_tnacht.IsEnabled = False
                self.bis_tnacht.IsEnabled = False

        else:
            self.LW_nacht.IsEnabled = False
            self.von_nacht.IsEnabled = False
            self.bis_nacht.IsEnabled = False
            self.LW_tnacht.IsEnabled = False
            self.von_tnacht.IsEnabled = False
            self.bis_tnacht.IsEnabled = False
            self.istiefenachtbetrieb.IsEnabled = False
 
    def reduziertchanged(self,sender,e):
        if self.alt_MEP == self.mepraum:
            checked = sender.IsChecked
            if checked:
                self.mepraum.istreduziert = 1
                self.mepraum.vol_faktorreduzierung_berechnen = self.mepraum.vol_faktorreduzierung
                self.redufaktor.IsEnabled = True
            else:
                self.mepraum.istreduziert = 0
                self.mepraum.vol_faktorreduzierung_berechnen = 1
                self.redufaktor.IsEnabled = False
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()
        
    def reduziertfaktorchanged(self,sender,e):
        if self.alt_MEP == self.mepraum:
            self.mepraum.vol_faktorreduzierung = float(str(sender.Text).replace(',', '.'))
            self.mepraum.vol_faktorreduzierung_berechnen = self.mepraum.vol_faktorreduzierung
            self.mepraum.Tagesbetrieb_Berechnen()
            self.mepraum.Druckstufe_Berechnen()

    def labnachteingabe_changed(self,sender,e):
        return
    def labtnachteingabe_changed(self,sender,e):
        return
    
    def nachtbetriebchanged(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.nachtbetrieb = 1 if sender.IsChecked else 0
                self.GUI_nachtbetrieb()
                self.mepraum.update_Nacht_Ueber_Lab()
                self.mepraum.Nachtbetrieb_Berechnen()
            except Exception as e:print(e)

    def tnachtbetriebchanged(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.tiefenachtbetrieb = 1 if sender.IsChecked else 0
                self.GUI_nachtbetrieb()
                self.mepraum.update_Tnacht_Ueber_Lab()
                self.mepraum.Nachtbetrieb_Berechnen()
            except Exception as e:print(e)

    def labmineingabe_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.ab_lab_min.soll = round(float(str(sender.Text).replace(',', '.')))
                self.mepraum.update_lab_min()
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Nachtbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()
                
            except:pass

    def labmaxeingabe_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.ab_lab_max.soll = round(float(str(sender.Text).replace(',', '.')))
                self.mepraum.update_lab_max()
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Nachtbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()
            except:pass

    def ab24heingabe_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.ab_24h.soll = round(float(str(sender.Text).replace(',', '.')))
                self.mepraum.update_24h()
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Nachtbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()
            except:pass

    def druckeingabe_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.Druckstufe.soll = round(float(str(sender.Text).replace(',', '.')))
                self.mepraum.update_druck()
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Nachtbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()
            except:pass
    
    def hoehechanged(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.hoehe = int(float(str(sender.Text)))
                self.mepraum.volumen = round(self.mepraum.flaeche * self.mepraum.hoehe / 1000,2)
                self.volumen.Text = str(self.mepraum.volumen)
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Nachtbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()

            except Exception as e:print(e)

    def LW_nacht_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.NB_LW = float(str(sender.Text).replace(',', '.'))
                self.mepraum.Nachtbetrieb_Berechnen()

            except:pass

    def LW_tnacht_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.T_NB_LW = float(str(sender.Text).replace(',', '.'))
                self.mepraum.Nachtbetrieb_Berechnen()
            except:pass

    def von_nacht_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.nb_von.soll = float(str(sender.Text).replace(',', '.'))
                self.mepraum.Nachtbetrieb_Berechnen()
            except:pass

    def bis_nacht_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.nb_bis.soll = float(str(sender.Text).replace(',', '.'))
                self.mepraum.Nachtbetrieb_Berechnen()
            except:pass

    def von_tnacht_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.tnb_von.soll = float(str(sender.Text).replace(',', '.'))
                self.mepraum.Nachtbetrieb_Berechnen()
            except:pass

    def bis_tnacht_changed(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.tnb_bis.soll = float(str(sender.Text).replace(',', '.'))
                self.mepraum.Nachtbetrieb_Berechnen()
            except:pass

    def faktorchanged(self, sender, args):
        if self.alt_MEP == self.mepraum:
            try:
                self.mepraum.faktor = float(str(sender.Text).replace(',', '.'))
            except:pass
            try:
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()

            except:pass

    def bezugsnameselectionchanged(self, sender, args):
        if self.alt_MEP == self.mepraum:
            self.mepraum.bezugsname = self.bezugsname.SelectedItem.ToString()
            for el in self.mepraum.berechnung_nach.keys():
                if self.mepraum.berechnung_nach[el] == self.bezugsname.SelectedItem.ToString():
                    self.mepraum.bezugsnummer = el
            self.einheit.Text = self.mepraum.einheit = self.mepraum.einheit_liste[self.mepraum.bezugsnummer]
            try:
                self.mepraum.Tagesbetrieb_Berechnen()
                self.mepraum.Nachtbetrieb_Berechnen()
                self.mepraum.Druckstufe_Berechnen()
            except:pass

            for el in self.ListeMEP:
                if el.name == self.mepraum.Raumnr:
                    el.berechnng = self.mepraum.bezugsname

    def MEP_changed(self, sender, args):
        try:
            if self.Raumnr.SelectedIndex != -1:
                text = str(self.Raumnr.SelectedItem)
                self.rauman.IsEnabled = True
                self.rauman1.IsEnabled = True
                self.temp_luftauslass.Clear()
                self.temp_vsr.Clear()
                self.temp_Iris.Clear()
                if text in self.mepraum_liste:
                    self.mepraum = self.mepraum_liste[text]
                    self.warnung.Visibility= self.hidden
                    self.auswahl.Visibility= self.hidden
                    for el in self.mepraum.list_vsr:
                        el.checked = False
                    
                    try:self.set_up()
                    except Exception as e:print(e)
        except Exception as e:print(e)

    def MEP_Suche_changed(self, sender, args):
        try:
            temp = []
            text = self.suche.Text
            if text in [None,'']:
                self.suche.Text = ''
                self.Raumnr.ItemsSource = sorted(self.mepraum_liste.keys())
            else:
                for el in sorted(self.mepraum_liste.keys()):
                    if el.upper().find(text.upper()) != -1:
                        temp.append(el)
                self.Raumnr.ItemsSource = temp
        except:pass

    def VSR_MEPRaum_changed(self, sender, args):
        text = str(self.auswahl.SelectedItem)
        text_original = str(self.Raumnr.SelectedItem)
        if text != text_original:
            self.rauman.IsEnabled = False
            self.Rauminfo.Visibility = self.visible
            self.Rauminfo.Text = "Die angezeigte Rauminfomationen beziehen sich auf Raum => " + text
        else:
            self.rauman.IsEnabled = True
            self.Rauminfo.Visibility = self.hidden
        try:
            temp = None
            for el in self.lv_vsr.Items:
                if el.checked:
                    temp = el
                    break
            self.mepraum = self.mepraum_liste[text]
            self.set_up()
            self.temp_luftauslass.Clear()
            self.temp_Iris.Clear()
            self.temp_vsr.Clear()
            for el in temp.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass

            if temp.IsHaupt:
                for el_vsr in temp.List_VSR:
                    if self.mepraum.list_vsr1.Contains(el_vsr):
                        self.temp_vsr.Add(el)
                self.lv_vsr1_getrennt.ItemsSource = self.temp_vsr
                self.lv_vsr1.ItemsSource = self.temp_vsr

                for el_iris in temp.List_Iris:
                    if self.mepraum.list_vsr2.Contains(el_iris):
                        self.temp_Iris.Add(el)
                self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
                self.lv_vsr2.ItemsSource = self.temp_Iris
            elif temp.IsIris:
                pass
            else:
                for el_iris in temp.List_Iris:
                    if self.mepraum.list_vsr2.Contains(el_iris):
                        self.temp_Iris.Add(el)
                self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
                self.lv_vsr2.ItemsSource = self.temp_Iris

        except Exception as e:print(e)

    def VSR_checked_changed(self, sender, args):
        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        Checked = sender.IsChecked
        item = sender.DataContext

        self.temp_luftauslass.Clear()
        self.temp_vsr.Clear()
        self.temp_Iris.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            for el in self.lv_vsr1.Items:
                el.checked = False
            for el in self.lv_vsr2.Items:
                el.checked = False
            item.checked = True

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass
            

            for el_vsr in item.List_VSR:
                if self.mepraum.list_vsr1.Contains(el_vsr):
                    self.temp_vsr.Add(el_vsr)
            self.lv_vsr1_getrennt.ItemsSource = self.temp_vsr
            self.lv_vsr1.ItemsSource = self.temp_vsr

            for el_iris in item.List_Iris:
                if self.mepraum.list_vsr2.Contains(el_iris):
                    self.temp_Iris.Add(el_iris)
            self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
            self.lv_vsr2.ItemsSource = self.temp_Iris

            if len(item.liste_Raum) > 1:
                self.raumwechsel_hinweis.Visibility= self.visible
                self.warnung.Visibility= self.visible
                self.auswahl.Visibility= self.visible
                self.auswahl.ItemsSource = sorted(item.liste_Raum)
            else:
                self.raumwechsel_hinweis.Visibility= self.hidden
                self.warnung.Visibility= self.hidden
                self.auswahl.Visibility= self.hidden

        else:
            self.raumwechsel_hinweis.Visibility= self.hidden
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
            self.lv_auslass.ItemsSource = self.mepraum.list_auslass
            self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass
            self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
            self.lv_vsr2.ItemsSource = self.mepraum.list_vsr2
            self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
            self.lv_vsr1.ItemsSource = self.mepraum.list_vsr1
    
    def VSR1_checked_changed(self, sender, args):
        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2.ItemsSource = self.mepraum.list_vsr1
        Checked = sender.IsChecked
        item = sender.DataContext
        self.temp_luftauslass.Clear()
        self.temp_Iris.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            for el in self.lv_vsr1.Items:
                el.checked = False
            for el in self.lv_vsr2.Items:
                el.checked = False
            item.checked = True

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass
            
            for el_iris in item.List_Iris:
                if self.mepraum.list_vsr2.Contains(el_iris):
                    self.temp_Iris.Add(el_iris)
            self.lv_vsr2_getrennt.ItemsSource = self.temp_Iris
            self.lv_vsr2.ItemsSource = self.temp_Iris

            if len(item.liste_Raum) > 1:
                self.raumwechsel_hinweis.Visibility= self.visible
                self.warnung.Visibility= self.visible
                self.auswahl.Visibility= self.visible
                self.auswahl.ItemsSource = sorted(item.liste_Raum)
            else:
                self.raumwechsel_hinweis.Visibility= self.hidden
                self.warnung.Visibility= self.hidden
                self.auswahl.Visibility= self.hidden

        else:            
            self.raumwechsel_hinweis.Visibility= self.hidden
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
            self.lv_auslass.ItemsSource = self.mepraum.list_auslass
            self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass
            self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
            self.lv_vsr2.ItemsSource = self.mepraum.list_vsr2
    
    def VSR2_checked_changed(self, sender, args):
        self.lv_vsr_getrennt.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr.ItemsSource = self.mepraum.list_vsr0
        self.lv_vsr1_getrennt.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2.ItemsSource = self.mepraum.list_vsr1
        self.lv_vsr2_getrennt.ItemsSource = self.mepraum.list_vsr2
        self.lv_vsr1.ItemsSource = self.mepraum.list_vsr2
        Checked = sender.IsChecked
        item = sender.DataContext
        self.temp_luftauslass.Clear()
        if Checked:
            for el in self.lv_vsr.Items:
                el.checked = False
            for el in self.lv_vsr1.Items:
                el.checked = False
            for el in self.lv_vsr2.Items:
                el.checked = False
            item.checked = True

            for el in item.Auslass:
                if el.raumid == self.mepraum.elemid.ToString():
                    self.temp_luftauslass.Add(el)
            self.lv_auslass.ItemsSource = self.temp_luftauslass
            self.lv_auslass_getrennt.ItemsSource = self.temp_luftauslass
            

            if len(item.liste_Raum) > 1:
                self.raumwechsel_hinweis.Visibility= self.visible
                self.warnung.Visibility= self.visible
                self.auswahl.Visibility= self.visible
                self.auswahl.ItemsSource = sorted(item.liste_Raum)
            else:
                self.raumwechsel_hinweis.Visibility= self.hidden
                self.warnung.Visibility= self.hidden
                self.auswahl.Visibility= self.hidden

        else:
            self.lv_auslass.ItemsSource = self.mepraum.list_auslass
            self.lv_auslass_getrennt.ItemsSource = self.mepraum.list_auslass
            self.raumwechsel_hinweis.Visibility= self.hidden
            self.warnung.Visibility= self.hidden
            self.auswahl.Visibility= self.hidden
    
    def Auslass_Volumen_changed_max_Un(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_unrelevant.SelectedIndex != -1:
            if item in self.lv_auslass_unrelevant.SelectedItems:
                for el in self.lv_auslass_unrelevant.SelectedItems:
                    el.Luftmengenmax = item.Luftmengenmax
    
    def Auslass_Volumen_changed_tnacht_Un(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_unrelevant.SelectedIndex != -1:
            if item in self.lv_auslass_unrelevant.SelectedItems:
                for el in self.lv_auslass_unrelevant.SelectedItems:
                    el.Luftmengentnacht = item.Luftmengentnacht

    
    def Auslass_Volumen_changed_min_Un(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_unrelevant.SelectedIndex != -1:
            if item in self.lv_auslass_unrelevant.SelectedItems:
                for el in self.lv_auslass_unrelevant.SelectedItems:
                    el.Luftmengenmin = item.Luftmengenmin
    
    def Auslass_Volumen_changed_nacht_Un(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_unrelevant.SelectedIndex != -1:
            if item in self.lv_auslass_unrelevant.SelectedItems:
                for el in self.lv_auslass_unrelevant.SelectedItems:
                    el.Luftmengennacht = item.Luftmengennacht

                  
    def Auslass_Volumen_changed_max(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    if el.art in ['24h','LAB']:
                        continue
                    el.Luftmengenmax = item.Luftmengenmax
                    self.Auslass_Volumen_changed(el,'max')
                    self.mepraum.update_ist_von_auslass(el,'max')
        else:
            self.Auslass_Volumen_changed(item,'max')
            self.mepraum.update_ist_von_auslass(item,'max')
    
    def Auslass_Volumen_changed_tnacht(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    if el.art in ['24h','LAB']:
                        continue
                    el.Luftmengentnacht = item.Luftmengentnacht
                    self.Auslass_Volumen_changed(el,'tnacht')
                    self.mepraum.update_ist_von_auslass(el,'tnacht')
        else:
            self.Auslass_Volumen_changed(item,'tnacht')
            self.mepraum.update_ist_von_auslass(item,'tnacht')
    
    def Auslass_Volumen_changed_min(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    
                    if el.art in ['24h','LAB']:
                        continue
                    el.Luftmengenmin = item.Luftmengenmin
                    self.Auslass_Volumen_changed(el,'min')
                    self.mepraum.update_ist_von_auslass(el,'min')

        else:
            self.Auslass_Volumen_changed(item,'min')
            self.mepraum.update_ist_von_auslass(item,'min')
    
    def Auslass_Volumen_changed_nacht(self,sender,args):
        item = sender.DataContext
        if self.lv_auslass_getrennt.SelectedIndex != -1:
            if item in self.lv_auslass_getrennt.SelectedItems:
                for el in self.lv_auslass_getrennt.SelectedItems:
                    if el.art in ['24h','LAB']:
                        continue
                    el.Luftmengennacht = item.Luftmengennacht
                    self.Auslass_Volumen_changed(el,'nacht')
                    self.mepraum.update_ist_von_auslass(el,'nacht')
        else:
            self.Auslass_Volumen_changed(item,'nacht')
            self.mepraum.update_ist_von_auslass(item,'nacht')

    def Auslass_Volumen_changed(self, auslass,zustand):
        try:
            vsr = auslass.VSR_Class
            if vsr:
                vsr.Luftmengenermitteln_new1(zustand)
                vsr.vsrauswaelten()
                vsr.vsrueberpruefen()
            iris = auslass.IRIS_Class
            if iris:
                iris.Luftmengenermitteln_new1(zustand)
            haupt = auslass.Haupt_Class
            if haupt:
                haupt.Luftmengenermitteln_new1(zustand)
                haupt.vsrauswaelten()
                haupt.vsrueberpruefen()
        except:pass
        
    def Auslass_Selected_Changed(self, sender, args):
        if self.lv_auslass.SelectedIndex != -1 and self.lv_auslass_getrennt.SelectedIndex != self.lv_auslass.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_auslass.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Auslass')
            self.lv_auslass_getrennt.SelectedIndex = self.lv_auslass.SelectedIndex
        else:
            return
    
    def Auslass_Selected_Changed_getrennt(self, sender, args):
        if self.lv_auslass_getrennt.SelectedIndex != -1 and self.lv_auslass_getrennt.SelectedIndex != self.lv_auslass.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_auslass_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Auslass')
            self.lv_auslass.SelectedIndex = self.lv_auslass_getrennt.SelectedIndex
        else:
            return
    
    def VSR_Selected_Changed(self, sender, args):
        if self.lv_vsr.SelectedIndex != -1 and self.lv_vsr_getrennt.SelectedIndex != self.lv_vsr.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Haupt')
            self.lv_vsr_getrennt.SelectedIndex = self.lv_vsr.SelectedIndex
        else:
            return
    
    def VSR_Selected_Changed_getrennt(self, sender, args):
        if self.lv_vsr_getrennt.SelectedIndex != -1 and self.lv_vsr_getrennt.SelectedIndex != self.lv_vsr.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Haupt')
            self.lv_vsr.SelectedIndex = self.lv_vsr_getrennt.SelectedIndex
        else:
            return   
        
    def VSR1_Selected_Changed(self, sender, args):
        if self.lv_vsr1.SelectedIndex != -1 and self.lv_vsr1_getrennt.SelectedIndex != self.lv_vsr1.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr1.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('VSR')
            self.lv_vsr1_getrennt.SelectedIndex = self.lv_vsr1.SelectedIndex
        else:
            return
    
    def VSR1_Selected_Changed_getrennt(self, sender, args):
        if self.lv_vsr1_getrennt.SelectedIndex != -1 and self.lv_vsr1_getrennt.SelectedIndex != self.lv_vsr1.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr1_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('VSR')
            self.lv_vsr1.SelectedIndex = self.lv_vsr1_getrennt.SelectedIndex
        else:
            return
    
    def VSR2_Selected_Changed(self, sender, args):
        if self.lv_vsr2.SelectedIndex != -1 and self.lv_vsr2_getrennt.SelectedIndex != self.lv_vsr2.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr2.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Iris')
            self.lv_vsr2_getrennt.SelectedIndex = self.lv_vsr2.SelectedIndex
        else:
            return
   
    def VSR2_Selected_Changed_getrennt(self, sender, args):
        if self.lv_vsr2_getrennt.SelectedIndex != -1 and self.lv_vsr2_getrennt.SelectedIndex != self.lv_vsr2.SelectedIndex:
            self.externaleventliste.Liste = [self.lv_vsr2_getrennt.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Iris')
            self.lv_vsr2.SelectedIndex = self.lv_vsr2_getrennt.SelectedIndex
        else:
            return

    def Ueber_aus_Selected_Changed(self, sender, args):
        if self.lv_ueber_aus.SelectedIndex != -1:
            self.externaleventliste.Liste = [self.lv_ueber_aus.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Üeberaus')
        else:
            return
    
    def Un_Auslass_Selected_Changed_getrennt(self,sender,e):
        if self.lv_auslass_unrelevant.SelectedIndex != -1:
            self.externaleventliste.Liste = [self.lv_auslass_unrelevant.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Unrelevant')
        else:
            return

    def Ueber_in_Selected_Changed(self, sender, args):
        if self.lv_ueber_in.SelectedIndex != -1:
            self.externaleventliste.Liste = [self.lv_ueber_in.SelectedItem.elemid]
            self.externaleventliste.Executeapp = self.externaleventliste.SelectElements
            self.externalevent.Raise()
            self.SelectedIndex('Üeberin')
        else:
            return
    
    def SelectedIndex(self,Art):
        if Art != 'Haupt':
            self.lv_vsr.SelectedIndex = -1
            self.lv_vsr_getrennt.SelectedIndex = -1
        if Art != 'VSR':
            self.lv_vsr1.SelectedIndex = -1
            self.lv_vsr1_getrennt.SelectedIndex = -1
        if Art != 'Iris':
            self.lv_vsr2.SelectedIndex = -1
            self.lv_vsr2_getrennt.SelectedIndex = -1
        if Art != 'Auslass':
            self.lv_auslass.SelectedIndex = -1
            self.lv_auslass_getrennt.SelectedIndex = -1
        if Art != 'Üeberaus':
            self.lv_ueber_aus.SelectedIndex = -1
        if Art != 'Üeberin':
            self.lv_ueber_in.SelectedIndex = -1
        if Art != 'Unrelevant':
            self.lv_auslass_unrelevant.SelectedIndex = -1
    
    def exportieren(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.ExportRaumluftbilanz
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def exportieren_schacht(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.ExportLuftmengenInSchacht
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def raumanzeigen(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.RaumAnzeigen
            self.externalevent.Raise()
        except Exception as e:print(e)

    def bauteilundraumanzeigen(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.RaumundBauteileanzeigen
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def Berechnen(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.RaumluftBerechnen
            self.externalevent.Raise()
        except Exception as e:print(e)

    def changevsrtype(self, sender, args):
        try:
            
            self.externaleventliste.Executeapp = self.externaleventliste.vsranpassen
            self.externalevent.Raise()
        except Exception as e:print(e)

    def luftmenge_gleich_verteilen_mep(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.LuftmengeverteilenMEP
            self.externalevent.Raise()
        except Exception as e:print(e)

    def luftmenge_gleich_verteilen_pro(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.LuftmengeverteilenProjekt
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def aenderung_uebernehmen_mep(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.AederungUebernehmenMEP
            self.externalevent.Raise()
        except Exception as e:print(e)

    def aenderung_uebernehmen_pro(self, sender, args):
        try:
            self.externaleventliste.Executeapp = self.externaleventliste.AederungUebernehmenProjekt
            self.externalevent.Raise()
        except Exception as e:print(e)
    
    def aus_revit_uebernehmen(self, sender, args):
        try:
            for el in self.mepraum.list_auslass:
                el.set_up()
                self.mepraum.update_ist_min(el)
                self.mepraum.update_ist_max(el)
                self.mepraum.update_ist_nacht(el)
                self.mepraum.update_ist_tnacht(el)

            for el in self.mepraum.list_vsr:
                el.set_up()
                el.Luftmengenermitteln_new()

        except Exception as e:print(e)
    
    def Textinput0(self, sender, args):
        try:args.Handled = self.regex2.IsMatch(args.Text)
        except:args.Handled = True
    
    def Textinput1(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find(',') != -1:
                if args.Text == ',':args.Handled = True
                return
        try:args.Handled = self.regex1.IsMatch(args.Text)
        except:args.Handled = True
    
    def Textinput2(self, sender, args):
        if sender.Text not in [None,'']:
            if sender.Text.find('-') != -1:
                if args.Text == '-':args.Handled = True
                return
        try:args.Handled = self.regex1.IsMatch(args.Text)
        except:args.Handled = True
    