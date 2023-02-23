# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.UI import TaskDialog,ExternalEvent
from System.Windows import GridLength
from pyrevit import script,forms
import os
from System.Windows.Forms import FolderBrowserDialog,DialogResult
from Luftbilanz_Klasse import Itemtemplate,MEPFOREXPORTITEM
from Luftbilanz_Config import config


XAML_FILES_DIR = os.path.dirname(__file__)
            
class ABFRAGE(forms.WPFWindow):
    def __init__(self,maininfo = '',checked=True,height = None,minmax = True):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR,'abfrage.xaml'))
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

class Familien(forms.WPFWindow):
    def __init__(self,IS_Auslass0,IS_Auslass1,IS_Auslass2,IS_Auslass3,IS_VSR,IS_Drossel):
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR,'Familien.xaml'))
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

class RaumluftbilanzExport(forms.WPFWindow):
    def __init__(self,path,Liste_MEP):
        self.path = path
        forms.WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR,'einstellung.xaml'))
        self.Liste_MEP = Liste_MEP
        self.liste_temp = ObservableCollection[MEPFOREXPORTITEM]()
        if os.path.exists(self.path):
            self.ordner.Text = self.path
        else:
            self.path = ''
        self.result = False
        self.LV_MEP.ItemsSource = self.Liste_MEP
    def suchetextchanged(self,sender,e):
        text = sender.Text
        self.liste_temp.Clear()
        if not text:
            self.LV_MEP.ItemsSource = self.Liste_MEP
            return
        else:
            for el in self.Liste_MEP:
                if el.name.upper().find(text.upper()) != -1:
                    self.liste_temp.Add(el)
            self.LV_MEP.ItemsSource = self.liste_temp
            return    
    def checkedchanged(self,sender,e):
        Checked = sender.IsChecked
        if self.LV_MEP.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_MEP.SelectedItems:
                    for item in self.LV_MEP.SelectedItems:
                        try:item.checked = Checked
                        except:pass
                else:pass
            except:pass

    def alle(self,sender,e):
        for el in self.LV_MEP.Items:
            el.checked = True

    def keine(self,sender,e):
        for el in self.LV_MEP.Items:
            el.checked = False
    
    def nur(self,sender,e):
        for el in self.LV_MEP.Items:
            if el.berechnng != 'keine':
                el.checked = True
    
    def ok(self,sender,e):
        if not self.ordner.Text:
            TaskDialog.Show('Info.','Kein Ordner ausgewählt')
            return
        elif os.path.exists(self.ordner.Text) is False:
            TaskDialog.Show('Info.','Ordner nicht vorhanden')
            self.ordner.Text = ''
            return
        self.result = True
        self.Close()
        return
    
    def schliessen(self,sender,e):
        self.result = False
        self.Close()
        return
    
    def movewindow(self, sender, args):
        self.DragMove()
    
    def durchsuchen(self,sender,args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Ordner auswählen"
        dialog.ShowNewFolderButton = True
        if dialog.ShowDialog() == DialogResult.OK:
            folder = dialog.SelectedPath
            self.ordner.Text = folder
            self.path = folder
