# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms
from eventhandler import ListeExternalEvent,ExternalEvent,IS_Params,IS_Types

__title__ = "Sichtbarkeit von Familie"

__doc__ = """

[2022.12.07]
Version: 1.0
"""
__authors__ = "Menghui Zhang"


try:getlog()
except:pass
try:getloglocal()
except:pass

class GUI(forms.WPFWindow):
    def __init__(self):
        self.IS_Params = IS_Params
        self.IS_Types = IS_Types
        self.externaleventhandler = ListeExternalEvent()
        self.externalevent = ExternalEvent.Create(self.externaleventhandler)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.LV_typs.ItemsSource = self.IS_Types
        self.LV_Params.ItemsSource = self.IS_Params

        
    def einstellen(self, sender, args):
        self.externaleventhandler.ExcuseApp = self.externaleventhandler.Sichtbarkeiteinstellen
        self.externalevent.Raise()
    
    def close(self, sender, args):
        self.Close()
    
    def Param_einsetllen(self):
        for param in self.IS_Params:
            liste = []
            for typ in self.IS_Types:
                if typ.checked:
                    paramvalue =typ.dict_values[param.paramname]
                    if paramvalue not in liste:liste.append(paramvalue)

            if len(liste) == 1:
                if liste[0] == 0:
                    param.wert = False
                else:
                    param.wert = True
            else:
                param.wert = None

    
    def typchanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.LV_typs.SelectedIndex != -1:
            if item in self.LV_typs.SelectedItems:
                for el in self.LV_typs.SelectedItems:el.checked = checked
        self.Param_einsetllen()
    
    def paramchecked(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.LV_Params.SelectedIndex != -1:
            if item in self.LV_Params.SelectedItems:
                for el in self.LV_Params.SelectedItems:el.checked = checked
    
    def wertchanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.LV_Params.SelectedIndex != -1:
            if item in self.LV_Params.SelectedItems:
                for el in self.LV_Params.SelectedItems:el.wert = checked

    

    
            
gui = GUI()
gui.externaleventhandler.GUI = gui
gui.Show()


