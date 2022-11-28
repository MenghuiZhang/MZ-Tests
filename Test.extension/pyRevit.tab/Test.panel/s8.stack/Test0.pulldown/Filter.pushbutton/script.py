# coding: utf8
from IGF_log import getloglocal
from rpw import revit,DB
from pyrevit import script, forms
from externalevent import ExternalEvent,ExternalEventListe,ObservableCollection,Bauteil
import os


__title__ = "Filter"
__doc__ = """

[2022.11.21]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

try:
    getloglocal(__title__)
except:
    pass

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc

class GUI(forms.WPFWindow):
    def __init__(self):
        self.temp_coll = ObservableCollection[Bauteil]()    
        self.externaleventhandler = ExternalEventListe()
        self.event = ExternalEvent.Create(self.externaleventhandler)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
  
    def catechanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.LV_category.SelectedIndex != -1:
            if item in self.LV_category.SelectedItems:
                for el in self.LV_category.SelectedItems:el.checked = checked
        self.cateorychanged()

    def cateorychanged(self):
        self.temp_coll.Clear()
        for el in self.LV_category.Items:
            if el.checked:
                for item in el.Items:
                    self.temp_coll.Add(item)
            else:
                for item in el.Items:item.checked = False
        self.DG_Familie.ItemsSource = self.temp_coll

    def familiechanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.DG_Familie.SelectedIndex != -1:
            if item in self.DG_Familie.SelectedItems:
                for el in self.DG_Familie.SelectedItems:el.checked = checked

    def ausrevit(self,sender,e):
        self.externaleventhandler.Executeapp = self.externaleventhandler.Reset
        self.event.Raise()
  
    def inrevit(self,sender,e):
        self.externaleventhandler.Executeapp = self.externaleventhandler.PostSelect
        self.event.Raise()

gui = GUI()
gui.externaleventhandler.GUI = gui
gui.externaleventhandler.Executeapp = gui.externaleventhandler.Reset
gui.event.Raise()
gui.Show()