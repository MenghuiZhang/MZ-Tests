# coding: utf8
from eventhandler import ItemsSource,EINAUS,ExternalEvent
from pyrevit import forms
import os
from IGF_log import getlog,getloglocal

__title__ = "Beschriftung ein-/ausblenden"
__doc__ = """
Beschriftung Ein-/Ausblenden

[2022.04.13]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

class WPF_UI(forms.WPFWindow):
    def __init__(self):
        self.liste = ItemsSource
        forms.WPFWindow.__init__(self,'window.xaml')
        self.treeView1.ItemsSource = ItemsSource
        self.einaus = EINAUS()
        self.einausevent = ExternalEvent.Create(self.einaus)

        self.set_icon(os.path.join(os.path.dirname(__file__), 'Test.png'))


    def all(self,sender,args):
        for item in self.treeView1.Items:
            item.checked = True
            for sub_item in item.children:
                sub_item.checked = True
                for sub_sub_item in sub_item.children:
                    sub_sub_item.checked = True
        self.treeView1.Items.Refresh()


    def kein(self,sender,args):
        for item in self.treeView1.Items:
            item.checked = False
            for sub_item in item.children:
                sub_item.checked = False
                for sub_sub_item in sub_item.children:
                    sub_sub_item.checked = False
        self.treeView1.Items.Refresh()

    def zu(self,sender,args):
        for item in self.treeView1.Items:
            item.expand = False
            for sub_item in item.children:
                sub_item.expand = False
        self.treeView1.Items.Refresh()

    def aus(self,sender,args):
        for item in self.treeView1.Items:
            item.expand = True
            for sub_item in item.children:
                sub_item.expand = True
        self.treeView1.Items.Refresh()


    def ab(self,sender,args):
        self.Close()
    
    def ok(self,sender,args):
        self.einaus.Liste = self.liste
        self.einausevent.Raise()
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        item = sender.DataContext
        if item.Art == 'Kategorie':
            for fam in item.children:
                fam.checked = Checked
                for typ in fam.children:
                    typ.checked = Checked
            self.treeView1.Items.Refresh()
            return
        if item.Art == 'Familie':
            for typ in item.children:
                typ.checked = Checked
            kate = item.parent
            Liste_typ = []
            for typ in kate.children:
                if typ.checked is True:
                    if 0 not in Liste_typ:
                        Liste_typ.append(0)
                elif typ.checked is False:
                    if 1 not in Liste_typ:
                        Liste_typ.append(1)
                else:
                    if 2 not in Liste_typ:
                        Liste_typ.append(2)
                
            if 2 in Liste_typ or len(Liste_typ) > 1:
                kate.checked = None
                self.treeView1.Items.Refresh()
                return
            else:
                if 0 in Liste_typ:
                    kate.checked = True
                    self.treeView1.Items.Refresh()
                    return
                if 1 in Liste_typ:
                    kate.checked = False
                    self.treeView1.Items.Refresh()
                    return
        fam = item.parent
        kate = fam.parent
        Liste_typ = []
        Liste_fam = []
        for typ in fam.children:
            if typ.checked is True:
                if 0 not in Liste_typ:
                    Liste_typ.append(0)
            elif typ.checked is False:
                if 1 not in Liste_typ:
                    Liste_typ.append(1)
            else:
                if 2 not in Liste_typ:
                    Liste_typ.append(2)
            
        if 2 in Liste_typ or len(Liste_typ) > 1:
            fam.checked = None
        else:
            if 0 in Liste_typ:
                fam.checked = True
            if 1 in Liste_typ:
                fam.checked = False

        for temp_fam in kate.children:
            if temp_fam.checked is True:
                if 0 not in Liste_fam:
                    Liste_fam.append(0)
            elif temp_fam.checked is False:
                if 1 not in Liste_fam:
                    Liste_fam.append(1)
            else:
                if 2 not in Liste_fam:
                    Liste_fam.append(2)
            
        if 2 in Liste_fam or len(Liste_fam) > 1:
            kate.checked = None
            self.treeView1.Items.Refresh()
            return
        else:
            if 0 in Liste_fam:
                kate.checked = True
                self.treeView1.Items.Refresh()
                return
            if 1 in Liste_fam:
                kate.checked = False
                self.treeView1.Items.Refresh()
                return
 
WPF_UI_ = WPF_UI()
WPF_UI_.Show()