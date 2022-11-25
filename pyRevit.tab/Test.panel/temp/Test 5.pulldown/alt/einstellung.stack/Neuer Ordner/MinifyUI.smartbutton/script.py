# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from System.Collections.ObjectModel import ObservableCollection
from IGF_log import getlog
from pyrevit import forms,revit
from pyrevit import script
from pyrevit.coreutils import ribbon

__title__ = "ausgewählte Tab ausblenden"
__doc__ = """

blendet ausgewählte Module aus.

[2021.12.06]
Version: 1.1
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
doc = revit.doc

class Tab(object):
    def __init__(self,name,checked):
        self.name = name
        self.checked = checked

liste_Module = []
all_module = ribbon.get_current_ui()
if doc.IsFamilyDocument:
    config = script.get_config('doc Tabs aus')
else:
    config = script.get_config('family Tabs aus')
for el in all_module:
    if doc.IsFamilyDocument:
        if not el.name in ['Architektur','Ingenieurbau','Stahlbau','Gebäudetechnik','Körpermodell & Grundstück','Zusammenarbeit','Berechnung']:
            liste_Module.append(el.name)
    else:
        if not el.name in ['Erstellen','Familieneditor']:
            liste_Module.append(el.name)


ausgeblenden_module = []
for el in all_module:
    if not el.visible:
        ausgeblenden_module.append(el)


liste_Module.sort()
if 'pyRevit' in liste_Module:
    liste_Module.remove('pyRevit')

Liste_Tab = ObservableCollection[Tab]()

for el in liste_Module:
    Liste_Tab.Add(Tab(el,False))

class WPF_UI(forms.WPFWindow):
    def __init__(self, xaml_file_name,liste_module):
        self.liste_module = liste_module
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.LVModule.ItemsSource = liste_module
        self.read_config()

    def read_config(self):
        try:
            hidetabs = config.hidetabs
            for el in self.liste_module:
                if (el.name in hidetabs) or (el.name in ausgeblenden_module):
                    el.checked = True
            self.LVModule.Items.Refresh()

        except:
            try:
                config.hidetabs = []
            except:
                pass

    def write_config(self):
        try:
            hide_tabs = []
            for el in self.liste_module:
                if el.checked:
                    hide_tabs.append(el.name)
            config.hidetabs = hide_tabs
        except:
            pass
        script.save_config()
    
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LVModule.SelectedItem is not None:
            try:
                if sender.DataContext in self.LVModule.SelectedItems:
                    for item in self.LVModule.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.LVModule.Items.Refresh()
                else:
                    pass
            except:
                pass


    def ok(self,sender,args):
        self.Close()
        self.write_config()
    def checkall(self,sender,args):
        for item in self.LVModule.Items:
            item.checked = True
        self.LVModule.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.LVModule.Items:
            item.checked = False
        self.LVModule.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.LVModule.Items:
            value = item.checked
            item.checked = not value
        self.LVModule.Items.Refresh()

    def abbrechen(self,sender,args):
        self.Close()
        script.exit()


def main(button = None):
    hidden_tabs = config.hidetabs

    if len(hidden_tabs) > 0:
        icon = script.get_bundle_file('on.png')
        if not button:
            button = script.get_button()
        button.set_icon(icon)
    else:
        icon = script.get_bundle_file('off.png')
        if not button:
            button = script.get_button()
        button.set_icon(icon)

    for el in all_module:
        if el.name in liste_Module:
            if el.name in hidden_tabs:
                el.visible = False
            else:
                el.visible = True
if __name__ == '__main__':
    try:
        getlog(__title__)
    except:
        pass
    ui_tabs = WPF_UI("window.xaml",Liste_Tab)
    try:
        ui_tabs.ShowDialog()
    except Exception as e:
        logger.error(e)
        ui_tabs.Close()
        script.exit()
    main()

# FIXME: need to figure out a way to fix the icon sizing of toggle buttons
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    hidden_tabs = config.hidetabs

    if len(hidden_tabs) > 0:
        icon = script_cmp.get_bundle_file('on.png')
        ui_button_cmp.set_icon(icon)
    else:
        icon = script_cmp.get_bundle_file('off.png')
        ui_button_cmp.set_icon(icon)

    for el in all_module:
        if el.name in liste_Module:
            if el.name in hidden_tabs:
                el.visible = False
            else:
                el.visible = True
                