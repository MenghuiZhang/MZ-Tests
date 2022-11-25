# pylint: skip-file
from pyrevit import HOST_APP, EXEC_PARAMS
from pyrevit import revit, script
import clr
clr.AddReference('AdWindows')
from Autodesk.Windows import ComponentManager

args = EXEC_PARAMS.event_args

def set_visible(item,visible):
    try:
        item.Visible = visible
    except:
        pass
    try:
        item.IsVisible = visible
    except:
        pass

DICT_ALLTABS = {}
all_module = ComponentManager.Ribbon.Tabs
for modul in all_module:
    revit_obj = modul.GetType()
    if revit_obj.ToString() == 'Autodesk.Windows.RibbonTab':
        DICT_ALLTABS[modul.Title] = modul

config = script.get_config('Module ausblenden')
try:
    try:
        hidden_tabs = config.hidetabs
    except:
        hidden_tabs = []
    try:
        hidden_panels = config.hidepanels
    except:
        hidden_panels = {}
    for el_Tab in DICT_ALLTABS.keys():
        if el_Tab in hidden_tabs:
            set_visible(DICT_ALLTABS[el_Tab],False)
            continue
        else:
            set_visible(DICT_ALLTABS[el_Tab],True)
        if el_Tab in hidden_panels.keys():
            hidden_items = hidden_panels[el_Tab]
            for el_panel in DICT_ALLTABS[el_Tab].Panels:
                panelname = el_panel.Source.AutomationName
                if el_panel.Source.Items.Count == 0:
                    set_visible(el_panel,False)
                else:
                    if panelname in hidden_items:
                        set_visible(el_panel,False)
                    else:
                        set_visible(el_panel,True)
except Exception as e:print(e)