"""Reduce the list of open Revit tabs.

Shift+Click:
Customize which tabs should be hidden in minified mode.
"""
#pylint: disable=C0103,E0401
from pyrevit import forms
from pyrevit.coreutils import ribbon
from pyrevit import script
from pyrevit.coreutils.ribbon import ICON_MEDIUM

config = script.get_config()

mlogger = script.get_logger()



class TabOption(forms.TemplateListItem):
    def __init__(self, uitab, hidden_tabs_list):
        super(TabOption, self).__init__(uitab)
        self.state = self.name in hidden_tabs_list

    @property
    def name(self):
        return self.item.name


def set_minifyui_config(hidden_tabs_list, config):
    config.hidden_tabs = hidden_tabs_list
    script.save_config()


def get_minifyui_config(config):
    return config.get_option('hidden_tabs', [])


def config_minifyui(config):
    this_ext_name = script.get_extension_name()
    hidden_tabs = get_minifyui_config(config)
    tabs = forms.SelectFromList.show(
        [TabOption(x, hidden_tabs) for x in ribbon.get_current_ui()
         if x.name not in this_ext_name],
        title='Module ein-/ausblenden',
        button_name='Hide Selected Modul',
        multiselect=True
        )
    

    if tabs:
        set_minifyui_config([x.name for x in tabs if x], config)


def update_ui(config):
    hidden_tabs = get_minifyui_config(config)
    for tab in ribbon.get_current_ui():
        if tab.name in hidden_tabs:
            # not new state since the visible value is reverse
            tab.visible = False
        else:
            tab.visible = True

def toggle_minifyui(config):
    new_state = not script.get_envvar(MINIFYUI_ENV_VAR)
    script.set_envvar(MINIFYUI_ENV_VAR, new_state)
    script.toggle_icon(new_state)
    update_ui(config)


config_minifyui(config)
update_ui(config)


# FIXME: need to figure out a way to fix the icon sizing of toggle buttons
# def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
#     off_icon = script_cmp.get_bundle_file('off.png')
#     ui_button_cmp.set_icon(off_icon, icon_size=ICON_MEDIUM)


# if __name__ == '__main__':
hidden_tabs = get_minifyui_config(config)
if len(hidden_tabs) > 0:
    icon = script.get_bundle_file('on.png')
    button = script.get_button()
    button.set_icon(icon, icon_size=ICON_MEDIUM)
    update_ui(config)
else:
    icon = script.get_bundle_file('off.png')
    button = script.get_button()
    button.set_icon(icon, icon_size=ICON_MEDIUM)
    update_ui(config)