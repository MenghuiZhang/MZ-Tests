from pyrevit import forms
import os.path as op

__doc__ = 'Test_Panel'

class DockableExample(forms.WPFPanel):
    panel_title = "pyRevit Dockable Panel Title"
    panel_id = "3110e336-f81c-4927-87da-4e0d30d4d64b"
    panel_source = op.join(op.dirname(__file__), "DockableExample.xaml")

    def do_something(self, sender, args):
        forms.alert("Voila!!!")


if not forms.is_registered_dockable_panel(DockableExample):
    forms.register_dockable_panel(DockableExample)
else:
    print("Skipped registering dockable pane. Already exists.")

from pyrevit import forms


test_panel_uuid = "3110e336-f81c-4927-87da-4e0d30d4d64b"

forms.open_dockable_panel(test_panel_uuid)
