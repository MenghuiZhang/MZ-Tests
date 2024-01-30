# coding: utf8
from pyrevit import forms
from IGF_Funktionen._Parameter import wert_schreibenbase

class WPFWindow(forms.WPFWindow):
    def __init__(self,wpfgui,handle_esc=True):
        forms.WPFWindow.__init__(self,wpfgui,handle_esc=handle_esc)
        self.Closed += self.ReleaseResources
        self.externalevent = None
        self.externaleventliste = None

    def ReleaseResources(self,sender,e):
        self._releaseresource()
    
    def _releaseresource(self):
        self.externaleventliste = None
        try:
            self.externalevent.Dispose()
        except:
            pass

class WPFPanel(forms.WPFPanel):
    def __init__(self):
        forms.WPFPanel.__init__(self)
        self.externalevent = None
        self.externaleventliste = None
        # self.set_owner()
    
    def setup_owner(self):
        wih = forms.Interop.WindowInteropHelper(self)
        wih.Owner = forms.AdWindows.ComponentManager.ApplicationWindow
    
    @staticmethod
    def open_dockable_panel(panel_type_or_id):
        forms.open_dockable_panel(panel_type_or_id)
    
    @staticmethod
    def close_dockable_panel(panel_type_or_id):
        forms.open_dockable_panel(panel_type_or_id,False)
    
    @staticmethod
    def get_rvt__dockable_panel(panel_type_or_id):
        dpanel_id = None
        if isinstance(panel_type_or_id, str):
            panel_id = forms.coreutils.Guid.Parse(panel_type_or_id)
            try:dpanel_id = forms.UI.DockablePaneId(panel_id)
            except:return
        elif issubclass(panel_type_or_id, WPFPanel):
            panel_id = forms.coreutils.Guid.Parse(panel_type_or_id.panel_id)
            try:dpanel_id = forms.UI.DockablePaneId(panel_id)
            except:return
        else:
            raise forms.PyRevitException("Given type is not a forms.WPFPanel")
            return
        if dpanel_id:
            if forms.UI.DockablePane.PaneIsRegistered(dpanel_id):
                dockable_panel = forms.HOST_APP.uiapp.GetDockablePane(dpanel_id)
                return dockable_panel

    @staticmethod
    def toggle_dockable_panel(panel_type_or_id):
        """Toggle previously registered dockable panel

        Args:
            panel_type_or_id (forms.WPFPanel, str): panel type or id
        """
        dpanel_id = None
        if isinstance(panel_type_or_id, str):
            panel_id = forms.coreutils.Guid.Parse(panel_type_or_id)
            dpanel_id = forms.UI.DockablePaneId(panel_id)
        elif issubclass(panel_type_or_id, WPFPanel):
            panel_id = forms.coreutils.Guid.Parse(panel_type_or_id.panel_id)
            dpanel_id = forms.UI.DockablePaneId(panel_id)
        else:
            raise forms.PyRevitException("Given type is not a forms.WPFPanel")

        if dpanel_id:
            if forms.UI.DockablePane.PaneIsRegistered(dpanel_id):
                dockable_panel = forms.HOST_APP.uiapp.GetDockablePane(dpanel_id)
                if dockable_panel.IsShown():
                    dockable_panel.Hide()
                else:

                    dockable_panel.Show()
            else:
                raise forms.PyRevitException(
                    "Panel with id \"%s\" is not registered" % panel_type_or_id
                    )

    @staticmethod
    def get_dockable_panel(panel_type_or_id):
        """Toggle previously registered dockable panel

        Args:
            panel_type_or_id (forms.WPFPanel, str): panel type or id
        """
        dpanel_id = None
        if isinstance(panel_type_or_id, str):
            panel_id = forms.coreutils.Guid.Parse(panel_type_or_id)
            dpanel_id = forms.UI.DockablePaneId(panel_id)
        elif issubclass(panel_type_or_id, WPFPanel):
            panel_id = forms.coreutils.Guid.Parse(panel_type_or_id.panel_id)
            dpanel_id = forms.UI.DockablePaneId(panel_id)
        else:
            raise forms.PyRevitException("Given type is not a forms.WPFPanel")

        if dpanel_id:
            if forms.UI.DockablePane.PaneIsRegistered(dpanel_id):
                dockable_panel = forms.HOST_APP.uiapp.GetDockablePane(dpanel_id)
                return dockable_panel
                        
    @staticmethod
    def register_dockable_panel(panel_type, default_visible=True):
        forms.register_dockable_panel(panel_type, default_visible)
    
    @staticmethod
    def is_registered_dockable_panel(panel_type):
        return forms.is_registered_dockable_panel(panel_type)



class WPF_Funktionen:
    @staticmethod
    def ClearValue_Combobox(combobox):
        itemssource = combobox.ItemsSource
        combobox.ItemsSource = None
        combobox.SelectedIndex = -1
        combobox.Text = ''
        combobox.ItemsSource = itemssource
        itemssource = None
        del itemssource
    
    @staticmethod
    def ClearValue_TextBox(textbox,Text = ''):
        textbox.Text = Text
    
    @staticmethod
    def ClearValue_CheckBox(checkbox,checked = False):
        checkbox.IsChecked = checked

    @staticmethod
    def wert_schreiben_ComboBox(param,Combobox):
        if param:
            if Combobox.SelectedIndex != -1:
                wert = Combobox.SelectedItem
                wert_schreibenbase(param, wert)

    @staticmethod
    def wert_schreiben_Textbox(param,Textbox):
        if param:
            wert = Textbox.Text
            wert_schreibenbase(param, wert)

    @staticmethod
    def wert_schreiben_Checkbox(param,Checkbox):
        if param:
            wert = Checkbox.IsChecked
            if wert:
                wert = 1
            else:
                wert = 0
            wert_schreibenbase(param, wert)


