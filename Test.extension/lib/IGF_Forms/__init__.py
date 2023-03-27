from pyrevit import forms

class WPFWindow(forms.WPFWindow):
    def __init__(self,wpfgui,handle_esc=True):
        forms.WPFWindow.__init__(self,wpfgui,handle_esc = handle_esc)
        self.Closed += self.ReleaseResources
        self.externalevent = None
        self.externaleventliste = None

    def ReleaseResources(self,sender,e):
        self.externaleventliste = None
        try:
            self.externalevent.Dispose()
        except:
            pass

