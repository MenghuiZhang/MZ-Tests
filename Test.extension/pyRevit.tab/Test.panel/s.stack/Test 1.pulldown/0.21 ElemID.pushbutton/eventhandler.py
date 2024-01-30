# coding: utf8
from IGF_ExternalEventHandler import TaskDialog,ItemTemplateEventHandler,ExternalEvent
from IGF_ExternalEventHandler._ExternalEventHandler_ReadOnly import ExternalEventHandlerFactory_READONLY

class ExternalEvenetListe(ItemTemplateEventHandler):
    def __init__(self):
        ItemTemplateEventHandler.__init__(self)

    def anzeigen(self,uiapp):
        self.name = 'Element anzeigen'
        try:elemid = self.GUI.elemetid.SelectedItem
        except:elemid = -1
        ExternalEventHandlerFactory_READONLY.ElementAnzeigenElementId(uiapp,elemid)
    
    def select(self,uiapp):
        self.name = 'Element auswählen'
        try:elemid = self.GUI.elemetid.SelectedItem
        except:elemid = -1
        ExternalEventHandlerFactory_READONLY.ElementAuswaehlenElementId(uiapp,elemid)
