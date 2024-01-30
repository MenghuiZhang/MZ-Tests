# coding: utf8
from IGF_Filters._selectionfilter import PickElementOptionFactory
from IGF_ExternalEventHandler import ItemTemplateEventHandler,ExternalEvent
import Autodesk.Revit.DB as DB

class ANZEIGEN(ItemTemplateEventHandler):
    def __init__(self):
        ItemTemplateEventHandler.__init__(self)
        self.MEP = None
        self.ExcuteApp = self.Nummerieren
        self.name = 'Nummer'
    
    def Nummerieren(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        nummer = 0
        while(True):
            auslass = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc,lambda x:(x.Category.Name == "Luftdurchlässe" and x.Name.find('24h') == -1))
            mepraum = auslass.Space[doc.GetElement(auslass.CreatedPhaseId)]
            t = DB.Transaction(doc,'1')
            if mepraum:

                t.Start()
                if self.MEP == None or self.MEP != mepraum.Number:
                    self.MEP = mepraum.Number
                    nummer = 0
                auslass.LookupParameter('IGF_X_Bauteil_ID_Text').Set('LAB_' + self.MEP + '_'+str(nummer))
                nummer+=1
                t.Commit()
                t.Dispose()
            else:
                return




    