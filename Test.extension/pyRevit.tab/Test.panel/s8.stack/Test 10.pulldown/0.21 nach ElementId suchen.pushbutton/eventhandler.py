# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class ExternalEvenetListe(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.ExcuteApp = None
        self.Name = ''
    
    def Execute(self,uiapp):
        self.Name = 'Material'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        t = DB.Transaction(doc,self.GUI.material)
        t.Start()
        for el in uidoc.Selection.GetElementIds():
            elem = doc.GetElement(el)
            if elem.Category.Name in ['Luftkanäle','Luftkanalformteile']:
                try:
                    elem.LookupParameter('IGF_X_Material_Text').Set(self.GUI.material)
                except:
                    pass
            try:
                elem.LookupParameter('IGF_X_Gewerkkürzel_Exemplar').Set('R')
            except:
                pass
        t.Commit()
        t.Dispose()
        del t
            
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
    def GetName(self):
        return self.Name
    
