# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from IGF_Filters._selectionfilter import PickElementOptionFactory


class Verbinden(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        elem0 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc,lambda x:x.Category.Id.ToString() == '-2003600')
        elem1 = PickElementOptionFactory.CreateCurrentDocumentOption(uidoc,lambda x:x.Category.Id.ToString() == '-2003600')
        t = DB.Transaction(doc,'1')
        t.Start()
        elem1.Number = elem0.Number
        t.Commit()
        

    def GetName(self):
        return "connect"

