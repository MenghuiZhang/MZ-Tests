# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB


class Verbinden(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = [doc.GetElement(el) for el in uidoc.Selection.GetElementIds()]
        conns = []
        for el in cl:
            print(el)
            cos = el.ConnectorManager.Connectors
            for co in cos:
                if not co.IsConnected:
                    conns.append(co)
                    
        print(conns)
        t = DB.Transaction(doc,'connect')
        t.Start()

        conns[0].ConnectTo(conns[1])
        t.Commit()

    def GetName(self):
        return "connect"

