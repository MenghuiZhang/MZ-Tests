# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
import math



class VERBINDEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        while (True):
            try:

                ref0 = uidoc.Selection.PickObject(Selection.ObjectType.Element)
                ref1 = uidoc.Selection.PickObject(Selection.ObjectType.Element)
                elem0= doc.GetElement(ref0)
                try:type0 = elem0.Symbol.Id
                except:type0 = elem0.GetTypeId()
                elem1= doc.GetElement(ref1)
                try:type1 = elem1.Symbol.Id
                except:type1 = elem1.GetTypeId()
                t = DB.Transaction(doc,'Test')
                t.Start()
                try:
                    try:mark = elem0.LookupParameter('ZS_MARK').AsString()
                    except:pass
                    elem0.ChangeTypeId(type1)
                    elem1.ChangeTypeId(type0)
                    try:elem1.LookupParameter('ZS_MARK').Set(mark)
                    except:pass
                except Exception as e:
                    print(e)
                t.Commit()
                t.Dispose()
            except Exception as e:
                print(e)
                break

    def GetName(self):
        return "Verbinden"

