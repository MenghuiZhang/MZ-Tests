# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List


class VERBINDEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        cl = uidoc.Selection.GetElementIds()
    
        dict_hls = {}
        dict_zu = {}
        for elid in cl:
            el =doc.GetElement(elid)
            if el.Category.Name == 'HLS-Bauteile':
                if el.Symbol.FamilyName == 'MC_ConnectionNode_Hydronic':
                    dict_hls[el] = el.Location.Point
            elif el.Category.Name == 'Rohrzubehör':
                if el.Symbol.FamilyName.find('6-Wege'):
                    Id = el.LookupParameter('SBI_Bauteilnummerierung').AsString()
                    if not Id:
                        continue
                    else:
                        if Id.find('6WV')!= -1:     dict_zu[el.Location.Point] = Id
        if len(dict_hls) != len(dict_zu):
            print('Anzahl passt nicht')
            return
        neu = {}
        for el0 in dict_hls.keys():
            dis = 100
            for el1 in dict_zu.keys():
                di = el1.DistanceTo(dict_hls[el0])
                if di < dis:
                    dis = di
                    neu[el0] = dict_zu[el1]
        
        test0 = neu.values()[:]
        test1 =set(test0)
        test2 = list(test1)
        if len(test1) != len(test2):
            print('elem bisitzen gleiche Id')
            return
        t = DB.Transaction(doc,'Segel')
        t.Start()
        for el in neu.keys():
            el.LookupParameter('SBI_Bauteilnummerierung').Set('DBH_'+neu[el][4:])
            el.ChangeTypeId(DB.ElementId(37932638))  
        t.Commit()
        print('erledigt!',len(neu))


    def GetName(self):
        return "Verbinden"