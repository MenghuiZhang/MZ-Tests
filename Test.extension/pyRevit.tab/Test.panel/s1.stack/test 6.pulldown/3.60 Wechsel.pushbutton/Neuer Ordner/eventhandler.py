# coding: utf8
import math
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit
import os
from System.Collections.Generic import List


class FamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self,familyInUse, overwriteParameterValues = False): return True
    def OnSharedFamilyFound(self,familyInUse, source, overwriteParameterValues = False): return True

class Familien(object):
    def __init__(self,elem,name):
        self.elem = elem
        self.name = name

def Get_IS():
    _dict = {}
    Liste = []
    FamilySymbols = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).WhereElementIsNotElementType()
    for el in FamilySymbols:
        if el.FamilyCategory.Name == 'Luftkanalzubehör':      
            _dict[el.Name] = el

    return _dict

LISTE_IS = Get_IS()

class CHANGEFAMILY(IExternalEventHandler):
    def __init__(self):
        self.GUI = None

        
    def Execute(self,app):
        if self.GUI.alt.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Bitte alte BSK-Familie auswählen!')
            return  
        if self.GUI.neu.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Bitte neue BSK-Familie auswählen!')
            return      

        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Auswahl = uidoc.Selection.GetElementIds()
        Liste = []
        for el in Auswahl:
            elem = doc.GetElement(el)
            try:
                if elem.Symbol.FamilyName == self.GUI.alt.SelectedItem.ToString():
                    Liste.append(elem)
            except:pass

        family = self.GUI.liste_is[self.GUI.neu.SelectedItem.ToString()]
        typs = {}
        for Id in family.GetFamilySymbolIds():
            try:
                type = doc.GetElement(Id)
                b = type.LookupParameter('MC Connection Size 1').AsValueString()
                h = type.LookupParameter('MC Connection Size 2').AsValueString()
                typs[Id] = [b,h]
            except Exception as e:print(e)
        
        if len(Liste) == 0:
            TaskDialog.Show('Info.','Keine BSK gefunden!')
            return

        t = DB.Transaction(doc,'BSK-Typ wechseln')
        t.Start()
        for n,item in enumerate(self.elems):
            b = item.LookupParameter('Breite').AsValueString()
            h = item.LookupParameter('Höhe').AsValueString()
            rotate = False
            id_neu = None
            for Id in typs.keys():
                b0 = typs[Id][0]
                h0 = typs[Id][1]
                if b0 == b and h0 == h:
                    id_neu = Id
                    break
                elif b0 == h and h0 == b:
                    id_neu = Id
                    rotate = True
                    break
            if id_neu == None:
                print(item.Id.ToString(),'  B:{}, H:{}, Kein Typ gefunden.'.format(b,h))
                continue
            if not rotate:
                try:
                    item.ChangeTypeId(id_neu)
                except:
                    try:
                        conns = list(item.MEPModel.ConnectorManager.Connectors)
                        elem0 = None
                        elem1 = None
                        conn0 = conns[0]
                        conn1 = conns[1]
                        refs0 = conn0.AllRefs
                        refs1 = conn1.AllRefs
                        for ref in refs0:
                            owner = ref.Owner
                            if owner.Category.Name != 'Luftkanäle':
                                continue
                            else:elem0 = owner
                            try:conns_new = owner.ConnectorManager.Connectors
                            except:continue
                            
                            for conn in conns_new:
                                if conn.IsConnectedTo(conn0):
                                    conn.DisconnectFrom(conn0)
                                    doc.Regenerate()
                                    
                                    break
                        for ref in refs1:
                            owner = ref.Owner
                            if owner.Category.Name != 'Luftkanäle':
                                continue
                            else:elem1 = owner
                            try:conns_new = owner.ConnectorManager.Connectors
                            except:continue
                    
                            for conn in conns_new:
                                if conn.IsConnectedTo(conn1):
                                    conn.DisconnectFrom(conn1)
                                    doc.Regenerate()
                                    break
                        item.ChangeTypeId(id_neu)
                        print("BSK {} ist nicht mit dem Hauptnetz verbunden".format(item.Id.toString()))
                    except Exception as e:print(e)
            else:
                conns = list(item.MEPModel.ConnectorManager.Connectors)
                elem0 = None
                elem1 = None
                conn0 = conns[0]
                conn1 = conns[1]
                refs0 = conn0.AllRefs
                refs1 = conn1.AllRefs
                for ref in refs0:
                    owner = ref.Owner
                    if owner.Category.Name != 'Luftkanäle':
                        continue
                    else:elem0 = owner
                    try:conns_new = owner.ConnectorManager.Connectors
                    except:continue
                    
                    for conn in conns_new:
                        if conn.IsConnectedTo(conn0):
                            conn.DisconnectFrom(conn0)
                            doc.Regenerate()
                            
                            break
                for ref in refs1:
                    owner = ref.Owner
                    if owner.Category.Name != 'Luftkanäle':
                        continue
                    else:elem1 = owner
                    try:conns_new = owner.ConnectorManager.Connectors
                    except:continue
                    
                    for conn in conns_new:
                        if conn.IsConnectedTo(conn1):
                            conn.DisconnectFrom(conn1)
                            doc.Regenerate()
                            break
                if not (elem0 and elem1):
                    print('Fehler mit BSK {} da keine Verbindung mit Luftkanäle'.format(item.Id.ToString()))
                    continue
                try:
                    item.ChangeTypeId(id_neu)
                    doc.Regenerate()
                    conns_new2 = list(item.MEPModel.ConnectorManager.Connectors)
                    line = DB.Line.CreateBound(conns_new2[0].Origin,conns_new2[1].Origin)
                    DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([item.Id]), line,0.5*math.pi)
                    distance = 100
                    conns1 = elem1.ConnectorManager.Connectors
                    c1 = ''
                    c2 = ''
                    for co1 in conns_new2:
                        for co2 in conns1:
                            distan = co1.Origin.DistanceTo(co2.Origin)
                            if distan <= distance:
                                distance = distan
                                c1 = co1
                                c2 = co2
                    try:
                        c1.ConnectTo(c2)
                        c1.Origin = c2.Origin
                    except Exception as e: print(e)

                    distance = 100
                    conns2 = elem0.ConnectorManager.Connectors
                    c1 = ''
                    c2 = ''
                    for co1 in conns_new2:
                        for co2 in conns2:
                            distan = co1.Origin.DistanceTo(co2.Origin)
                            if distan <= distance:
                                distance = distan
                                c1 = co1
                                c2 = co2
                    try:
                        c1.ConnectTo(c2)
                        c1.Origin = c2.Origin
                    except Exception as e: print(e)
                except Exception as e:print(e)

        t.Commit()
        t.Dispose()
        TaskDialog.Show('Fertig','Erledigt!')

    def GetName(self):
        return "BSK austauschen"