# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection,RevitCommandId,PostableCommand
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from pyrevit import script
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Windows import Visibility
from System.Windows.Input import Key
import System
import os
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as ex
from System.Runtime.InteropServices import Marshal


Visible = Visibility.Visible
Hidden = Visibility.Hidden

doc = __revit__.ActiveUIDocument.Document
projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number
config = script.get_config('Schema-TS-Param -' + projectinfo)

_hls = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
_params = []
for el in _hls:
    params = el.Parameters
    for pa in params:
        try:
            name = pa.Definition.Name
            if name not in _params:_params.append(name)
        except:
            pass
    break
_params.sort()

class Rohr:
    def __init__(self,elem,nr):
        self.nr = nr
        self.elem = elem
    def wert_schreiben(self):
        self.elem.LookupParameter('IGF_X_RVT_TS_Nr').Set(str(self.nr))

class FilterNeben(Selection.ISelectionFilter):
    def __init__(self,liste):
        self.Liste = liste

    def AllowElement(self,element):
        if element.Id.ToString() in self.Liste:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter(Selection.ISelectionFilter):

    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class AlleBauteile:
    def __init__(self,Liste,elem,param):
        self.elem = elem
        self.param = param
        self.rohrliste = Liste
        self.liste = []
        self.hls_liste = {}
        self.rohr = []
        self.formteil = []
        self.enddeckel = []
        self.zubehoer = []
        for el in self.rohrliste:
            self.liste.append(el)
            self.rohr.append(el)
        if self.elem.Category.Id.ToString() == '-2008044':
            self.rohr.append(self.elem.Id.ToString())
        elif self.elem.Category.Id.ToString() == '-2008049':
            self.formteil.append(self.elem.Id.ToString())
        elif self.elem.Category.Id.ToString() == '-2008055':
            self.zubehoer.append(self.elem.Id.ToString())  
        self.get_t_st(self.elem)
        self.alle = []
        self.alle.extend(self.rohr)
        self.alle.extend(self.zubehoer)
        self.alle.extend(self.formteil)
        self.alle.extend(self.enddeckel)

    def get_t_st(self,elem):       
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste:  
                            category_temp = owner.Category.Id.ToString()
                            if category_temp == '-2001140':
                                try:
                                    nr = owner.LookupParameter(self.param).AsString() 
                                    if not nr:
                                        print('Element {} hat kein Bauteilid'.format(owner.Id.ToString()))
                                    if nr in self.hls_liste.keys():
                                        print('Element {} besitzt doppelte Bauteilid {}'.format(owner.Id.ToString(),nr))
                                except:
                                    print(owner.Id)
                                if nr:self.hls_liste[nr] = owner
                                return
                            elif category_temp == '-2008044':
                                self.rohr.append(owner.Id.ToString())
                            elif category_temp == '-2008049':
                                conn_temp = owner.MEPModel.ConnectorManager.Connectors
                                if conn_temp.Size == 1:self.enddeckel.append(owner.Id.ToString())
                                else:self.formteil.append(owner.Id.ToString())
                            elif category_temp == '-2008055':
                                self.zubehoer.append(owner.Id.ToString())                 
                            self.get_t_st(owner)

class HLS:
    def __init__(self,elem,nr,vorhanden):
        self.elem = elem
        self.vorhanden = vorhanden
        self.nr = nr
        self.wege_3 = ''
        self.liste = []
        self.Liste_rohr = []
        self.Liste_rohrid = []
        self.Liste_rohrfid = []
        self.Liste_rohrzid = []
        self.temp_liste = [] 
        self.alle = []
        self.get_liste(self.elem)
        self.alle.extend(self.Liste_rohrid)
        self.alle.extend(self.Liste_rohrfid)
        self.alle.extend(self.Liste_rohrzid)

    def get_liste(self,elem):
        if self.wege_3:return
        
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cateid = elem.Category.Id.ToString()
        
        if cateid not in ['-2008043','-2008122']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and cateid != '-2001140':
                    self.wege_3 = elemid
                    self.Liste_rohr.extend(self.temp_liste)
                    return
                elif cateid == '-2008049':
                    conns_liste = []
                    for conn in conns:
                        _dn = int(conn.Radius*2*304.8)
                        if _dn not in conns_liste:conns_liste.append(_dn)
                    if len(self.temp_liste) > 0 and len(conns_liste) > 1:
                        self.nr += 0.1
                        self.Liste_rohr.extend(self.temp_liste)
                        self.temp_liste = []

                for conn in conns:
                    refs = conn.AllRefs
                    if elemid != self.elem.Id.ToString():
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() not in self.liste and owner.Id.ToString() in self.vorhanden:  
                                if owner.Category.Id.ToString() == '-2008044':  
                                    self.temp_liste.append(Rohr(owner,self.nr))  
                                    self.Liste_rohrid.append(owner.Id.ToString())
                                elif owner.Category.Id.ToString() == '-2008055':
                                    self.Liste_rohrzid.append(owner.Id.ToString())
                                    if len(self.temp_liste) >0:
                                        self.nr += 0.1   
                                        self.Liste_rohr.extend(self.temp_liste) 
                                        self.temp_liste = []
                                elif owner.Category.Id.ToString() == '-2008049':
                                    self.Liste_rohrfid.append(owner.Id.ToString())
                                self.get_liste(owner)
                    else:
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() in self.vorhanden: 
                                if owner.Category.Id.ToString() == '-2008044':
                                    self.temp_liste.append(Rohr(owner,self.nr))
                                    self.Liste_rohrid.append(owner.Id.ToString())
                                self.get_liste(owner)
                            else:
                                continue

class Rohr1:
    def __init__(self,elemid,doc):
        self.elemid = DB.ElementId(int(elemid))
        self.doc = doc
        self.elem = self.doc.GetElement(self.elemid)
        self.durchmesser = self.elem.LookupParameter('IGF_X_SM_Durchmesser').AsValueString()
    def wert_schreiben(self):
        self.elem.LookupParameter('Durchmesser').SetValueString(self.durchmesser)

class Rohrformteil:
    def __init__(self,elem):
        self.doc = None
        self.elem = elem
        self.conns0 = ''
        self.conns1 = ''
        self.conns2 = ''
        self.liste = []
        self.art = ''
        self.get_conns()
    
    def get_conns(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        if conns.Size == 2:
            conns = list(conns)
            conn_sizes = [int(conn.Radius*304.8*2) for conn in conns]
            if conn_sizes[0] != conn_sizes[1]:
                self.art = 'Übergang'
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Category.Id.ToString() in ['-2008044','-2008055','-2001140','-2001160']: # Rohre, Zubehör, HLS, Sanitärinstallationen
                                try:
                                    conns_owner = owner.ConnectorManager.Connectors
                                except:
                                    conns_owner = owner.MEPModel.ConnectorManager.Connectors
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp):
                                        if self.conns0 == '':
                                            self.conns0 = conn_temp
                                        else:
                                            self.conns1 = conn_temp
            else:
                self.art = 'Bogen'
                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            elemid = owner.Id.ToString()
                            if elemid in self.liste:
                                continue
                            self.liste.append(elemid)
                            if owner.Category.Id.ToString() == '-2008044':
                                conns_owner = owner.ConnectorManager.Connectors
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp):
                                        if self.conns0 == '':
                                            self.conns0 = conn_temp
                                        else:
                                            self.conns1 = conn_temp
                            elif owner.Category.Id.ToString() == '-2008049':
                                conns_owner = owner.MEPModel.ConnectorManager.Connectors
                                conn_ander = ''
                                for conn_temp in conns_owner:
                                    if conn.IsConnectedTo(conn_temp) == False:
                                        conn_ander = conn_temp
                                        break 
                                allrefs2 = conn_ander.AllRefs
                                for ref2 in allrefs2:
                                    owner2 = ref2.Owner
                                    if owner2.Category.Id.ToString() == '-2008044':
                                        conns_owner2 = owner2.ConnectorManager.Connectors
                                        for conn_temp2 in conns_owner2:
                                            if conn_ander.IsConnectedTo(conn_temp2):
                                                if self.conns0 == '':
                                                    self.conns0 = conn_temp2
                                                else:
                                                    self.conns1 = conn_temp2
        else:
            self.art = 'T-Stück'
            for conn in conns:
                if conn.IsConnected:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        elemid = owner.Id.ToString()
                        if elemid in self.liste:
                            continue
                        self.liste.append(elemid)
                        if owner.Category.Id.ToString() == '-2008044':
                            conns_owner = owner.ConnectorManager.Connectors
                            for conn_temp in conns_owner:
                                if conn.IsConnectedTo(conn_temp):
                                    if self.conns0 == '':
                                        self.conns0 = conn_temp
                                    elif self.conns1 == '':
                                        self.conns1 = conn_temp
                                    else:
                                        self.conns2 = conn_temp
                        elif owner.Category.Id.ToString() == '-2008049':
                            conns_owner = owner.MEPModel.ConnectorManager.Connectors
                            conn_ander = ''
                            for conn_temp in conns_owner:
                                if conn.IsConnectedTo(conn_temp) == False:
                                    conn_ander = conn_temp
                                    break 
                            allrefs2 = conn_ander.AllRefs
                            for ref2 in allrefs2:
                                owner2 = ref2.Owner
                                if owner2.Category.Id.ToString() == '-2008044':
                                    conns_owner2 = owner2.ConnectorManager.Connectors
                                    for conn_temp2 in conns_owner2:
                                        if conn_ander.IsConnectedTo(conn_temp2):
                                            if self.conns0 == '':
                                                self.conns0 = conn_temp2
                                            elif self.conns1 == '':
                                                self.conns1 = conn_temp2
                                            else:
                                                self.conns2 = conn_temp2
        

    def createfittings(self):
        if self.art == 'Bogen':
            try:self.doc.Create.NewElbowFitting(self.conns0, self.conns1)
            except:self.doc.Create.NewElbowFitting(self.conns1, self.conns0)
        elif self.art == 'Übergang':
            try:self.doc.Create.NewTransitionFitting(self.conns0, self.conns1)
            except:self.doc.Create.NewTransitionFitting(self.conns1, self.conns0)
        elif self.art == 'T-Stück':
            try:self.doc.Create.NewTeeFitting(self.conns0, self.conns1,self.conns2)
            except:
                try:self.doc.Create.NewTeeFitting(self.conns2, self.conns0, self.conns1)
                except:self.doc.Create.NewTeeFitting(self.conns1, self.conns2, self.conns0)

class WEGE3:
    def __init__(self,Liste_Rohre,elemid,nr,nicht_bearbeitet,angepasst,doc,vorhanden):
        self.doc = doc
        self.Eingang_Liste_Rohre = Liste_Rohre
        self.angepasst = angepasst
        self.nicht_bearbeitet = nicht_bearbeitet
        self.nr = nr
        self.elemid = elemid
        self.vorhanden = vorhanden
        self.elem = self.doc.GetElement(DB.ElementId(int(self.elemid)))

        self.liste = []
        for el in self.Eingang_Liste_Rohre:
            self.liste.append(el)

        self.wege_3_0 = ''
        self.wege_3_1 = ''
        self.Liste_rohr_0 = []
        self.Liste_rohr_1 = []
        self.Liste_rohrid_0 = []
        self.Liste_rohrid_1 = []
        self.Liste_rohrfid_0 = []
        self.Liste_rohrfid_1 = []
        self.Liste_rohrzid_0 = []
        self.Liste_rohrzid_1 = []
        self.temp_liste = []
        self.alle_0 =[]
        self.alle_1 =[]
        self.t_dict = {}

        self.get_liste(self.elem)
        self.korrigieren()
        
    def korrigieren(self):
        if self.wege_3_0 in self.angepasst:
            self.Liste_rohr_0 = []
            self.Liste_rohrid_0 = []
            self.wege_3_0 = 'Angepasst'
        
        if self.wege_3_1 in self.angepasst:
            self.Liste_rohr_1 = []
            self.Liste_rohrid_1 = []
            self.wege_3_1 = 'Angepasst'
        
        if len(self.Liste_rohrid_0) == 0:
            if len(self.Liste_rohrid_1) > 0:
                for el in self.Liste_rohr_1:
                    el.nr = el.nr - 1
        
        if self.wege_3_0 == '' or self.wege_3_0 == 'Angepasst':pass
        elif self.wege_3_0 == 'Enddeckel':
            if self.wege_3_1 in ['','Angepasst']:
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1

        elif self.wege_3_0 == 'Offen':
            if self.wege_3_1 in ['','Angepasst','Enddeckel']:
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
        elif self.wege_3_0 == 'HLS-Bauteile':
            if self.wege_3_1 in ['','Angepasst','Enddeckel','Offen']:
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
        else:   
            if self.wege_3_1 in ['','Angepasst','Enddeckel','Offen','HLS-Bauteile']:
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
            elif self.wege_3_0 in self.nicht_bearbeitet and self.wege_3_1 in self.nicht_bearbeitet:
                index1 = self.nicht_bearbeitet.index(self.wege_3_0)
                index2 = self.nicht_bearbeitet.index(self.wege_3_1)
                if index2 > index1:
                    if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                        for el in self.Liste_rohr_0:
                            el.nr += 1
                        for el in self.Liste_rohr_1:
                            el.nr -= 1

            elif self.wege_3_0 not in self.nicht_bearbeitet and self.wege_3_1 in self.nicht_bearbeitet:
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
            elif self.wege_3_0 in self.nicht_bearbeitet and self.wege_3_1 not in self.nicht_bearbeitet:
                pass
            else:
                self.Liste_rohr_0 = []
                self.Liste_rohr_1 = []
                self.Liste_rohrid_0 = []
                self.Liste_rohrid_1 = []
                
                self.wege_3_0 = self.elemid
                self.wege_3_1 = 'Erneuen'
        self.alle_0.extend(self.Liste_rohrid_0)
        self.alle_0.extend(self.Liste_rohrfid_0)
        self.alle_0.extend(self.Liste_rohrzid_0)
        self.alle_1.extend(self.Liste_rohrid_1)
        self.alle_1.extend(self.Liste_rohrfid_1)
        self.alle_1.extend(self.Liste_rohrzid_1)
    def get_liste(self,elem):        
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cateid = elem.Category.Id.ToString()

        if cateid not in ['-2008043','-2008122']: # entspricht entweder RoheSystem oder Rohrdämmung
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size == 1: # Enddeckel oder Entlerrungsventile
                    if self.wege_3_0 == '':
                        self.wege_3_0 = 'Enddeckel'
                        self.nr = int(self.nr) + 1
                        if len(self.temp_liste) != 0: 
                            self.Liste_rohr_0.extend(self.temp_liste)
                            self.temp_liste = []
                    elif self.wege_3_1 == '':
                        self.wege_3_1 = 'Enddeckel'
                        if len(self.temp_liste) != 0: 
                            self.Liste_rohr_1.extend(self.temp_liste)
                            self.temp_liste = []

                    return
                
                elif conns.Size == 2: # Übergang
                    if cateid == '-2008049':
                        conns_liste = []
                        for conn in conns:
                            _dn = int(conn.Radius*2*304.8)
                            if _dn not in conns_liste:conns_liste.append(_dn)
                        if len(conns_liste) > 1:
                            if len(self.temp_liste) > 0:
                                if self.wege_3_0 == '':
                                    self.Liste_rohr_0.extend(self.temp_liste)
                                else:
                                    self.Liste_rohr_1.extend(self.temp_liste)
                                self.nr += 0.1  
                                self.temp_liste = []

                elif conns.Size > 2: # T-Stück
                    if elemid != self.elemid:   
                        if self.wege_3_0 == '':
                            self.wege_3_0 = elemid
                            self.nr = int(self.nr) + 1
                            if len(self.temp_liste) > 0: 
                                self.Liste_rohr_0.extend(self.temp_liste)
                                self.temp_liste = []
                        else:
                            self.wege_3_1 = elemid
                            if len(self.temp_liste) > 0: 
                                self.Liste_rohr_1.extend(self.temp_liste)
                                self.temp_liste = []

                        return

                for conn in conns:
                    if conn.IsConnected == False:
                        if self.wege_3_0 == '':
                            self.wege_3_0 = 'Offen'
                            self.nr = int(self.nr) + 1
                            if len(self.temp_liste) > 0:
                                self.Liste_rohr_0.extend(self.temp_liste)
                                self.temp_liste = []
                            
                        elif self.wege_3_1 == '':
                            self.wege_3_1 = 'Offen'
                            if len(self.temp_liste) > 0:
                                self.Liste_rohr_1.extend(self.temp_liste)
                                self.temp_liste = []
                        return
                            
                    else:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            ownerid = owner.Id.ToString()
                            ownercateid = owner.Category.Id.ToString()
                            if ownercateid not in ['-2008043','-2008122']:
                                if ownerid not in self.liste:
                                    if ownerid in self.vorhanden:  
                                        if ownercateid == '-2008044': # Rohre
                                            self.temp_liste.append(Rohr(owner,self.nr))  
                                            if self.wege_3_0 == '':self.Liste_rohrid_0.append(ownerid)
                                            else:self.Liste_rohrid_1.append(ownerid)
                                        
                                        elif ownercateid == '-2008055': # Rohrzubehör
                                            if self.wege_3_0 == '':self.Liste_rohrzid_0.append(ownerid)
                                            else:self.Liste_rohrzid_1.append(ownerid)
                                            if len(self.temp_liste) >0:
                                                self.nr += 0.1
                                                if self.wege_3_0 == '':
                                                    self.Liste_rohr_0.extend(self.temp_liste)
                                                    self.temp_liste = []
                                                else:
                                                    self.Liste_rohr_1.extend(self.temp_liste)
                                                    self.temp_liste = []
                                        
                                        elif ownercateid == '-2008049': # Rohrformteile
                                            if self.wege_3_0 == '':self.Liste_rohrfid_0.append(ownerid)
                                            else:self.Liste_rohrfid_1.append(ownerid)
                                                    
                                        elif ownercateid == '-2001140': # HLS-Bauteile
                                            if self.wege_3_0 == '': 
                                                self.wege_3_0 = 'HLS-Bauteile'
                                                self.nr = int(self.nr) + 1
                                                if len(self.temp_liste) > 0:
                                                    self.Liste_rohr_0.extend(self.temp_liste)
                                            else:
                                                self.wege_3_1 = 'HLS-Bauteile'
                                                if len(self.temp_liste) > 0:
                                                    self.Liste_rohr_1.extend(self.temp_liste)
                                            return                                     
                                        
                                        self.get_liste(owner)
                                    else:
                                        if self.wege_3_0 == '':
                                            self.wege_3_0 = 'Offen'
                                            self.nr = int(self.nr) + 1
                                            if len(self.temp_liste) > 0:
                                                self.Liste_rohr_0.extend(self.temp_liste)
                                                self.temp_liste = []
                                        elif self.wege_3_1 == '':
                                            self.wege_3_1 = 'Offen'
                                            if len(self.temp_liste) > 0:
                                                self.Liste_rohr_1.extend(self.temp_liste)
                                                self.temp_liste = []
                            
    def get_Liste_T_Piece(self,elem):
        if (self.T_Piece_0 or self.Endverbraucher_0) and (self.T_Piece_1 or self.Endverbraucher_1) :return
        
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and cate == 'Rohrformteile' and elemid != self.elemid.ToString():
                    if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                        self.T_Piece_0 = elem.Id.ToString()
                    else:
                        self.T_Piece_1 = elem.Id.ToString()
                    return
                if conns.Size == 1:
                    if self.Endverbraucher_0 == None:
                        self.Endverbraucher_0 = 'Pass'
                    else:
                        self.Endverbraucher_1 = 'Pass'

                for conn in conns:
                    if not conn.IsConnected:
                        if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                            self.Endverbraucher_0 = 'Pass'
                        else:self.Endverbraucher_1 = 'Pass'
                        continue
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste and owner.Id.ToString() not in self.Liste_Pre_Rohre:  
                            if owner.Category.Id.ToString() == '-2008044':
                                if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                                    self.Liste_Rohre_1.append(owner.Id.ToString())
                                else:
                                    self.Liste_Rohre_2.append(owner.Id.ToString())
                            elif owner.Category.Id.ToString() in ['-2008099','-2001160','-2001140']:
                                if self.T_Piece_0 == None and self.Endverbraucher_0 == None:
                                    self.Endverbraucher_0 = owner.Id.ToString()
                                else:self.Endverbraucher_1 = owner.Id.ToString()
                                return
                           
                            self.get_Liste_T_Piece(owner)

    def get_Liste_Verbraucher(self):
        if self.T_Piece_0 in self.t_dict.keys():
            if self.t_dict[self.T_Piece_0].angepasst == False:
                self.logger.error(self.t_dict[self.T_Piece_0].elemid.ToString())
            self.Anzahl_Endverbraucher += self.t_dict[self.T_Piece_0].Anzahl_Endverbraucher
        if self.T_Piece_1 in self.t_dict.keys():
            if self.t_dict[self.T_Piece_1].angepasst == False:
                self.logger.error(self.t_dict[self.T_Piece_1].elemid.ToString())
            self.Anzahl_Endverbraucher += self.t_dict[self.T_Piece_1].Anzahl_Endverbraucher
 
    def get_Dimension(self):
        for n in self.List_dimension:
            if self.Anzahl_Endverbraucher >= n:
                self.Dimension = self.dict_dimension_Neu[n]
                break

    def wert_schreiben(self):
        if self.Endverbraucher_0 != None:
            for rohr in self.Liste_Rohre_1:
                try:
                    self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.grunddimension))
                except Exception as e:
                    self.logger.error(e)
        if self.Endverbraucher_1 != None:
            for rohr in self.Liste_Rohre_2:
                try:
                    self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.grunddimension))
                except Exception as e:
                    self.logger.error(e)
        
        for rohr in self.Liste_Pre_Rohre:
            try:
                self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.Dimension))
            except Exception as e:
                self.logger.error(e)
            try:
                self.doc.GetElement(DB.ElementId(int(rohr))).LookupParameter('IGF_X_Anzahl_Sprinkler').Set(self.Anzahl_Endverbraucher)
            except Exception as e:
                self.logger.error(e)
                                          
class NUMMERIEREN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        try:
            self.GUI.write_config()
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document

            Rohr_liste = []
            Rohrid_liste = []
            gesamt_liste = []
            T_Stueck_Liste = []
            T_Stueck_Dict = {}

            Ids = self.GUI.allebauteile.hls_liste.keys()[:]  
            Ids.sort()
            print('Insgesamt {} HLS-Bauteile gefunden. Folgende werden die Bauteilids aufgelistet.'.format(len(Ids)))
            print(Ids)
            tsnr = int(self.GUI.startnr.Text) - 1

            for Id in Ids:
                tsnr += 1
                el = self.GUI.allebauteile.hls_liste[Id]
                hls = HLS(el,tsnr,self.GUI.allebauteile.alle)
                Rohr_liste.extend(hls.Liste_rohr)
                if hls.wege_3 != '':
                    if hls.wege_3 not in T_Stueck_Liste:
                        T_Stueck_Liste.append(hls.wege_3)
                        T_Stueck_Dict[hls.wege_3] = hls.alle
                    else:
                        T_Stueck_Dict[hls.wege_3].extend(hls.alle)
                
                gesamt_liste.extend(hls.liste)
                Rohrid_liste.extend(hls.Liste_rohrid)
                        
            t_stueck_neu = []
            t_stueck_nichtbearbeitet = T_Stueck_Liste[:]
            t_stueck_bearbeitet = []

            for n,wege_3 in enumerate(T_Stueck_Liste):
                tsnr += 1
                wege3 = WEGE3(T_Stueck_Dict[wege_3],wege_3,tsnr,t_stueck_nichtbearbeitet,t_stueck_bearbeitet,\
                    doc,self.GUI.allebauteile.alle)

                liste = [len(wege3.Liste_rohr_0) == 0, len(wege3.Liste_rohr_1) == 0]

                liste_neu = [x for x in liste if x == True]
                if len(liste_neu) == 2:
                    tsnr = tsnr - 1
                elif len(liste_neu) == 1:
                    pass
                elif len(liste_neu) == 0:
                    tsnr = tsnr + 1
                
                Rohr_liste.extend(wege3.Liste_rohr_0)
                Rohr_liste.extend(wege3.Liste_rohr_1)

                if not (wege3.wege_3_0 == wege3.elemid and wege3.wege_3_1 == 'Erneuen'):
                    t_stueck_nichtbearbeitet.remove(wege_3)
                    t_stueck_bearbeitet.append(wege_3)
                    if (wege3.wege_3_0 not in ['','Enddeckel','Angepasst','HLS-Bauteile','Offen']) and (wege3.wege_3_0 not in t_stueck_neu):
                        if wege3.wege_3_0 in T_Stueck_Liste:
                            T_Stueck_Dict[wege3.wege_3_0].extend(wege3.alle_0)
                        else:
                            T_Stueck_Dict[wege3.wege_3_0] = wege3.alle_0
                            t_stueck_neu.append(wege3.wege_3_0)
                    if (wege3.wege_3_1 not in ['','Enddeckel','Angepasst','HLS-Bauteile','Offen']) and (wege3.wege_3_1 not in t_stueck_neu):
                        if wege3.wege_3_1 in T_Stueck_Liste:
                            T_Stueck_Dict[wege3.wege_3_1].extend(wege3.alle_1)
                        else:
                            T_Stueck_Dict[wege3.wege_3_1] = wege3.alle_1
                            t_stueck_neu.append(wege3.wege_3_1)  

                gesamt_liste.extend(wege3.liste)
                Rohrid_liste.extend(wege3.Liste_rohrid_0)
                Rohrid_liste.extend(wege3.Liste_rohrid_1)

            for el in t_stueck_nichtbearbeitet:
                T_Stueck_Liste.remove(el)
            t_stueck_neu.extend(t_stueck_nichtbearbeitet)
            t_stueck_nichtbearbeitet = t_stueck_neu[:]

            _n = 0

            while(len(t_stueck_neu)>0):
                T_Stueck_Liste.extend(t_stueck_neu)
                t_stueck_nichtbearbeitet = t_stueck_neu[:]
                t_stueck_neu_temp = []
                
                for n,wege_3 in enumerate(t_stueck_neu):
                    tsnr = tsnr + 1

                    wege3 = WEGE3(T_Stueck_Dict[wege_3],wege_3,tsnr,t_stueck_nichtbearbeitet,t_stueck_bearbeitet,\
                    doc,self.GUI.allebauteile.alle)

                    Rohr_liste.extend(wege3.Liste_rohr_0)
                    Rohr_liste.extend(wege3.Liste_rohr_1)
                    liste = [len(wege3.Liste_rohr_0) == 0, len(wege3.Liste_rohr_1) == 0]
                    liste_neu = [x for x in liste if x == True]
                    if len(liste_neu) == 2:
                        tsnr = tsnr - 1
                    elif len(liste_neu) == 1:
                        pass
                    elif len(liste_neu) == 0:
                        tsnr = tsnr + 1
                    
                    if not (wege3.wege_3_0 == wege3.elemid and wege3.wege_3_1 == 'Erneuen'):
                        t_stueck_nichtbearbeitet.remove(wege_3)
                        t_stueck_bearbeitet.append(wege_3)
                        if (wege3.wege_3_0 not in ['','Enddeckel','Angepasst','HLS-Bauteile','Offen']) and (wege3.wege_3_0 not in t_stueck_neu_temp): 
                            if wege3.wege_3_0 not in T_Stueck_Liste:
                                T_Stueck_Dict[wege3.wege_3_0] = wege3.alle_0
                                t_stueck_neu_temp.append(wege3.wege_3_0)
                            else:T_Stueck_Dict[wege3.wege_3_0].extend(wege3.alle_0)
                        if (wege3.wege_3_1 not in ['','Enddeckel','Angepasst','HLS-Bauteile','Offen']) and (wege3.wege_3_1 not in t_stueck_neu_temp):
                            if wege3.wege_3_1 not in T_Stueck_Liste:
                                T_Stueck_Dict[wege3.wege_3_1] = wege3.alle_1
                                t_stueck_neu_temp.append(wege3.wege_3_1)  
                            else:T_Stueck_Dict[wege3.wege_3_1] = wege3.alle_1

                    gesamt_liste.extend(wege3.liste)
                    Rohrid_liste.extend(wege3.Liste_rohrid_0)
                    Rohrid_liste.extend(wege3.Liste_rohrid_1)

                t_stueck_neu = t_stueck_neu_temp
                for el in t_stueck_nichtbearbeitet:
                    T_Stueck_Liste.remove(el)
                t_stueck_neu.extend(t_stueck_nichtbearbeitet)

                if len(t_stueck_neu) < 10:
                    _n += 1
                if _n > 100:
                    return
            
            self.GUI.endnr.Text = str(tsnr)

            t = DB.Transaction(doc,'Teilstrecke nummerieren')
            t.Start()
            for n,el in enumerate(Rohr_liste):
                if n % 20  == 0:
                    self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                self.GUI.value = n+1
                self.GUI.maxvalue = len(Rohr_liste)
                self.GUI.PB_text = str(n+1) + ' / '+ str(len(Rohr_liste)) + ' Rohre'
                el.wert_schreiben()

            t.Commit()
            t.Dispose()
            self.GUI.maxvalue = 100
            self.GUI.value = 0
            self.GUI.PB_text = ''
            self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)

        except Exception as e:
            print(e)
    
    def GetName(self):
        return "nummerieren"
    
    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        try:
            self.GUI.write_config()
            uidoc = app.ActiveUIDocument
            doc = uidoc.Document
            el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter(),'Wählt ein Rohr aus')
            el0 = doc.GetElement(el0_ref)
            conns = el0.ConnectorManager.Connectors
            liste = []
            Liste_Rohr = []
            Liste_Rohr.append(el0.Id.ToString())
            for conn in conns:
                if conn.IsConnected:
                    refs = conn.AllRefs
                    for _ref in refs:
                        if _ref.Owner.Category.Name not in ['Rohr Systeme','Rohrdämmung'] and _ref.Owner.Id.ToString() != el0.Id.ToString():
                            liste.append(_ref.Owner.Id.ToString())

            if len(liste) ==  0:
                TaskDialog.Show('Info','kein Teil damit verbunden!')
            ref_1 = uidoc.Selection.PickObject(Selection.ObjectType.Element,FilterNeben(liste),'Wählt den nächsten Teil aus')
            self.GUI.allebauteile = AlleBauteile(Liste_Rohr,doc.GetElement(ref_1),self.GUI.config.param)
            self.GUI.butter_nummer.IsEnabled = True
            self.GUI.butter_infos.IsEnabled = True
            self.GUI.butter_dimension.IsEnabled = True
        except Exception as e:print(e)
    
    def GetName(self):
        return "auswählen"

class DIMENSIONIEREN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        self.GUI.write_config()
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        logger = script.get_logger()
        self.GUI.butter_nummer.IsEnabled = False
        self.GUI.butter_infos.IsEnabled = False
        self.GUI.butter_dimension.IsEnabled = False
        try:
            liste_formteil = []
            formteile = [DB.ElementId(int(elid)) for elid in self.GUI.allebauteile.formteil]
            for elid in self.GUI.allebauteile.formteil:

                el = doc.GetElement(DB.ElementId(int(elid)))
                formteil = Rohrformteil(el)
                if formteil.art:
                    if formteil.art == 'Übergang':
                        if formteil.conns0 and formteil.conns1:
                            liste_formteil.append(formteil)
                    else:
                        liste_formteil.append(formteil)
          
            t = DB.Transaction(doc,'Dimension')
            t.Start()
            
            doc.Delete(List[DB.ElementId](formteile))
            app.ActiveUIDocument.Document.Regenerate()
            for el in self.GUI.allebauteile.rohr:
                try:
                    Rohr1(el,doc).wert_schreiben()
                except Exception as e:
                    logger.error(e)
            app.ActiveUIDocument.Document.Regenerate()
            self.GUI.maxvalue = len(liste_formteil)

            for n,el in enumerate(liste_formteil):
                self.GUI.value = n+1
                self.GUI.PB_text = str(n+1) + ' / '+ str(len(liste_formteil)) + ' Formteile'
                self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                try:
                    el.doc = doc
                    el.createfittings()
                    doc.Regenerate()
                except:pass

            t.Commit()
            t.Dispose()
            self.GUI.maxvalue = 100
            self.GUI.value = 0
            self.GUI.PB_text = ''
            self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
        except Exception as e:
            self.GUI.maxvalue = 100
            self.GUI.value = 0
            self.GUI.PB_text = ''
            self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
            logger.error(e)
    
    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text


    def GetName(self):
        return "Dimension übernehmen"

class UEBERNEHMEN(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        self.GUI.write_config()
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        logger = script.get_logger()
        excel = self.GUI.excel.Text
        if os.path.exists(excel):
            Parameter_Dict = {}
            exapplication = ex.ApplicationClass()
            book1 = exapplication.Workbooks.Open(excel)
            Parameter_Liste = []
            Bauteilnummer = ''
            zukorrigieren = []
            try:
                for sheet in book1.Worksheets:
                    rows = sheet.UsedRange.Rows.Count
                    cols = sheet.UsedRange.Columns.Count
                    Bauteilnummer = sheet.Cells[1, 1].Value2
                    for col in range(2, cols + 1):
                        para = sheet.Cells[1, col].Value2
                        if para == '' or  para == None:
                            break
                        Parameter_Liste.append(para)

                    for row in range(2, rows + 1):
                        liste = []
                        nummer = sheet.Cells[row, 1].Value2
                        if nummer == '' or  nummer == None:
                            continue
                        for col in range(2, len(Parameter_Liste) + 2):
                            value = sheet.Cells[row, col].Value2
                            liste.append(value)
                        
                        # if nummer in Parameter_Dict.keys() and nummer not in zukorrigieren:
                        #     temp_list = Parameter_Dict[nummer]
                            # for n in range(len(temp_list)):
                            #     if temp_list[n] != liste[n]:
                            #         # zukorrigieren.append(nummer)
                            #         # logger.error("TS {} passt nicht.".format(nummer))
                            #         break
                        # else:
                        Parameter_Dict[nummer] = liste
                book1.Save()
                book1.Close()

            except Exception as e:
                logger.error(e)
                Marshal.FinalReleaseComObject(sheet)
                Marshal.FinalReleaseComObject(book1)
                exapplication.Quit()
                Marshal.FinalReleaseComObject(exapplication)
                return
            
            benutzt = []
            rohre = self.GUI.allebauteile.rohr
            self.GUI.maxvalue = len(rohre)
            t = DB.Transaction(doc,'Rohre')
            t.Start()
            for n,elid in enumerate(rohre):
                if n % 10 == 0:
                    self.GUI.value = n+1
                    self.GUI.PB_text = str(n+1) + ' / '+ str(len(rohre)) + ' Rohre'
                    self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                        System.Windows.Threading.DispatcherPriority.Background)
                el = doc.GetElement(DB.ElementId(int(elid)))
                bauteil = el.LookupParameter(Bauteilnummer) 

                if bauteil == None:
                    TaskDialog.Show('Fehler','Parameter für TS in excel ist falsch. Bitte korrigieren!')
                    t.RollBack()
                    return
                Id = bauteil.AsString()
                if Id in Parameter_Dict.keys():
                    if Id not in benutzt:benutzt.append(Id)
                    for n,para in enumerate(Parameter_Liste):
                        param = el.LookupParameter(para)
                        if param.StorageType.ToString() == 'String':
                            param.Set(str(Parameter_Dict[Id][n]))
                        else:
                            param.SetValueString(str(Parameter_Dict[Id][n]))
                elif str(int(float(Id))) in Parameter_Dict.keys():
                    if str(int(float(Id))) not in benutzt:benutzt.append(str(int(float(Id))))
                    for n,para in enumerate(Parameter_Liste):
                        param = el.LookupParameter(para)
                        if param.StorageType.ToString() == 'String':
                            param.Set(str(Parameter_Dict[str(int(float(Id)))][n]))
                        else:
                            param.SetValueString(str(Parameter_Dict[str(int(float(Id)))][n]))
                else:
                    logger.error('TS {} (ElementId {}) existiert nur in Schema-Modell. Bitte einmal prüfen, ob die Bauteilnummer in Schema-Modell richtig eingetragen wird.'.format(Id,el.Id.ToString()))
            t.Commit()

            for el in sorted(Parameter_Dict.keys()):
                if el not in benutzt:
                    logger.error('TS {} ist nicht verwendet'.format(el))
            
            self.GUI.maxvalue = 100
            self.GUI.value = 0
            self.GUI.PB_text = ''
            self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
    
    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text


    def GetName(self):
        return "Infos übernehmen"