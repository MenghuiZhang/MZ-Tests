# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
from rpw import DB,revit,UI
import os
from System.Windows.Input import Key
from System.Collections.Generic import List
import time

__title__ = "3.2 Strang dimensionieren"
__doc__ = """

Parameter: IGF_X_SM_Durchmesser

[2022.08.12]
Version: 1.0

"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc


class Filters(UI.Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class FilterNeben(UI.Selection.ISelectionFilter):
    def __init__(self,liste):
        self.Liste = liste

    def AllowElement(self,element):
        if element.Id.ToString() in self.Liste:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class AlleBauteile:
    def __init__(self,Liste,elem):
        self.elem = elem
        self.rohrliste = Liste
        self.liste = []
        self.rohr = []
        self.formteil = []
        for el in self.rohrliste:
            self.liste.append(el)
            self.rohr.append(el)
        self.get_t_st(self.elem)
        if self.elem.Category.Id.ToString() == '-2008049':self.formteil.append(self.elem.Id)       
        
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
                    if conn.IsConnected:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() not in self.liste:  
                                if owner.Category.Id.ToString() == '-2001140':
                                    return     
                                elif owner.Category.Id.ToString() == '-2008044':
                                    self.rohr.append(owner.Id.ToString())
                                elif owner.Category.Id.ToString() == '-2008049':
                                    self.formteil.append(owner.Id)               
                                self.get_t_st(owner)

el0_ref = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,Filters(),'Wählt ein Rohr aus')
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
    script.exit()
ref_1 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,FilterNeben(liste),'Wählt den nächsten Teil aus')
a = AlleBauteile(Liste_Rohr,doc.GetElement(ref_1))


class Rohr:
    def __init__(self,elemid):
        self.elemid = DB.ElementId(int(elemid))
        self.elem = doc.GetElement(self.elemid)
        self.durchmesser = self.elem.LookupParameter('IGF_X_SM_Durchmesser').AsValueString()
    def wert_schreiben(self):
        self.elem.LookupParameter('Durchmesser').SetValueString(self.durchmesser)

class Rohrformteil:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.doc = None
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
                            if owner.Category.Id.ToString() in ['-2008044','-2008055']:
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
                            if owner.Category.Id.ToString() == '-2008044':#,'-2008055']:
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
            except:
                self.doc.Create.NewElbowFitting(self.conns1, self.conns0)
                
        elif self.art == 'Übergang':
            try:self.doc.Create.NewTransitionFitting(self.conns0, self.conns1)
            except:
                self.doc.Create.NewTransitionFitting(self.conns1, self.conns0)
                
        elif self.art == 'T-Stück':
            try:self.doc.Create.NewTeeFitting(self.conns0, self.conns1,self.conns2)
            except:
                try:self.doc.Create.NewTeeFitting(self.conns2, self.conns0, self.conns1)
                except:self.doc.Create.NewTeeFitting(self.conns1, self.conns2, self.conns0)
                    

liste_formteil = []

for elid in a.formteil:
    formteil = Rohrformteil(elid)
    if formteil.art:
        if formteil.art == 'Übergang':
            if formteil.conns0 and formteil.conns1:
                liste_formteil.append(formteil)
        else:
            liste_formteil.append(formteil)


t = DB.Transaction(doc,'Strang dimensionieren')
t.Start()
doc.Delete(List[DB.ElementId](a.formteil))
doc.Regenerate()
for el in a.rohr:
    try:
        Rohr(el).wert_schreiben()
    except Exception as e:
        logger.error(e)
doc.Regenerate()

with forms.ProgressBar(title="{value}/{max_value} Formteile werden erstellt",cancellable=True, step=1) as pb:
    nicht_bearbeitet = []
    for n,el in enumerate(liste_formteil):
        pb.update_progress(n+1,len(liste_formteil))
        try:
            el.doc = doc
            el.createfittings()
           
        except:
            try:
                doc = __revit__.ActiveUIDocument.Document
                doc.Regenerate()
                el.doc = doc
                el.createfittings()
            except:pass

t.Commit()