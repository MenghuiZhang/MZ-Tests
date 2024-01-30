# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection,RevitCommandId,PostableCommand
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from pyrevit import script
from System.Windows.Media import Brushes
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Windows import Visibility
from System.Windows.Input import Key
import System


class AlleBauteile:
    def __init__(self,StartRohr,ZweiteBauteil):
        self.StartRohr = StartRohr
        self.ZweiteBauteil = ZweiteBauteil
        self.Liste_Rohr = []
        self.Liste_Formteile = []
        # self.Liste_EndDecke = []
        # self.Liste_Flex = []
        self.Liste_Endverbraucher = []
        # self.Liste_HLSBauteile = []

        self.Liste = []
        self.Liste.append(self.StartRohr.Id.ToString())

        self.Liste_Rohr.append(self.StartRohr.Id.ToString())

        # if self.ZweiteBauteil.Category.Id.ToString() == '-2001140':
        #     self.Liste_HLSBauteile.append(self.ZweiteBauteil.Id.ToString())
        if self.ZweiteBauteil.Category.Id.ToString() == '-2008049':
            self.Liste_Formteile.append(self.ZweiteBauteil.Id.ToString())

        self.get_Alle()


    def set_up(self):
        self.Liste_Rohr = []
        self.Liste_Formteile = []
        self.Liste_Endverbraucher = []
        self.Liste = []
        self.Liste.append(self.StartRohr.Id.ToString())
        self.Liste_Rohr.append(self.StartRohr.Id.ToString())
        if self.ZweiteBauteil.Category.Id.ToString() == '-2008049':
            self.Liste_Formteile.append(self.ZweiteBauteil.Id.ToString())
        self.get_Alle()
    
    def get_Alle(self):
        try:self.get_Alle_Bauteile(self.ZweiteBauteil)
        except Exception as e:print(e) 
        try:self.Daten_Bearbeiten()
        except Exception as e:print(e) 

    
    def get_Alle_Bauteile(self,elem):
        self.Liste.append(elem.Id.ToString())
        cate = elem.Category.Name

        if cate not in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:                    
                for conn in conns:
                    if conn.IsConnected:
                        for ref in conn.AllRefs:
                            if ref.Owner.Id.ToString() not in self.Liste and ref.Owner.Category.Name not in ['Rohr Systeme','Rohrdämmung']:  
                                if ref.Owner.Category.Id.ToString() == '-2008049': ## Rohr
                                    if self.IstEndVerbraucher_Formteil(ref.Owner):
                                        self.Liste.append(ref.Owner.Id.ToString())
                                        return
                                self.get_Alle_Bauteile(ref.Owner)
    
    def Daten_Bearbeiten(self):
        self.Liste = set(self.Liste)
        self.Liste = list(self.Liste)
        for elid in self.Liste:
            if elid == self.StartRohr.Id.ToString():
                continue
            elem = self.StartRohr.Document.GetElement(DB.ElementId(int(elid)))
            if self.IstEndVerbraucher(elem):
                self.Liste_Endverbraucher.append(elid)
                elem.Dispose()
                continue

            if elem.Category.Id.ToString() == '-2008044': ## Rohr
                self.Liste_Rohr.append(elid)
            elif elem.Category.Id.ToString() == '-2008049': ## Rohrfromteile
                if self.IstEndVerbraucher_Formteil(elem):
                    self.Liste_Endverbraucher.append(elid)
                    elem.Dispose()
                    continue
                self.Liste_Formteile.append(elid)
            elif elem.Category.Id.ToString() == '-2008099':
                self.Liste_Endverbraucher.append(elid)
                elem.Dispose()
                continue
            elem.Dispose()
    
    def IstEndVerbraucher_Formteil(self,elem):
        try:
            if elem.LookupParameter('IGF_X_Anzahl_Sprinkler_soll').AsInteger() > 0:
                return True
        except:pass
        return False


    # def get_Alle_Bauteile(self,elem):
    #     self.Liste.append(elem.Id.ToString())
    #     cate = elem.Category.Name

    #     if cate not in ['Rohr Systeme','Rohrdämmung']:
    #         conns = None
    #         try:conns = elem.MEPModel.ConnectorManager.Connectors
    #         except:
    #             try:conns = elem.ConnectorManager.Connectors
    #             except:pass
    #         if conns:                    
    #             for conn in conns:
    #                 if conn.IsConnected:
    #                     for ref in conn.AllRefs:
    #                         owner = ref.Owner
    #                         if owner.Id.ToString() not in self.Liste and owner.Category.Name not in ['Rohr Systeme','Rohrdämmung']:  
    #                             if owner.Category.Id.ToString() == '-2001140': ## HLS-Bauteile
    #                                 if self.IstEndVerbraucher(owner):
    #                                     self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                     return

    #                             elif owner.Category.Id.ToString() == '-2008044': ## Rohr
    #                                 self.Liste_Rohr.append(owner.Id.ToString())
    #                                 if self.IstEndVerbraucher(owner):
    #                                     self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                     return
                                    
    #                             elif owner.Category.Id.ToString() == '-2008049': ## Rohrfromteile
    #                                 try:
    #                                     if owner.LookupParameter('IGF_X_Anzahl_Sprinkler_soll').AsInteger() > 0:
    #                                         self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                         return
    #                                 except:pass

    #                                 if owner.MEPModel.ConnectorManager.Connectors.Size != 1:
    #                                     self.Liste_Formteile.append(owner.Id.ToString())      
    #                                 else:
    #                                     self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                     return

    #                                 if self.IstEndVerbraucher(owner):
    #                                     self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                     return

    #                             elif owner.Category.Id.ToString() == '-2008099':
    #                                 self.Liste_Endverbraucher.append(owner.Id.ToString())    
    #                                 return   

    #                             elif owner.Category.Id.ToString() == '-2008055':
    #                                 if self.IstEndVerbraucher(owner):
    #                                     self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                     return  
    #                             elif owner.Category.Id.ToString() == '-2008050':
    #                                 if self.IstEndVerbraucher(owner):
    #                                     self.Liste_Endverbraucher.append(owner.Id.ToString())
    #                                     return

    #                             self.get_Alle_Bauteile(owner)


    def IstEndVerbraucher(self,elem):
        if elem.Category.Id.ToString() in ['-2008044','-2008050']:
            try:return elem.ConnectorManager.UnusedConnectors.Size + 1 == elem.ConnectorManager.Connectors.Size
            except:return False
        try:return elem.MEPModel.ConnectorManager.UnusedConnectors.Size + 1 == elem.MEPModel.ConnectorManager.Connectors.Size
        except:return False
        
class Endverbraucher:
    def __init__(self,elem,AlleListe,dict_dimension_Neu = None,List_dimension = [],grunddimension = None):
        self.elem = elem
        self.T_Stueck = None
        self.dict_dimension_Neu = dict_dimension_Neu
        self.List_dimension = List_dimension
        self.grunddimension = grunddimension
        self.Liste = AlleListe
        self.TMP_Liste = []
        self.Liste_Rohr = []
        self.Anzahl = 0
        self.Dimension = grunddimension
        self.Category = self.elem.Category.Id.ToString()
        if self.Category == '-2008099':self.Anzahl = 1
        else:
            if self.elem.LookupParameter('IGF_X_Anzahl_Sprinkler_soll'):
                self.Anzahl = self.elem.LookupParameter('IGF_X_Anzahl_Sprinkler_soll').AsInteger()
                    
        self.get_T_Stueck(self.elem)
        self.TMP_Liste = []
        
    def get_T_Stueck(self,elem):
        elemid = elem.Id.ToString()
        if self.T_Stueck:return
        self.TMP_Liste.append(elemid)
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
                            if owner.Id.ToString() not in self.TMP_Liste and owner.Id.ToString() in self.Liste:  
                                if owner.Category.Id.ToString() in ['-2001140','-2008044']:
                                    self.Liste_Rohr.append(owner.Id.ToString())  
                                elif owner.Category.Id.ToString() == '-2008049':
                                    conn_temp = owner.MEPModel.ConnectorManager.Connectors
                                    if conn_temp.Size == 3:
                                        self.T_Stueck = owner.Id.ToString()
                                        conn_temp.Dispose()
                                        return
                                    conn_temp.Dispose()
                                self.get_T_Stueck(owner)
    
    def get_Dimension(self):
        for n in self.List_dimension:
            if self.Anzahl >= n:
                self.Dimension = self.dict_dimension_Neu[n]
                break
    
    def wert_schreiben(self):
        for elid in self.Liste_Rohr:
            elem = self.elem.Document.GetElement(DB.ElementId(int(elid)))
            if elem.Category.Id.ToString() == '-2008044':
                try:
                    elem.LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.Dimension))
                except Exception as e:
                    print(e)
            try:
                elem.LookupParameter('IGF_X_Anzahl_Sprinkler').Set(self.Anzahl)
            except Exception as e:
                print(e)
        
class T_Stueck:
    def __init__(self,elem,StartListe_0,StartListe_1,Anzahl0,Anzahl1,dict_dimension_Neu = None,List_dimension = [],grunddimension = None,AlleBauteile = []):
        self.elem = elem
        self.T_Stueck = None
        self.dict_dimension_Neu = dict_dimension_Neu
        self.AlleBauteile = AlleBauteile
        self.List_dimension = List_dimension
        self.grunddimension = grunddimension
        self.StartListe_0 = StartListe_0
        self.StartListe_1 = StartListe_1
        self.Anzahl0 = Anzahl0
        self.Anzahl1 = Anzahl1
        self.Anzahl2 = 0
        self.TMP_Liste = []
        self.Liste_Rohr = []
        self.Schleife = 0
        self.Dimension = grunddimension
    
    def get_T_Stueck(self,elem):
        elemid = elem.Id.ToString()
        if self.T_Stueck:return
        self.TMP_Liste.append(elemid)
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
                            if owner.Id.ToString() not in self.TMP_Liste and owner.Id.ToString() not in self.StartListe_0 and owner.Id.ToString() not in self.StartListe_1 and owner.Id.ToString() in self.AlleBauteile:  
                                if owner.Category.Id.ToString() in ['-2001140','-2008044']:
                                    self.Liste_Rohr.append(owner.Id.ToString())  
                                elif owner.Category.Id.ToString() == '-2008049':
                                    conn_temp = owner.MEPModel.ConnectorManager.Connectors
                                    if conn_temp.Size == 3:
                                        self.T_Stueck = owner.Id.ToString()
                                        return
                                    conn_temp.Dispose()  
           
                                self.get_T_Stueck(owner)
    
    def get_Dimension(self):
        for n in self.List_dimension:
            if self.Anzahl2 >= n:
                self.Dimension = self.dict_dimension_Neu[n]
                break
    
    def wert_schreiben(self):
        for elid in self.Liste_Rohr:
            elem = self.elem.Document.GetElement(DB.ElementId(int(elid)))
            if elem.Category.Id.ToString() == '-2008044':
                try:
                    elem.LookupParameter('IGF_X_SM_Durchmesser').SetValueString(str(self.Dimension))
                except Exception as e:
                    print(e)
            try:
                elem.LookupParameter('IGF_X_Anzahl_Sprinkler').Set(self.Anzahl2)
            except Exception as e:
                print(e)
