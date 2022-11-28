from calendar import c
from rpw import DB,revit
import math

doc = revit.doc

class Rund_Uebergang:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.groess_conn = None
        self.klein_conn = None
        
        self.Abstand = None
        self.Analyse()
        self.Flaeche = self.get_Flaeche()
    
    def Analyse(self):
        conns = list(self.elem.MEPModel.ConnectorManager.Connectors)
        temp1 = conns[0].Radius
        temp2 = conns[1].Radius
        if temp1 > temp2:
            self.groess_conn = temp1*0.3048
            self.klein_conn = temp2*0.3048
        else:
            self.klein_conn = temp1*0.3048
            self.groess_conn = temp2*0.3048
        
        self.Abstand = conns[0].Origin.DistanceTo(conns[1].Origin)*0.3048
    def get_Flaeche(self):
        if self.groess_conn == self.klein_conn:
            fl = math.pi * 2 * self.groess_conn * self.Abstand
        else:
            fl = math.pi*self.groess_conn*math.sqrt(self.groess_conn*self.Abstand/(self.groess_conn-self.klein_conn)**2+self.groess_conn**2)-\
            math.pi*self.klein_conn*math.sqrt(self.klein_conn*self.Abstand/(self.groess_conn-self.klein_conn)**2+self.klein_conn**2)
        if fl > 1:
            return fl
        else:
            return 1.0

class Bogen_Rund:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.groess_conn = None
        self.klein_conn = None
        self.Abstand = None
        self.Angel = self.elem.LookupParameter('Angle').AsDouble()
        self.Analyse()
        self.Flaeche = self.get_Flaeche()
    
    def Analyse(self):
        conns = list(self.elem.MEPModel.ConnectorManager.Connectors)
        temp1 = conns[0].Radius
        temp2 = conns[1].Radius
        if temp1 > temp2:
            self.groess_conn = temp1*0.3048
            self.klein_conn = temp2*0.3048
        else:
            self.klein_conn = temp1*0.3048
            self.groess_conn = temp2*0.3048
        
        self.Abstand = conns[0].Origin.DistanceTo(conns[1].Origin)*0.3048
    def get_Flaeche(self):
        fl = self.Abstand/2/math.sin(self.Angel/2)*self.Angel*2*math.pi*self.groess_conn
        if fl > 1:
            return fl
        else:
            return 1.0   

class RE_RU_Ueber:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.rund = None
        self.breite = None
        self.hoehe = None
        self.Abstand = None
        self.Angel = self.elem.LookupParameter('Angle').AsDouble()
        self.Analyse()
        self.Flaeche = self.get_Flaeche()
    
    def Analyse(self):
        conns = list(self.elem.MEPModel.ConnectorManager.Connectors)
        try:
            self.rund = conns[0].Radius*0.3048
            self.breite = conns[1].Height*0.3048
            self.hoehe = conns[1].Width*0.3048        
        except:
            self.rund = conns[1].Radius*0.3048
            self.breite = conns[0].Height*0.3048
            self.hoehe = conns[0].Width*0.3048
        self.Abstand = conns[0].Origin.DistanceTo(conns[1].Origin)*0.3048

    def get_Flaeche(self):
        fl = self.Abstand/2*((self.breite+self.hoehe)*2+self.rund*2*math.pi)+self.breite*self.hoehe
        if fl > 1:
            return fl
        else:
            return 1.0 


class Bogen_Re:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.A = self.elem.LookupParameter('a').AsDouble()*0.3048
        self.B = self.elem.LookupParameter('b').AsDouble()*0.3048
        self.R = self.elem.LookupParameter('r').AsDouble()*0.3048
        self.Angel = self.elem.LookupParameter('alpha').AsDouble()
        self.Flaeche = self.get_Flaeche()
    
    def get_Flaeche(self):
        fl = self.Angel*self.R*self.A+self.Angel*(self.R+self.B)*self.A+math.pi*((self.B+self.R)**2-self.R**2)/2
        if fl > 1:
            return fl
        else:
            return 1.0  