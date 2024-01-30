# coding: utf8
from IGF_Massensuazug import ItemTemplate,DB

class Trassen(ItemTemplate):
    def __init__(self,elem):
        ItemTemplate.__init__(self,elem)
        self._Connectors = None
    
    def Reset(self,elem):
        self.elem = elem
        self._Connectors = None
    
    @property
    def Connectors(self):
        if not self._Connectors:
            self._Connectors = list(self.elem.ConnectorManager.Connectors)
        return self._Connectors
    
    @property
    def Connector(self):
        return self.Connector[0]
    
    @property
    def Shape(self):
        if self.Connector.Domain.ToString() == 'Round':return 'RU'
        return 'RE'
    
    @property
    def Durchmesser(self):
        if self.Shape == 'RU':return int(round(self.Connector.Radius*304.8*2))
        return None
    
    @property
    def Hoehe(self):
        if self.Shape == 'RE':return int(round(self.Connector.Height*304.8))
        return None
    
    @property
    def Breite(self):
        if self.Shape == 'RE':return int(round(self.Connector.Width*304.8))
        return None
    
    @property
    def Laenge(self):
        return round(self.elem.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8/1000,2)
    
    @property
    def Flaeche(self):
        return round(self.elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() * 304.8 * 304.8/1000000,2)
    
    @property
    def Groesse(self):
        if self.Shape == 'RU':return 'DN'+str(int(self.Durchmesser))
        return str(max(self.Breite,self.Hoehe))+'x'+str(min(self.Hoehe,self.Breite))
    
    
    
    


        
