# coding: utf8
from IGF_Klasse import DB, ItemTemplateMitName
from IGF_lib import get_value

class HKFamilie(ItemTemplateMitName):
    def __init__(self,name,elems):
        ItemTemplateMitName.__init__(self,name)
        self._selectedindex_H = -1
        self._selectedindex_K = -1
        self._selectedart_K = -1
        self._selectedart_H = -1

        self.Paras = []
        self.checked = False
        self.elems = elems
        self.liste_art_H = ['Heizkörper','Segel','Umlufthitzer','Sonstige','FBH']
        self.liste_art_k = ['Segel','Umluftkühler','Sonstige']

    @property
    def selectedindex_H(self):
        return self._selectedindex_H
    @selectedindex_H.setter
    def selectedindex_H(self,value):
        if value != self._selectedindex_H:
            self._selectedindex_H = value
            self.RaisePropertyChanged('selectedindex_H')
    @property
    def selectedindex_K(self):
        return self._selectedindex_K
    @selectedindex_K.setter
    def selectedindex_K(self,value):
        if value != self._selectedindex_K:
            self._selectedindex_K = value
            self.RaisePropertyChanged('selectedindex_K')
    
    @property
    def selectedart_H(self):
        return self._selectedart_H
    @selectedart_H.setter
    def selectedart_H(self,value):
        if value != self._selectedart_H:
            self._selectedart_H = value
            self.RaisePropertyChanged('selectedart_H')
    
    @property
    def selectedart_K(self):
        return self._selectedart_K
    @selectedart_K.setter
    def selectedart_K(self,value):
        if value != self._selectedart_K:
            self._selectedart_K = value
            self.RaisePropertyChanged('selectedart_K')

class FamilienInstance(object):
    def __init__(self,elem):
        self.elem = elem
        self.raumid = ''
        if not self.Muster_Pruefen():
            self.get_Raum()

    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
    def get_Raum(self):
        try:
            self.raumid = self.elem.Space[self.elem.Document.GetElement(self.elem.CreatedPhaseId)].Id.IntegerValue

        except:
            self.raumid = ''
            print('kein MEP-Raum für Element {} gefunden'.format(self.elem.Id.ToString()))

class Deckensegel(FamilienInstance):
    def __init__(self,elem,param_H,param_K):
        self.param_H = param_H
        self.param_K = param_K
        FamilienInstance.__init__(self,elem)       

        self.heizleistung = get_value(self.elem.LookupParameter(self.param_H))
        self.kuehlleistung = get_value(self.elem.LookupParameter(self.param_K))

class Heizkoerper(FamilienInstance):
    def __init__(self,elem,param_H):
        self.param_H = param_H
        FamilienInstance.__init__(self,elem)       
        self.heizleistung = get_value(self.elem.LookupParameter(self.param_H))

class Umluftkuehler(FamilienInstance):
    def __init__(self,elem,param_K):
        self.param_K = param_K
        FamilienInstance.__init__(self,elem)       
        self.kuehlleistung = get_value(self.elem.LookupParameter(self.param_K))

class MEPRaum_HK:
    def __init__(self,elem,Liste_DES = [],Liste_HK = [],Liste_ULH = [],Liste_FBH = [],Liste_SON = [],Liste_ULK = []):
        self.elem = elem
        self.Liste_DES = Liste_DES
        self.Liste_HK = Liste_HK
        self.Liste_ULH = Liste_ULH
        self.Liste_FBH = Liste_FBH
        self.Liste_SON = Liste_SON
        self.Liste_ULK = Liste_ULK
        self.H_DES = 0
        self.H_FBH = 0
        self.H_HK = 0
        self.H_SON = 0
        self.H_ULH = 0
        self.K_DES = 0
        self.K_ULK = 0

    def DES_berechnen(self):
        for el in self.Liste_DES:
            self.H_DES += el.heizleistung
            self.K_DES += el.kuehlleistung
    
    def HK_berechnen(self):
        for el in self.Liste_HK:
            self.H_HK += el.heizleistung

    def ULK_berechnen(self):
        for el in self.Liste_ULK:
            self.K_ULK += el.kuehlleistung
                    
    
    def wert_schreiben(self,param,wert):
        if wert is not None:
            para = self.elem.LookupParameter(param)
            if para:
                try:
                    if para.StorageType.ToString() == 'Double':para.SetValueString(str(wert))
                    elif para.StorageType.ToString() == 'Integer':para.Set(int(wert))
                    else:para.Set(int(wert))
                except:
                    pass
            else:
                print('Parameter {} nicht vorhanden'.format(param))

    
    def werte_schreiben_DES(self):
        self.wert_schreiben("IGF_H_DeS_Leistung", self.H_DES)
        self.wert_schreiben("IGF_K_DeS_Leistung", self.K_DES)

    def werte_schreiben_HK(self):
        self.wert_schreiben("IGF_H_HK_Leistung", self.H_HK)
    
    def werte_schreiben_ULK(self):
        self.wert_schreiben("IGF_K_ULK_Leistung", self.K_ULK)


    