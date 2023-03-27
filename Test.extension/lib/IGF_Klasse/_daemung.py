# coding: utf8
from IGF_Klasse import DB,TemplateItemBase,ItemTemplateMitElemUndName,ObservableCollection
from System.Collections.Generic import List
import math

class VorgabeDaemung(TemplateItemBase):
    def __init__(self, durchmesser = '', dicke = ''):
        TemplateItemBase.__init__(self)
        self._durchmesser_von = durchmesser
        self._durchmesser_bis = ''
        self._dicke = dicke
    
    @property
    def durchmesser_von(self):
        return self._durchmesser_von
    @durchmesser_von.setter
    def durchmesser_von(self,value):
        if value != self._durchmesser_von:
            self._durchmesser_von = value
            self.RaisePropertyChanged('durchmesser_von')
    
    @property
    def durchmesser_bis(self):
        return self._durchmesser_bis
    @durchmesser_bis.setter
    def durchmesser_bis(self,value):
        if value != self._durchmesser_bis:
            self._durchmesser_bis = value
            self.RaisePropertyChanged('durchmesser_bis')
    
    @property
    def dicke(self):
        return self._dicke
    @dicke.setter
    def dicke(self,value):
        if value != self._dicke:
            self._dicke = value
            self.RaisePropertyChanged('dicke')

class TypDaemung(ItemTemplateMitElemUndName):
    def __init__(self,name,elem):
        ItemTemplateMitElemUndName.__init__(self,name,elem)

class Bauteildaemung(object):
    def __init__(self,elem,list_vorgabe = None,liste_daemung = None,typdaemung=None):
        self.elem = elem
        self.doc = self.elem.Document
        self.list_vorgabe = list_vorgabe
        self.typdaemung = typdaemung
        self.liste_daemung = liste_daemung
        self.dicke_daemung = 0
        self.vorhanden_daemung = None

    def change_daemung(self):
        if self.typdaemung and self.dicke_daemung:
            self.vorhanden_daemung.ChangeTypeId(self.typdaemung.elem.Id)
            self.vorhanden_daemung.Thickness = self.dicke_daemung

class Rohrdaemung(Bauteildaemung):
    def __init__(self,elem,list_vorgabe = None,liste_daemung = None,typdaemung=None):
        Bauteildaemung.__init__(self,elem,list_vorgabe,liste_daemung,typdaemung)
        self.get_vorhanden()
        
    
    def get_vorhanden(self):
        # elem_Daemung = self.elem.GetDependentElements(DB.ElementMulticategoryFilter(List[DB.BuiltInCategory]([DB.BuiltInCategory.OST_PipeInsulations])))
        elem_Daemung = self.elem.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_PipeInsulations))
        if elem_Daemung.Count > 0:
            for elid in elem_Daemung:
                self.vorhanden_daemung = self.doc.GetElement(elid)
        
        del elem_Daemung
    
    def create_daemung(self):
        if self.typdaemung and self.dicke_daemung:
            bearbeitungsbereich = self.elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
            self.doc.GetWorksetTable().SetActiveWorksetId(DB.WorksetId(bearbeitungsbereich))
            self.doc.Regenerate()
            DB.Plumbing.PipeInsulation.Create(self.doc,self.elem.Id,self.typdaemung.elem.Id,self.dicke_daemung)

    def get_daemung_dicke(self,art,dicke):
        if art == 'Prozent':
            self.get_daemung_dicke_Prozent(dicke)
        else:
            self.get_daemung_dicke_Fest(dicke)
    
    def get_daemung_dicke_system(self):
        systemtyp = self.doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId())
        try:ISO_Dicke = systemtyp.LookupParameter('IGF_X_Vorgabe_ISO_Dicke').AsString()
        except:ISO_Dicke = None
        if ISO_Dicke:
            if ISO_Dicke.find('%') != -1:
                self.get_daemung_dicke_Prozent(ISO_Dicke.replace('%',''))
            elif ISO_Dicke.find('mm') != -1:
                self.get_daemung_dicke_Fest(ISO_Dicke.replace('mm',''))

    def get_typ_daemung_aus_system(self):
        systemtyp = self.doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId())
        try:ISO_Art = systemtyp.LookupParameter('IGF_X_Vorgabe_ISO_Art').AsString()
        except:ISO_Art = None
        if ISO_Art:
            for el in self.liste_daemung:
                if el.name == ISO_Art:
                    self.typdaemung = el
                    return

    def get_dicke_100(self):
        durch = 0
        if self.elem.Category.Id.ToString() in ['-2008044','-2008050']:
            durch = self.elem.LookupParameter('Außendurchmesser').AsDouble()*304.8
        else:
            conns = self.elem.MEPModel.ConnectorManager.Connectors
            if conns:
                dict_conn = {}
                for temp in conns:
                    if int(temp.Radius*304.8*2) not in dict_conn.keys():
                        dict_conn[int(temp.Radius*304.8*2)] = [temp]
                    else:
                        dict_conn[int(temp.Radius*304.8*2)].Add(temp)
                conns = dict_conn[sorted(dict_conn.keys())[-1]]
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Category.Id.ToString() in ['-2008044','-2008050']:
                            durch = owner.LookupParameter('Außendurchmesser').AsDouble()*304.8
                            break
                if not durch:durch = sorted(dict_conn.keys())[-1]
        dicke_temp = 100
        for el in self.list_vorgabe:
            try:
                if durch < el[0]:
                    dicke_temp = el[1]
                    break 
            except:
                dicke_temp = el[1]
                break 
        return dicke_temp
                    
    def get_daemung_dicke_Prozent(self,prozent):
        temp_dicke = self.get_dicke_100()
        dicke_temp_neu = temp_dicke*int(prozent) / 100
        self.dicke_daemung = math.ceil(dicke_temp_neu) / 304.8

    def get_daemung_dicke_Fest(self,fest):
        self.dicke_daemung = round(float(fest)) / 304.8

    def get_info_aus_system(self):
        try:
            systemtyp = self.doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId())
        except:
            print('Systemtyp von Element {} kann nicht ermittelt werden'.format(self.elem.Id.ToString()))
            return
        try:ISO_Dicke = systemtyp.LookupParameter('IGF_X_Vorgabe_ISO_Dicke').AsString()
        except:ISO_Dicke = None
        try:ISO_Art = systemtyp.LookupParameter('IGF_X_Vorgabe_ISO_Art').AsString()
        except:ISO_Art = None
        if not (ISO_Dicke and ISO_Art):
            print('Dämmungtyp/ -Dicke von Element {} kann nicht über Systemtyp ermittelt werden.'.format(self.elem.Id.ToString()))
            return

        if self.elem.Category.Id.ToString() in ['-2008044','-2008050']:
            if ISO_Art not in self.ISO_Rohr.keys():
                print('Rohrdämmung {} nicht vorhanden'.format(ISO_Art))
                return
            self.ISO_Art = self.ISO_Rohr[ISO_Art]
            if ISO_Dicke.find('mm') != -1:
                self.dicke_daemung = float(ISO_Dicke[:ISO_Dicke.find('mm')]) / 304.8
                return

            if ISO_Dicke.find('%') != -1:
                durch = self.elem.LookupParameter('Außendurchmesser').AsDouble()*304.8
                pro = ISO_Dicke[:ISO_Dicke.find('%')]
           
                if durch < 28:
                    dicke_temp = 20
                elif durch < 42:
                    dicke_temp = 30
                elif durch < 48:
                    dicke_temp = 40
                elif durch < 54:
                    dicke_temp = 50
                elif durch < 70:
                    dicke_temp = 60
                elif durch < 76:
                    dicke_temp = 70
                elif durch < 89:
                    dicke_temp = 80
                else:
                    dicke_temp = 100
                dicke_temp_neu = dicke_temp*int(pro)/1000.0
                self.dicke_daemung = math.ceil(dicke_temp_neu)/30.48
                return
        if self.elem.Category.Id.ToString() in ['-2008049','-2008055']:
            if ISO_Art not in self.ISO_Rohr.keys():
                print('Rohrdämmung {} nicht vorhanden'.format(ISO_Art))
                return
            self.ISO_Art = self.ISO_Rohr[ISO_Art]
            if ISO_Dicke.find('mm') != -1:
                self.dicke_daemung = float(ISO_Dicke[:ISO_Dicke.find('mm')]) / 304.8
                return
            if ISO_Dicke.find('%') != -1:
                pro = ISO_Dicke[:ISO_Dicke.find('%')]
        
                conns = self.elem.MEPModel.ConnectorManager.Connectors
                if conns:
                    conn = {}
                    for temp in conns:
                        if int(temp.Radius*304.8*2) not in conn.keys():
                            conn[int(temp.Radius*304.8*2)] = [temp]
                        else:
                            conn[int(temp.Radius*304.8*2)].Add(temp)
                    conns = conn[sorted(conn.keys())[-1]]
                    for conn in conns:
                        refs = conn.AllRefs
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Category.Id.ToString() in ['-2008044','-2008050']:
                                durch = owner.LookupParameter('Außendurchmesser').AsDouble()*304.8
                                if durch < 28:
                                    dicke_temp = 20
                                elif durch < 42:
                                    dicke_temp = 30
                                elif durch < 48:
                                    dicke_temp = 40
                                elif durch < 54:
                                    dicke_temp = 50
                                elif durch < 70:
                                    dicke_temp = 60
                                elif durch < 76:
                                    dicke_temp = 70
                                elif durch < 89:
                                    dicke_temp = 80
                                else:
                                    dicke_temp = 100
                                dicke_temp_neu = dicke_temp*int(pro)/1000.0
                                self.dicke_daemung = math.ceil(dicke_temp_neu)/30.48
                                return
                            
                    print('Dicke für Element {} kann nicht ermittelt werden. Grund: Formteil/Zubehör nicht mit Rohr verbunden'.format(self.elem.Id.ToString()))
                    return
                print('Dicke für Element {} kann nicht ermittelt werden. Grund: Keine Connectors gefunden'.format(self.elem.Id.ToString()))
                return   
        
    def get_info_eingabe(self):
        if self.elem.Category.Id.ToString() in ['-2008044','-2008050']:
            if self.dicke_daemung_vorgeben_mm:
                self.dicke_daemung = float(self.dicke_daemung_vorgeben_mm) / 304.8
                return
            durch = self.elem.LookupParameter('Außendurchmesser').AsDouble()*304.8

            if durch < 28:
                dicke_temp = 20
            elif durch < 42:
                dicke_temp = 30
            elif durch < 48:
                dicke_temp = 40
            elif durch < 54:
                dicke_temp = 50
            elif durch < 70:
                dicke_temp = 60
            elif durch < 76:
                dicke_temp = 70
            elif durch < 89:
                dicke_temp = 80
            else:
                dicke_temp = 100
            dicke_temp_neu = dicke_temp*int(self.dicke_daemung_vorgeben_Pro)/1000.0
            self.dicke_daemung = math.ceil(dicke_temp_neu)/30.48
            return
            
        elif self.elem.Category.Id.ToString() in ['-2008049','-2008055']:
            if self.dicke_daemung_vorgeben_mm:
                self.dicke_daemung = float(self.dicke_daemung_vorgeben_mm) / 304.8
                return
            conns = self.elem.MEPModel.ConnectorManager.Connectors
            if conns:
                conn = {}
                for temp in conns:
                    if int(temp.Radius*304.8*2) not in conn.keys():
                        conn[int(temp.Radius*304.8*2)] = [temp]
                    else:
                        conn[int(temp.Radius*304.8*2)].Add(temp)
                conns = conn[sorted(conn.keys())[-1]]
                
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Category.Id.ToString() in ['-2008044','-2008050']:
                            durch = owner.LookupParameter('Außendurchmesser').AsDouble()*304.8
                            if durch < 28:
                                dicke_temp = 20
                            elif durch < 42:
                                dicke_temp = 30
                            elif durch < 48:
                                dicke_temp = 40
                            elif durch < 54:
                                dicke_temp = 50
                            elif durch < 70:
                                dicke_temp = 60
                            elif durch < 76:
                                dicke_temp = 70
                            elif durch < 89:
                                dicke_temp = 80
                            else:
                                dicke_temp = 100
                            dicke_temp_neu = dicke_temp*int(self.dicke_daemung_vorgeben_Pro)/1000.0
                            self.dicke_daemung = math.ceil(dicke_temp_neu)/30.48
                            return
                            
                print('Dicke für Element {} kann nicht ermittelt werden. Grund: Formteil/Zubehör nicht mit Rohr verbunden'.format(self.elem.Id.ToString()))
                return
            print('Dicke für Element {} kann nicht ermittelt werden. Grund: Keine Connectors gefunden'.format(self.elem.Id.ToString()))
            return

class Daemung:
    @staticmethod
    def get_rohrdaemung():
        dict_Daemung = {el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():el for el in DB.FilteredElementCollector(__revit__.ActiveUIDocument.Document).\
                        OfClass(DB.Plumbing.PipeInsulationType).ToElements()}
        liste = ObservableCollection[TypDaemung]()
        for name in sorted(dict_Daemung.keys()):
            liste.Add(TypDaemung(name,dict_Daemung[name]))
        return liste
    
        