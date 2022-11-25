# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit,forms
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
import math


class ISO(object):
    def __init__(self,name,elem):
        self.elem = elem
        self.name = name
IS_ISO = ObservableCollection[ISO]()
IS_ISO_LUFT = ObservableCollection[ISO]()

def get_ISO():
    temp = DB.FilteredElementCollector(revit.doc).OfClass(DB.Mechanical.DuctInsulationType).ToElements()
    temp_dict = {}
    for el in temp:
        name = el.LookupParameter('Typname').AsString()
        temp_dict[name] = el.Id
    for el in sorted(temp_dict.keys()):
        IS_ISO_LUFT.Add(ISO(el,temp_dict[el]))
    
    temp = DB.FilteredElementCollector(revit.doc).OfClass(DB.Plumbing.PipeInsulationType).ToElements()
    temp_dict = {}
    for el in temp:
        name = el.LookupParameter('Typname').AsString()
        temp_dict[name] = el.Id
    for el in sorted(temp_dict.keys()):
        IS_ISO.Add(ISO(el,temp_dict[el]))

get_ISO()

liste_category = List[DB.BuiltInCategory]()
liste_category.Add(DB.BuiltInCategory.OST_PipeInsulations)
liste_category.Add(DB.BuiltInCategory.OST_DuctInsulations)
Filter = DB.ElementMulticategoryFilter(liste_category)

class Systemtyp(object):
    def __init__(self,name,liste):
        self.checked = False
        self.name = name
        self.liste = liste

Liste_Systemtyp = ObservableCollection[Systemtyp]()
Liste_Systemtyp_1 = ObservableCollection[Systemtyp]()
def get_Systemtyp():
    coll = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType().ToElements()
    Dict = {}
    for el in coll:
        typ = el.LookupParameter('Typ').AsValueString()
        if typ not in Dict.keys():
            Dict[typ] = [el]
        else:
            Dict[typ].append(el)
    for el in sorted(Dict.keys()):
        Liste_Systemtyp.Add(Systemtyp(el,Dict[el]))
get_Systemtyp()

def get_Systemtyp1():
    coll = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType().ToElements()
    Dict = {}
    for el in coll:
        typ = el.LookupParameter('Typ').AsValueString()
        if typ not in Dict.keys():
            Dict[typ] = [el]
        else:
            Dict[typ].append(el)
    for el in sorted(Dict.keys()):
        Liste_Systemtyp_1.Add(Systemtyp(el,Dict[el]))
get_Systemtyp1()

class Bauteil(object):
    def __init__(self,elem,doc):
        self.elem = elem
        self.doc = doc
        self.Filter = Filter
        self.ISO_Dicke_vorgeben_Pro = ''
        self.ISO_Dicke_vorgeben_mm = ''
        self.ISO_Dicke = 0
        self.ISO_Art = None
        self.ISO_Rohr = ISO_Rohr
        self.vorhanden = self.get_vorhanden() 
    
    def get_vorhanden(self):
        if self.elem.GetDependentElements(self.Filter).Count > 0:
            for el in self.elem.GetDependentElements(self.Filter):
                return self.doc.GetElement(el)
        else:
            return None
    
    def change_ISO(self):
        try:
            self.vorhanden.ChangeTypeId(self.ISO_Art)
            self.vorhanden.Thickness = self.ISO_Dicke
        except:pass

    def get_ISO_Info_0(self):
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
                self.ISO_Dicke = float(ISO_Dicke[:ISO_Dicke.find('mm')]) / 304.8
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
                self.ISO_Dicke = math.ceil(dicke_temp_neu)/30.48
                return

        if self.elem.Category.Id.ToString() in ['-2008049','-2008055']:
            if ISO_Art not in self.ISO_Rohr.keys():
                print('Rohrdämmung {} nicht vorhanden'.format(ISO_Art))
                return
            self.ISO_Art = self.ISO_Rohr[ISO_Art]
            if ISO_Dicke.find('mm') != -1:
                self.ISO_Dicke = float(ISO_Dicke[:ISO_Dicke.find('mm')]) / 304.8
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
                                self.ISO_Dicke = math.ceil(dicke_temp_neu)/30.48
                                return
                            
                    print('Dicke für Element {} kann nicht ermittelt werden. Grund: Formteil/Zubehör nicht mit Rohr verbunden'.format(self.elem.Id.ToString()))
                    return
                print('Dicke für Element {} kann nicht ermittelt werden. Grund: Keine Connectors gefunden'.format(self.elem.Id.ToString()))
                return             

    def get_ISO_Info_1(self):
        if self.elem.Category.Id.ToString() in ['-2008044','-2008050']:
            if self.ISO_Dicke_vorgeben_mm:
                self.ISO_Dicke = float(self.ISO_Dicke_vorgeben_mm) / 304.8
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
            dicke_temp_neu = dicke_temp*int(self.ISO_Dicke_vorgeben_Pro)/1000.0
            self.ISO_Dicke = math.ceil(dicke_temp_neu)/30.48
            return
            
        elif self.elem.Category.Id.ToString() in ['-2008049','-2008055']:
            if self.ISO_Dicke_vorgeben_mm:
                self.ISO_Dicke = float(self.ISO_Dicke_vorgeben_mm) / 304.8
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
                            dicke_temp_neu = dicke_temp*int(self.ISO_Dicke_vorgeben_Pro)/1000.0
                            self.ISO_Dicke = math.ceil(dicke_temp_neu)/30.48
                            return
                            
                print('Dicke für Element {} kann nicht ermittelt werden. Grund: Formteil/Zubehör nicht mit Rohr verbunden'.format(self.elem.Id.ToString()))
                return
            print('Dicke für Element {} kann nicht ermittelt werden. Grund: Keine Connectors gefunden'.format(self.elem.Id.ToString()))
            return
    
    def get_ISO_Info_2(self):
        if self.ISO_Dicke_vorgeben_mm:
            self.ISO_Dicke = float(self.ISO_Dicke_vorgeben_mm) / 304.8
            return       

    def get_ISO_Info_3(self):
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
        self.ISO_Dicke = float(ISO_Dicke[:ISO_Dicke.find('mm')]) / 304.8       
    
    def create_ISO(self):
        try:
            DB.Plumbing.PipeInsulation.Create(self.doc,self.elem.Id,self.ISO_Art,self.ISO_Dicke)
        except:pass
    
    def create_ISO1(self):
        try:
            DB.Mechanical.DuctInsulation.Create(self.doc,self.elem.Id,self.ISO_Art,self.ISO_Dicke)
        except:pass

class EINGABE(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self,'eingabe.xaml',handle_esc=False)
        self.iso = None
    
    def close(self,sender,arg):
        self.Close()

    def changpro(self,sender,arg):
        self.dicke_mm.Text = ''
        self.dicke_pro.Text = '100'
        self.dicke_mm.IsEnabled = False
        self.dicke_pro.IsEnabled = True

    def changmm(self,sender,arg):
        self.dicke_mm.Text = '25'
        self.dicke_pro.Text = ''
        self.dicke_mm.IsEnabled = True
        self.dicke_pro.IsEnabled = False

class EINGABE_LUFT(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self,'eingabe_luft.xaml',handle_esc=False)
        self.iso = None
    
    def close(self,sender,arg):
        self.Close()
    
class ADDISO(IExternalEventHandler):
    def __init__(self):
        self.elems = None
        self.vorhandenart = None
        self.vorhandendicke_mm = None
        self.vorhandendicke_pro = None
        self.vorhandenbearbeiten = False
        self.typ = []
        self.elem = None
        self.system = None
        self.rohr = False
        self.rohrformteil = False
        self.flexrohr = False
        self.rohraccessory = False
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Liste = []
        self.elems = []
        if self.elem:
            for el in uidoc.Selection.GetElementIds():
                elem = doc.GetElement(el)
                if elem.Category.Id.ToString() == '-2008044' and self.rohr:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008049' and self.rohrformteil:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008050' and self.flexrohr:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008055' and self.rohraccessory:
                    self.elems.append(elem)


        elif self.system:
            try:
                system = doc.GetElement(list(uidoc.Selection.GetElementIds())[0])
                elems = system.PipingNetwork
                for elem in elems:
                    if elem.Category.Id.ToString() == '-2008044' and self.rohr:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008049' and self.rohrformteil:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008055' and self.rohraccessory:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008050' and self.flexrohr:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
            except:
                TaskDialog.Show('Fehler','Bitte Rohrsystem auswählen')
                return

        elif self.typ.Count > 0:
            for typ in self.typ:
                for el in typ.liste:
                    for elem in el.PipingNetwork:
                        if elem.Category.Id.ToString() == '-2008044' and self.rohr:
                            if elem.MEPSystem.Id.ToString() == el.Id.ToString():
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008049' and self.rohrformteil:
                            if elem.LookupParameter('Systemname').AsString() == el.Name:
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008055' and self.rohraccessory:
                            if elem.LookupParameter('Systemname').AsString() == el.Name:
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008050' and self.flexrohr:
                            if elem.MEPSystem.Id.ToString() == el.Id.ToString():
                                self.elems.append(elem)
        if self.elems.Count == 0:
            TaskDialog.Show('Info','Keine Rohre/Rohrformteile/Flexrohre/Rohrzubehör gefunden')
            return
        
        for el in self.elems:
            bauteil = Bauteil(el,doc)
            if self.vorhandenart:bauteil.ISO_Art = self.vorhandenart
            if self.vorhandendicke_mm:
                bauteil.ISO_Dicke_vorgeben_mm = self.vorhandendicke_mm
                bauteil.get_ISO_Info_1()
            elif self.vorhandendicke_pro:
                bauteil.ISO_Dicke_vorgeben_Pro = self.vorhandendicke_pro
                bauteil.get_ISO_Info_1()
            else:
                bauteil.get_ISO_Info_0()
            
            if not bauteil.ISO_Dicke and bauteil.ISO_Art:
                continue
            if self.vorhandenbearbeiten:
                Liste.append(bauteil)
            else:
                if bauteil.vorhanden is None:
                    Liste.append(bauteil)

        t = DB.Transaction(doc,'Dämmung erstellen/ampassen')
        t.Start()
        for bauteil in Liste:
            if bauteil.vorhanden:
                bauteil.change_ISO()
            else:
                bauteil.create_ISO()
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Dämmung erstellen/anpassen"

class ADDISO_LUFT(IExternalEventHandler):
    def __init__(self):
        self.elems = None
        self.vorhandenart = None
        self.vorhandendicke_mm = None
        self.vorhandenbearbeiten = False
        self.typ = []
        self.elem = None
        self.system = None
        self.duct = False
        self.ductformteil = False
        self.flexduct = False
        self.ductaccessory = False
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Liste = []
        self.elems = []
        if self.elem:
            for el in uidoc.Selection.GetElementIds():
                elem = doc.GetElement(el)
                if elem.Category.Id.ToString() == '-2008000' and self.duct:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008010' and self.ductformteil:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008020' and self.flexduct:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008016' and self.ductaccessory:
                    self.elems.append(elem)


        elif self.system:
            try:
                system = doc.GetElement(list(uidoc.Selection.GetElementIds())[0])
                elems = system.DuctNetwork
                for elem in elems:
                    if elem.Category.Id.ToString() == '-2008000' and self.duct:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008010' and self.ductformteil:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008016' and self.ductaccessory:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008020' and self.flexduct:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
            except:
                TaskDialog.Show('Fehler','Bitte Rohrsystem auswählen')
                return

        elif self.typ.Count > 0:
            for typ in self.typ:
                for el in typ.liste:
                    for elem in el.DuctNetwork:
                        if elem.Category.Id.ToString() == '-2008000' and self.duct:
                            if elem.MEPSystem.Id.ToString() == el.Id.ToString():
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008010' and self.ductformteil:
                            if elem.LookupParameter('Systemname').AsString() == el.Name:
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008016' and self.ductaccessory:
                            if elem.LookupParameter('Systemname').AsString() == el.Name:
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008020' and self.flexduct:
                            if elem.MEPSystem.Id.ToString() == el.Id.ToString():
                                self.elems.append(elem)
        if self.elems.Count == 0:
            TaskDialog.Show('Info','Keine Luftkanäle/Kanalformteile/Flexkanäle/Kanalzubehör gefunden')
            return
        
        for el in self.elems:
            bauteil = Bauteil(el,doc)
            if self.vorhandenart:bauteil.ISO_Art = self.vorhandenart
            if self.vorhandendicke_mm:
                bauteil.ISO_Dicke_vorgeben_mm = self.vorhandendicke_mm
                bauteil.get_ISO_Info_2()
            
            else:
                bauteil.get_ISO_Info_3()
            
            if not bauteil.ISO_Dicke and bauteil.ISO_Art:
                continue
            if self.vorhandenbearbeiten:
                Liste.append(bauteil)
            else:
                if bauteil.vorhanden is None:
                    Liste.append(bauteil)

        t = DB.Transaction(doc,'Dämmung erstellen/ampassen')
        t.Start()
        for bauteil in Liste:
            if bauteil.vorhanden:
                bauteil.change_ISO()
            else:
                bauteil.create_ISO1()
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Dämmung erstellen/anpassen"

class REMOVEISO(IExternalEventHandler):
    def __init__(self):
        self.elems = None
        self.typ = []
        self.elem = None
        self.system = None
        self.rohr = False
        self.rohrformteil = False
        self.flexrohr = False
        self.rohraccessory = False
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Liste = []

        self.elems = []
        if self.elem:
            for el in uidoc.Selection.GetElementIds():
                elem = doc.GetElement(el)
                if elem.Category.Id.ToString() == '-2008044' and self.rohr:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008049' and self.rohrformteil:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008050' and self.flexrohr:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008055' and self.rohraccessory:
                    self.elems.append(elem)

        elif self.system:
            try:
                system = doc.GetElement(list(uidoc.Selection.GetElementIds())[0])
                elems = system.PipingNetwork
                for elem in elems:
                    if elem.Category.Id.ToString() == '-2008044' and self.rohr:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008049' and self.rohrformteil:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008050' and self.flexrohr:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008055' and self.rohraccessory:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
            except:
                TaskDialog.Show('Fehler','Bitte Rohrsystem auswählen')
                return

        elif self.typ.Count > 0:
            for typ in self.typ:
                for el in typ.liste:
                    for elem in el.PipingNetwork:
                        if elem.Category.Id.ToString() == '-2008044' and self.rohr:
                            if elem.MEPSystem.Id.ToString() == el.Id.ToString():
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008049' and self.rohrformteil:
                            if elem.LookupParameter('Systemname').AsString() == el.Name:
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008050' and self.flexrohr:
                            if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008055' and self.rohraccessory:
                            if elem.LookupParameter('Systemname').AsString() == system.Name:
                                self.elems.append(elem)

        if self.elems.Count == 0:
            TaskDialog.Show('Info','Keine Rohre/Rohrformteile/Flexrohre/Rohrzubehör gefunden')
            return

        Liste = []
        for el in self.elems:
            bauteil = Bauteil(el,doc)
            if bauteil.vorhanden is not None:
                Liste.append(bauteil)

        t = DB.Transaction(doc,'Dämmung entfernen')
        t.Start()
        for bauteil in Liste:
            doc.Delete(bauteil.vorhanden.Id)
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Dämmung entfernen"

class REMOVEISO_LUFT(IExternalEventHandler):
    def __init__(self):
        self.elems = None
        self.typ = []
        self.elem = None
        self.system = None
        self.duct = False
        self.ductaccessory = False
        self.ductformteil = False
        self.flexduct = False
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Liste = []

        self.elems = []
        if self.elem:
            for el in uidoc.Selection.GetElementIds():
                elem = doc.GetElement(el)
                if elem.Category.Id.ToString() == '-2008000' and self.duct:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008010' and self.ductformteil:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008020' and self.flexduct:
                    self.elems.append(elem)
                elif elem.Category.Id.ToString() == '-2008016' and self.ductaccessory:
                    self.elems.append(elem)

        elif self.system:
            try:
                system = doc.GetElement(list(uidoc.Selection.GetElementIds())[0])
                elems = system.DuctNetwork
                for elem in elems:
                    if elem.Category.Id.ToString() == '-2008000' and self.duct:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008010' and self.ductformteil:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008020' and self.flexduct:
                        if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                            self.elems.append(elem)
                    elif elem.Category.Id.ToString() == '-2008016' and self.ductaccessory:
                        if elem.LookupParameter('Systemname').AsString() == system.Name:
                            self.elems.append(elem)
            except:
                TaskDialog.Show('Fehler','Bitte Luftkanal System auswählen')
                return

        elif self.typ.Count > 0:
            for typ in self.typ:
                for el in typ.liste:
                    for elem in el.DuctNetwork:
                        if elem.Category.Id.ToString() == '-2008000' and self.duct:
                            if elem.MEPSystem.Id.ToString() == el.Id.ToString():
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008010' and self.ductformteil:
                            if elem.LookupParameter('Systemname').AsString() == el.Name:
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008020' and self.flexduct:
                            if elem.MEPSystem.Id.ToString() == system.Id.ToString():
                                self.elems.append(elem)
                        elif elem.Category.Id.ToString() == '-2008016' and self.ductaccessory:
                            if elem.LookupParameter('Systemname').AsString() == system.Name:
                                self.elems.append(elem)

        if self.elems.Count == 0:
            TaskDialog.Show('Info','Keine Luftkanäle/Kanalformteile/Flexkanäle/Kanalzubehör gefunden')
            return

        Liste = []
        for el in self.elems:
            bauteil = Bauteil(el,doc)
            if bauteil.vorhanden is not None:
                Liste.append(bauteil)

        t = DB.Transaction(doc,'Dämmung entfernen')
        t.Start()
        for bauteil in Liste:
            doc.Delete(bauteil.vorhanden.Id)
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Dämmung entfernen"