# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
from rpw import DB,revit,UI
import os
from System.Windows.Input import Key

__title__ = "3.0 TS nummerieren"
__doc__ = """

Ts Nummerieren

Parameter: IGF_X_RVT_TS_Nr

Hinweis:
In Schema-Modell und TGA-Modell das Skript durchlaufen und überprüfen: ob die Anzahl der TS in Beide Modell gleich sind.

1. Startnum und Parameter für HLS-Bauteil Id auswählen.
2. Start anklicken.

[2022.08.04]
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

class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.Key = Key
        self.start = False
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.btid.ItemsSource = _params
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
        self.read_config()

    def ok(self, sender, args):
        self.start = True
        self.write_config()
        self.Close() 

    def read_config(self):
        try:
            param = config.param
            if param not in _params:
                config.param = ''
            else:
                self.btid.SelectedItem = param
        except:
            pass
    
    def write_config(self):
        try:
            config.param = self.btid.SelectedItem.ToString()
        except:
            pass
        script.save_config()


    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        else:
            args.Handled = True
gui = AktuelleBerechnung()
try:
    gui.ShowDialog()
except Exception as e:
    logger.error(e)
    gui.Close()
    script.exit()
if gui.start == False:
    script.exit()

Startnum = int(gui.startnr.Text)-1

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
        self.hls_liste = {}
        self.rohr = []
        self.formteil = []
        self.zubehoer = []
        for el in self.rohrliste:
            self.liste.append(el)
            self.rohr.append(el)
        if self.elem.Category.Id.ToString() == '-2008044':
            self.formteil.append(self.elem.Id.ToString())
        elif self.elem.Category.Id.ToString() == '-2008055':
            self.zubehoer.append(self.elem.Id.ToString())  
        self.get_t_st(self.elem)
        
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
                            if owner.Category.Id.ToString() == '-2001140':
                                nr = owner.LookupParameter(config.param).AsString() 
                                self.hls_liste[nr] = owner
                                return     
                            elif owner.Category.Id.ToString() == '-2008044':
                                self.rohr.append(owner.Id.ToString())
                            elif owner.Category.Id.ToString() == '-2008049':
                                self.formteil.append(owner.Id.ToString())
                            elif owner.Category.Id.ToString() == '-2008055':
                                self.zubehoer.append(owner.Id.ToString())                 
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

if len(liste)  ==  0:
    script.exit()
ref_1 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,FilterNeben(liste),'Wählt den nächsten Teil aus')
a = AlleBauteile(Liste_Rohr,doc.GetElement(ref_1))


class Rohr:
    def __init__(self,elem,nr):
        self.nr = nr
        self.elem = elem
    def wert_schreiben(self):
        self.elem.LookupParameter('IGF_X_RVT_TS_Nr').Set(str(self.nr))

class HLS:
    def __init__(self,elem,nr):
        self.elem = elem
        self.nr = nr
        self.wege_3 = ''
        self.liste = []
        self.Liste_rohr = []
        self.Liste_rohrid = []
        self.dimension = ''
        self.get_vsr(self.elem)
    def get_vsr(self,elem):
        if self.wege_3:return
        
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
                if conns.Size > 2 and cate != 'HLS-Bauteile':
                    self.wege_3 = elem.Id.ToString()
                    return
                elif cate == 'Rohrformteile':
                    conns_liste = []
                    for conn in conns:
                        _dn = int(conn.Radius*2*304.8)
                        if _dn not in conns_liste:conns_liste.append(_dn)
                    if len(self.Liste_rohr) > 0 and len(conns_liste) > 1:
                        self.nr += 0.1
                for conn in conns:
                    refs = conn.AllRefs
                    if elem.Id.ToString() != self.elem.Id.ToString():
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() not in self.liste:  
                                if owner.Category.Id.ToString() == '-2008044':  
                                    self.Liste_rohr.append(Rohr(owner,self.nr))  
                                    self.Liste_rohrid.append(owner.Id.ToString())
                                elif owner.Category.Id.ToString() == '-2008055':
                                    if len(self.Liste_rohr) >0:
                                        self.nr += 0.1                                   
                                self.get_vsr(owner)
                    else:
                        for ref in refs:
                            owner = ref.Owner
                            if owner.Id.ToString() in a.rohr or owner.Id.ToString() in a.formteil or owner.Id.ToString() in a.zubehoer:
                                self.dimension = int(conn.Radius*304.8)
                                if owner.Category.Id.ToString() == '-2008044':
                                    self.Liste_rohr.append(Rohr(owner,self.nr))
                                    self.Liste_rohrid.append(owner.Id.ToString())
                                self.get_vsr(owner)
                            else:
                                continue

Rohr_liste = []
Rohrid_liste = []
gesamt_liste = []
T_Stueck_Liste = []
T_Stueck_Liste_Bearbeitet = []

Ids = a.hls_liste.keys()[:]  
Ids.sort()
tsnr = Startnum

for Id in Ids:
    tsnr = 1 + tsnr
    el = a.hls_liste[Id]
    hls = HLS(el,tsnr)
    Rohr_liste.extend(hls.Liste_rohr)
    if hls.wege_3 != '' and hls.wege_3 not in T_Stueck_Liste:
        T_Stueck_Liste.append(hls.wege_3)
    gesamt_liste.extend(hls.liste)
    Rohrid_liste.extend(hls.Liste_rohrid)

class WEGE3:
    def __init__(self,elemid,nr,liste,liste2):
        self.elemid = elemid
        self.t_liste = liste
        self.t_liste2 = liste2
        self.elem = doc.GetElement(DB.ElementId(int(self.elemid)))
        self.nr = nr
        self.wege_3_0 = ''
        self.wege_3_1 = ''
        self.liste = []
        self.Liste_rohr = []
        self.Liste_rohrid = []
        self.Liste_rohr_0 = []
        self.Liste_rohrid_0 = []
        self.Liste_rohr_1 = []
        self.Liste_rohrid_1 = []
        self.dimension = ''
        self.nr_origin = nr
        self.get_vsr(self.elem)
        if self.wege_3_0 in self.t_liste2 and self.wege_3_1 in self.t_liste2:
            index1 = self.t_liste2.index(self.wege_3_0)
            index2 = self.t_liste2.index(self.wege_3_1)
            if index2 > index1:
                for el in self.Liste_rohr_0:
                    el.nr += 1
                for el in self.Liste_rohr_1:
                    el.nr -= 1
        elif self.wege_3_0 not in self.t_liste2 and self.wege_3_1 in self.t_liste2:
            if len(self.Liste_rohr_0) != 0:
                for el in self.Liste_rohr_0:
                    el.nr += 1
                for el in self.Liste_rohr_1:
                    el.nr -= 1
        elif self.wege_3_0 in self.t_liste2 and self.wege_3_1 not in self.t_liste2:
            pass

        elif self.wege_3_0 not in self.t_liste2 and self.wege_3_1 not in self.t_liste2:
            # '',''
            if self.wege_3_0 == '' and self.wege_3_1 == '':pass
            # 'Pass','Pass'
            elif self.wege_3_0 == 'Pass' and self.wege_3_1 == 'Pass':
                if len(self.Liste_rohr_0) > len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
            # 'Pass',''
            elif self.wege_3_0 == 'Pass' and self.wege_3_1 == '':
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
            # '','Pass'
            elif self.wege_3_0 == '' and self.wege_3_1 == 'Pass':
                pass
            # 'Pass','Id'
            elif self.wege_3_0 == 'Pass' and self.wege_3_1 != '':
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
            # 'Id', 'Pass'
            elif self.wege_3_0 != '' and self.wege_3_1 == 'Pass':
                pass
            # '','Id'
            elif self.wege_3_0 == '' and self.wege_3_1 != '':
                if len(self.Liste_rohr_0) > 0 and len(self.Liste_rohr_1) > 0:
                    for el in self.Liste_rohr_0:
                        el.nr += 1
                    for el in self.Liste_rohr_1:
                        el.nr -= 1
            # 'Id', ''
            elif self.wege_3_0 != '' and self.wege_3_1 == '':
                pass
            # 'Id', 'Id'
            else:
                self.Liste_rohr_0 = []
                self.Liste_rohr_1 = []
                self.Liste_rohrid_0 = []
                self.Liste_rohrid_1 = []
                self.wege_3_0 = self.elemid
                self.wege_3_1 = 'Pass'

        self.Liste_rohr.extend(self.Liste_rohr_0)
        self.Liste_rohr.extend(self.Liste_rohr_1)
        self.Liste_rohrid.extend(self.Liste_rohrid_0)
        self.Liste_rohrid.extend(self.Liste_rohrid_1)

    def get_vsr(self,elem):        
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
                if conns.Size == 1:
                    if self.wege_3_0 == '':self.wege_3_0 = 'Pass'
                    elif self.wege_3_1 == '':self.wege_3_1 = 'Pass'
                    if int(self.nr) == int(self.nr_origin):
                        if self.wege_3_0 != '':
                            if len(self.Liste_rohr_0) != 0: 
                                self.nr = int(self.nr_origin) + 1
                    return

                if conns.Size > 2 and elemid != self.elemid:
                    if self.wege_3_0 == '':self.wege_3_0 = elemid
                    elif self.wege_3_1 == '':self.wege_3_1 = elemid

                    if self.wege_3_0 in self.t_liste:
                        self.Liste_rohr_0 = []
                        self.Liste_rohrid_0 = []
                        self.wege_3_0 = 'Pass'
                    if self.wege_3_1 in self.t_liste:
                        self.Liste_rohr_1 = []
                        self.Liste_rohrid_1 = []
                        self.wege_3_1 = 'Pass'

                    if int(self.nr) == int(self.nr_origin):
                        if self.wege_3_0 != '':
                            if len(self.Liste_rohr_0) != 0: 
                                self.nr = int(self.nr_origin) + 1

                    return
                elif cate == 'Rohrformteile':
                    conns_liste = []
                    for conn in conns:
                        _dn = int(conn.Radius*2*304.8)
                        if _dn not in conns_liste:conns_liste.append(_dn)
                    if len(conns_liste) > 1:
                        if len(self.Liste_rohr_0) > 0 and int(self.nr) == int(self.nr_origin):
                            self.nr += 0.1  
                        elif len(self.Liste_rohr_1) > 0:self.nr += 0.1

                for conn in conns:
                    if conn.IsConnected:
                        refs = conn.AllRefs

                        for ref in refs:
                            owner = ref.Owner
                            if self.wege_3_0 == '':
                                if owner.Id.ToString() not in self.liste and owner.Id.ToString() not in Rohrid_liste:  
                                    if len(self.Liste_rohr) == 0:
                                        self.dimension = int(conn.Radius*304.8)   
                                    if owner.Category.Id.ToString() == '-2008044':
                                        self.Liste_rohr_0.append(Rohr(owner,self.nr))  
                                        self.Liste_rohrid_0.append(owner.Id.ToString())
                                    elif owner.Category.Id.ToString() == '-2008055':
                                        if len(self.Liste_rohr_0) >0:
                                            self.nr += 0.1 
                                    elif owner.Category.Id.ToString() == '-2001140':
                                        self.wege_3_0 = 'Pass'
                                        if len(self.Liste_rohrid_0) != 0:
                                            self.nr = int(self.nr_origin) + 1
                                        return                                     
                                    self.get_vsr(owner)
                            elif self.wege_3_1 == '':
                                if owner.Id.ToString() not in self.liste and owner.Id.ToString() not in Rohrid_liste:  
                                    if len(self.Liste_rohr) == 0:
                                        self.dimension = int(conn.Radius*304.8)   
                                    if owner.Category.Id.ToString() == '-2008044':
                                        self.Liste_rohr_1.append(Rohr(owner,self.nr))  
                                        self.Liste_rohrid_1.append(owner.Id.ToString())
                                    elif owner.Category.Id.ToString() == '-2001140':
                                        self.wege_3_1 = 'Pass'
                                        return
                                    elif owner.Category.Id.ToString() == '-2008055':
                                        if len(self.Liste_rohr_1) >0:
                                            self.nr += 0.1 
                                                                      
                                    self.get_vsr(owner)
                            
                    else:
                        if self.wege_3_0 == '':
                            self.wege_3_0 = 'Pass'
                            if len(self.Liste_rohrid_0) != 0:
                                self.nr = int(self.nr_origin) + 1
                        elif self.wege_3_1 == '':self.wege_3_1 = 'Pass'
                   

t_stueck_neu = []
t_stueck_nichtbearbeitet = []

for n,wege_3 in enumerate(T_Stueck_Liste):
    tsnr = tsnr + 1
    wege3 = WEGE3(wege_3,tsnr,T_Stueck_Liste[n+1:],T_Stueck_Liste[:n+1])
    Rohr_liste.extend(wege3.Liste_rohr)
    liste = [len(wege3.Liste_rohr_0) == 0, len(wege3.Liste_rohr_1) == 0]
    liste_neu = [x for x in liste if x == True]
    if len(liste_neu) == 2:
        tsnr = tsnr - 1
    elif len(liste_neu) == 1:
        pass
    elif len(liste_neu) == 0:
        tsnr = tsnr + 1
    if wege3.wege_3_0 == wege3.elemid and wege3.wege_3_1 == 'Pass':
        t_stueck_nichtbearbeitet.append(wege3.wege_3_0)
    else:
        if wege3.wege_3_0 != '' and wege3.wege_3_0 != 'Pass' and wege3.wege_3_0 not in t_stueck_neu and wege3.wege_3_0 not in T_Stueck_Liste:
            t_stueck_neu.append(wege3.wege_3_0)
        if wege3.wege_3_1 != '' and wege3.wege_3_1 != 'Pass' and wege3.wege_3_1 not in t_stueck_neu and wege3.wege_3_1 not in T_Stueck_Liste:
            t_stueck_neu.append(wege3.wege_3_1)
    

    gesamt_liste.extend(wege3.liste)
    Rohrid_liste.extend(wege3.Liste_rohrid)

for el in t_stueck_nichtbearbeitet:
    T_Stueck_Liste.remove(el)
t_stueck_neu.extend(t_stueck_nichtbearbeitet)


while(len(t_stueck_neu)>0):
    T_Stueck_Liste.extend(t_stueck_neu)
    t_stueck_nichtbearbeitet = []
    t_stueck_neu_temp = []
    anzahl = len(T_Stueck_Liste)
    for n,wege_3 in enumerate(t_stueck_neu):
        tsnr = tsnr + 1
        wege3 = WEGE3(wege_3,tsnr,t_stueck_neu[n+1:],T_Stueck_Liste[:n+anzahl])
        Rohr_liste.extend(wege3.Liste_rohr)
        liste = [len(wege3.Liste_rohr_0) == 0, len(wege3.Liste_rohr_1) == 0]
        liste_neu = [x for x in liste if x == True]
        if len(liste_neu) == 2:
            tsnr = tsnr - 1
        elif len(liste_neu) == 1:
            pass
        elif len(liste_neu) == 0:
            tsnr = tsnr + 1
        
        if wege3.wege_3_0 == wege3.elemid and wege3.wege_3_1 == 'Pass':
            t_stueck_nichtbearbeitet.append(wege3.wege_3_0)
        else:
            if wege3.wege_3_0 != '' and wege3.wege_3_0 != 'Pass' and wege3.wege_3_0 not in t_stueck_neu_temp and wege3.wege_3_0 not in T_Stueck_Liste:
                t_stueck_neu_temp.append(wege3.wege_3_0)
            if wege3.wege_3_1 != '' and wege3.wege_3_1 != 'Pass' and wege3.wege_3_1 not in t_stueck_neu_temp and wege3.wege_3_1 not in T_Stueck_Liste:
                t_stueck_neu_temp.append(wege3.wege_3_1)

        gesamt_liste.extend(wege3.liste)
        Rohrid_liste.extend(wege3.Liste_rohrid)

    t_stueck_neu = t_stueck_neu_temp
    for el in t_stueck_nichtbearbeitet:
        T_Stueck_Liste.remove(el)
    t_stueck_neu.extend(t_stueck_nichtbearbeitet)



t = DB.Transaction(doc,'Test')
t.Start()
for el in Rohr_liste:
    el.wert_schreiben()

t.Commit()
t.Dispose()

print('Insgesamt {} Teilstrecken'.format(tsnr-Startnum))
print('Endnr. {} '.format(tsnr))