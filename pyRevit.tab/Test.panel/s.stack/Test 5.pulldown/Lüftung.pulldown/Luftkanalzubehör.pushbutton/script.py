# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from IGF_lib import get_value
from pyrevit import script, forms
from rpw import revit,DB

__title__ = "3.51 Einbauteile in MEP-Räume" 
__doc__ = """


[2021.10.15]
Version: 1.0"""
__author__ = "Menghui Zhang"


try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

doc = revit.doc

Ductterminal_coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType()
Ductterminalids = Ductterminal_coll.ToElementIds()
Ductterminal_coll.Dispose()


class Luftauslass:
    def __init__(self,elementid):
        self.elem_id = elementid
        self.elem = doc.GetElement(self.elem_id)
        self.RoutingListe = []
        self.VSR = ''
        self.Einbauteile = []
        self.Luftmengenmin = 0
        self.Luftmengenmax = 0
        self.Luftmengennacht = 0
        self.Luftmengentiefe = 0
        self.Routingend = False
        try:
            self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:
            self.System = ''

        self.Luftmengenermitteln()
        self.Wirkungsort(self.elem)
        self.einbauort = ''
        self.Einbauort()
        if not self.einbauort:
            logger.error('Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elem_id.ToString()))
    
    def Luftmengenermitteln(self):
        try:
            self.Luftmengenmin = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin')))
        except:
            pass
        try:
            self.Luftmengenmax = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax')))
        except:
            pass
        try:
            self.Luftmengennacht = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht')))
        except:
            pass
        try:
            self.Luftmengentiefe = float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht')))
        except:
            pass
    def Einbauort(self):
        if self.elem.Space[doc.Phases[0]]:
            self.einbauort = self.elem.Space[doc.Phases[0]].Number
    
    def Wirkungsort(self,element):
        elemid = element.Id.ToString()
        self.RoutingListe.append(elemid)
        cate = element.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Luftkanaldämmung außen', 'Luftkanaldämmung innen', 'Rohre', 'Rohrformteile', 'Rohrzubehör','Rohrdämmung']:
            conns = None
            try:
                conns = element.MEPModel.ConnectorManager.Connectors
            except:
                try:
                    conns = element.ConnectorManager.Connectors
                except:
                    pass
            
            if conns:
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner

                        if not owner.Id.ToString() in self.RoutingListe:
                            if owner.Category.Name in ['Luftkanalzubehör','HLS-Bauteile']:
                                faminame = owner.Symbol.FamilyName
                                if not faminame in ['_L_IGF_430_Netztrennung Kanäle','CAX RE Kombirahmen 2fach vertikal','Segeltuchstutzen']:
                                    self.Einbauteile.append(owner.Id.ToString())
                                if faminame.upper().find('VOLUMENSTROMREGLER') != -1:
                                    self.VSR = owner.Id.ToString()
                                
                            elif owner.Category.Name == 'Luftkanalformteile':
                                if self.VSR:
                                    if owner.MEPModel.ConnectorManager.Connectors.Size > 2:
                                        self.Routingend = True
                                        return

                                
                                # elif faminame.find('Schalldämpfer') != -1:
                                #     self.Einbauteile.append(owner.Id.ToString())

                            self.Wirkungsort(owner)

                            if self.VSR and self.Routingend:
                                return


class Einbauteile:
    def __init__(self,elementid,list_durchlass):
        self.elem_id = elementid
        self.liste_auslass = list_durchlass
        self.elem = doc.GetElement(DB.ElementId(int(self.elem_id)))
        self.einbauort = ''
        self.wirkungsort = ''
        self.Luftmengenmin = 0
        self.Luftmengenmax = 0
        self.Luftmengennacht = 0
        self.Luftmengentiefe = 0
        self.MEPSpace = ''
        self.kurztext = ''
        self.category = self.elem.Category.Name
        self.Luftmengenermitteln()
        self.Einbauort()
        self.Wirkungsort()
        self.Text()
        if not self.einbauort:
            logger.error('Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elem_id))

    def Luftmengenermitteln(self):
        for el in self.liste_auslass:
            self.Luftmengenmin += el.Luftmengenmin
            self.Luftmengenmax += el.Luftmengenmax
            self.Luftmengennacht += el.Luftmengennacht
            self.Luftmengentiefe += el.Luftmengentiefe

    def Einbauort(self):
        if self.elem.Space[doc.Phases[0]]:
            self.einbauort = self.elem.Space[doc.Phases[0]].Number
            self.MEPSpace = self.elem.Space[doc.Phases[0]].Id.ToString()

    def Text(self):
        luftart = 'XXX'
        FamName = 'Familienamen'
        typname = 'Typnamen'
        luftmengmin = '0000'
        luftmengmax = '0000'
        luftmengnacht = '0000'
        luftmengtiefe = '0000'
        durchmesse = ''
        systyp = ''
        for el in self.liste_auslass:
            if el.System:
                systyp = el.System
                break

        if systyp:
            if systyp.upper().find('TIERHALTUNG') != -1:
                if systyp.upper().find('ZULUFT') != -1:
                    luftart = 'TZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    luftart = 'TAB'
            else:
                if systyp.upper().find('ZULUFT') != -1:
                    luftart = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    luftart = 'RAB'
        else:
            luftart = 'XXX'
            logger.error('Systemtyp nicht gefunden, ElementId: {}'.format(self.elem_id))
        try:
            FamName = self.elem.Symbol.FamilyName
        except:
            FamName = 'Familienamen'
            logger.error('Familienamen nicht gefunden, ElementId: {}'.format(self.elem_id))
        try:
            typname = self.elem.Name
        except:
            typname = 'Typnamen'
            logger.error('Typnamen nicht gefunden, ElementId: {}'.format(self.elem_id))
        
        try:
            diameter = str(int(list(self.elem.MEPModel.ConnectorManager.Connectors)[0].Radius*609.6))
            while len(diameter) < 4:
                diameter = '0' + diameter
            durchmesse = 'DN ' + diameter
        except:
            try:
                conn = list(self.elem.MEPModel.ConnectorManager.Connectors)[0]
                Height = str(int(conn.Height*304.8))
                Width = str(int(conn.Width*304.8))
                while len(Height) < 4:
                    Height = '0' + Height
                while len(Width) < 4:
                    Width = '0' + Width
                durchmesse = Width + 'x' + Height
            except:
                pass
        def luftmeng(luftmeng):
            if luftmeng is not None:
                luftmeng = str(int(luftmeng))
                while(len(luftmeng) < 4):
                   luftmeng = '0' + luftmeng
                return luftmeng
            else:
                luftmeng = '0000'
        luftmengmin = luftmeng(self.Luftmengenmin)
        luftmengmax = luftmeng(self.Luftmengenmax)
        luftmengnacht = luftmeng(self.Luftmengennacht)
        luftmengtiefe = luftmeng(self.Luftmengentiefe)
       # print(self.Luftmengenmin,self.Luftmengenmax,self.Luftmengennacht,self.Luftmengentiefe)
       # print(luftart,FamName,typname,luftmengmin,luftmengmax,luftmengnacht,luftmengtiefe)
        if luftart == 'XXX':
            try:
                print(self.elem_id,self.elem.LookupParameter('Systemtyp').AsValueString())
            except:
                print(self.elem_id)
        if self.category == 'HLS-Bauteile':
            self.kurztext = luftart + '_' + FamName + '_' + typname + '_' + luftmengmin + '_' + luftmengmax + '_' + luftmengnacht + '_' + luftmengtiefe
        else:
            self.kurztext = luftart + '_' + FamName + '_' + typname + '_' + durchmesse + '_' + luftmengmin + '_' + luftmengmax + '_' + luftmengnacht + '_' + luftmengtiefe

    def Wirkungsort(self):
        for el in self.liste_auslass:
            if self.wirkungsort:
                if self.wirkungsort.find(el.einbauort) == -1:
                    #logger.error('Wirkungsort konnte nicht ermittelt werden, ElementId: {}'.format(self.elem_id))
                    self.wirkungsort += el.einbauort + ', '
                    #return
            else:
                self.wirkungsort = el.einbauort + ', '
        if self.wirkungsort:
            try:
                self.wirkungsort = self.wirkungsort[:-2]
            except:
                pass

    def werte_schreiben(self):
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.elem.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)
        
        try:
            self.Luftmengenmin = get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin'))
        except:
            pass
        try:
            self.Luftmengenmax = get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax'))
        except:
            pass
        try:
            self.Luftmengennacht = get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht'))
        except:
            pass
        try:
            self.Luftmengentiefe = get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht'))
        except:
            pass

        wert_schreiben("IGF_RLT_AuslassVolumenstromMin", round(self.Luftmengenmin))
        wert_schreiben("IGF_RLT_AuslassVolumenstromMax", round(self.Luftmengenmax))
        wert_schreiben("IGF_RLT_AuslassVolumenstromNacht", round(self.Luftmengennacht))
        wert_schreiben("IGF_RLT_AuslassVolumenstromTiefeNacht", round(self.Luftmengentiefe))
        wert_schreiben("IGF_X_Einbauort", self.einbauort)
        wert_schreiben("IGF_X_Wirkungsort", self.wirkungsort)

class MEPRaum:
    def __init__(self,elementid,list_einbauteile):
        self.elem_id = elementid
        self.liste_einbauteile = list_einbauteile
        self.elem = doc.GetElement(DB.ElementId(int(self.elem_id)))
        self.VSR_dict = {}
        self.Bauteile_dict = {}
        self.Zuordnen()


    def Zuordnen(self):
        for el in self.liste_einbauteile:
            if el.kurztext.upper().find('VOLUMENSTROMREGLER') != -1:
                if not el.kurztext in self.VSR_dict.keys():
                    self.VSR_dict[el.kurztext] = 1
                else:
                    self.VSR_dict[el.kurztext] += 1
            else:
                if not el.kurztext in self.Bauteile_dict.keys():
                    self.Bauteile_dict[el.kurztext] = 1
                else:
                    self.Bauteile_dict[el.kurztext] += 1

    def werte_schreiben(self):
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.elem.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)
        
        VSR_Params = ['IGF_RLT_MEP-Raum_VRs_00', 'IGF_RLT_MEP-Raum_VRs_01', 'IGF_RLT_MEP-Raum_VRs_02', 'IGF_RLT_MEP-Raum_VRs_03', 'IGF_RLT_MEP-Raum_VRs_04',
                      'IGF_RLT_MEP-Raum_VRs_05', 'IGF_RLT_MEP-Raum_VRs_06', 'IGF_RLT_MEP-Raum_VRs_07', 'IGF_RLT_MEP-Raum_VRs_08', 'IGF_RLT_MEP-Raum_VRs_09']
        
        Bauteile_Params = ['IGF_RLT_MEP-Raum_Einbau_00', 'IGF_RLT_MEP-Raum_Einbau_01', 'IGF_RLT_MEP-Raum_Einbau_02', 'IGF_RLT_MEP-Raum_Einbau_03', 'IGF_RLT_MEP-Raum_Einbau_04',
                           'IGF_RLT_MEP-Raum_Einbau_05', 'IGF_RLT_MEP-Raum_Einbau_06', 'IGF_RLT_MEP-Raum_Einbau_07', 'IGF_RLT_MEP-Raum_Einbau_08', 'IGF_RLT_MEP-Raum_Einbau_09']

        Liste_VAR = self.VSR_dict.keys()[:]
        Liste_Bauteil = self.Bauteile_dict.keys()[:]
        if len(Liste_VAR) > 0:
            for n in range(len(Liste_VAR)):
                try:
                    anzahl = str(self.VSR_dict[Liste_VAR[n]])
                    while ((len(anzahl)) < 2):
                        anzahl = '0' + anzahl
                    werte = anzahl + '_' + Liste_VAR[n]
                    wert_schreiben(VSR_Params[n],werte)
                except Exception as e:
                    logger.error(e)
                    logger.error("Index: {}, Art: {}, Anzahl: {}".format(n,Liste_VAR[n],self.VSR_dict[Liste_VAR[n]]))
                    logger.error('MEP Raum: {}, ElementId: {}'.format(self.elem.Number,self.elem_id))
                    print(30*'-')
        if len(Liste_Bauteil) > 0:
            for n in range(len(Liste_Bauteil)):
                try:
                    anzahl = str(self.Bauteile_dict[Liste_Bauteil[n]])
                    while ((len(anzahl)) < 2):
                        anzahl = '0' + anzahl
                    werte = anzahl + '_' + Liste_Bauteil[n]
                    wert_schreiben(Bauteile_Params[n],werte)
                except Exception as e:
                    logger.error(e)
                    logger.error("Index: {}, Art: {}, Anzahl: {}".format(n,Liste_Bauteil[n],self.Bauteile_dict[Liste_Bauteil[n]]))
                    logger.error('MEP Raum: {}, ElementId: {}'.format(self.elem.Number,self.elem_id))
                    print(30*'-')
        


auslass_liste = []
einbauteile_dict = {}

with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe",cancellable=True, step=10) as pb:
    for n, space_id in enumerate(Ductterminalids):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(Ductterminalids))
        luftdurchlass = Luftauslass(space_id)

        auslass_liste.append(luftdurchlass)
        for elid in luftdurchlass.Einbauteile:
            if not elid in einbauteile_dict.keys():
                einbauteile_dict[elid] = [luftdurchlass]
            else:
                einbauteile_dict[elid].append(luftdurchlass)


einbauteil_liste = einbauteile_dict.keys()
bauteile_liste = []
MEPSpace_dict = {}
with forms.ProgressBar(title = "{value}/{max_value} Einbauteile -- Luftmengen berechnen",cancellable=True, step=10) as pb1:
    for n, einbauteil_id in enumerate(einbauteil_liste):
        if pb1.cancelled:
            script.exit()

        pb1.update_progress(n + 1, len(einbauteil_liste))
        einbauteil = Einbauteile(einbauteil_id,einbauteile_dict[einbauteil_id])
        bauteile_liste.append(einbauteil)
        
        if not einbauteil.MEPSpace in MEPSpace_dict.keys():
            MEPSpace_dict[einbauteil.MEPSpace] = [einbauteil]
        else:
            MEPSpace_dict[einbauteil.MEPSpace].append(einbauteil)


# Werte zuückschreiben + Abfrage
if forms.alert("Daten in Einbauteile schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Bauteile",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start("Einbau-/Wirkungsort")
        for n, bauteil in enumerate(bauteile_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n + 1, len(bauteile_liste))
            bauteil.werte_schreiben()
        doc.Regenerate()
        t.Commit()

#print(30*'=')
MEPSpace_Liste = MEPSpace_dict.keys()
MEPSpace_liste = []

with forms.ProgressBar(title = "{value}/{max_value} MEP-Räume -- berechnen",cancellable=True, step=5) as pb3:
    for n, MEPId in enumerate(MEPSpace_Liste):
        if pb3.cancelled:
            script.exit()

        pb3.update_progress(n + 1, len(MEPSpace_Liste))
        if MEPId ==  '':
            continue

        mepraum = MEPRaum(MEPId,MEPSpace_dict[MEPId])
        
        MEPSpace_liste.append(mepraum)


# Werte zuückschreiben + Abfrage
if forms.alert("Daten in MEP-Räume schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=5) as pb4:
        t = DB.Transaction(doc)
        t.Start("MEP-Räume")
        for n, mepspace in enumerate(MEPSpace_liste):
            if pb4.cancelled:
                t.RollBack()
                script.exit()
            pb4.update_progress(n + 1, len(MEPSpace_liste))
            mepspace.werte_schreiben()
        t.Commit()