# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
import time


__title__ = "Anlagennummer"
__doc__ = """
Anlagennummern aus Luftauslasssystem in MEP-Raum übertragen (für RAB und RZU)
[2022.02.21]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
start = time.time()
uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass

Ductterminal = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElementIds()

class Luftauslass:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        try:
            self.System = doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId())
            self.SystemArt = self.elem.LookupParameter('Systemtyp').AsValueString()
            self.Anlagennummer = self.System.LookupParameter('IGF_X_AnlagenNr').AsValueString()
        except:
            if not self.Muster:
                logger.error("Element {} hat keine Systemtyp zugewiesen".format(self.elemid.ToString()))
            self.System = ''
            self.SystemArt = ''
            self.Anlagennummer = ''
        if not self.Einbauort:
            if not self.Muster:
                logger.error('Einbauort konnte nicht ermittelt werden, ElementId: {}'.format(self.elemid.ToString()))
        
        if self.SystemArt:
            if self.SystemArt.upper().find('24H') != -1:
                self.SystemArt_neu = '24h'
            elif self.SystemArt.upper().find('TIERHALTUNG') != -1:
                if self.SystemArt.upper().find('ZULUFT') != -1:
                    self.SystemArt_neu = 'TZU'
                elif self.SystemArt.upper().find('ABLUFT') != -1:
                    self.SystemArt_neu = 'TAB'
            else:
                if self.SystemArt.upper().find('ZULUFT') != -1:
                    self.SystemArt_neu = 'RZU'
                elif self.SystemArt.upper().find('ABLUFT') != -1:
                    self.SystemArt_neu = 'RAB'
                elif self.SystemArt.upper().find('ÜBER') != -1:
                    self.SystemArt_neu = 'Über'
                else:self.SystemArt_neu = 'XXX'
        else:
            self.SystemArt_neu = 'XXX'

        
        try:
            self.Typen = self.elem.Symbol.LookupParameter('Typenkommentare').AsString()
            if self.Typen.find('IGF_RLT_Laboranschluss_LAB') != -1:
                self.SystemArt_neu = 'LAB'
        except:
            pass

        # if self.SystemArt_neu == 'XXX':
        #     if not self.Muster:
        #         print(self.elemid)

    @property
    def Einbauort(self):
        return self.elem.Space[doc.Phases[0]]
    
    @property
    def Muster(self):
        try:

            if self.elem.LookupParameter('Bearbeitungsbereich').AsValueString() != 'KG4xx_Musterbereich':
                return False
            else:
                return True
        except:
            return False



dict_MEP_Auslass = {}

with forms.ProgressBar(title='{value}/{max_value} Luftauslässe',cancellable=True, step=int(len(Ductterminal)/200)+1) as pb:
    for n,elemid in enumerate(Ductterminal):
        if pb.cancelled:
            script.exit()
        
        pb.update_progress(n+1,len(Ductterminal))

        luftauslass = Luftauslass(elemid)

        if luftauslass.Muster:
            continue
            
        if luftauslass.SystemArt_neu == 'XXX' or luftauslass.SystemArt_neu == 'Über':
            continue

        if luftauslass.Einbauort:
            if not luftauslass.Einbauort.Id.ToString() in dict_MEP_Auslass.keys():
                dict_MEP_Auslass[luftauslass.Einbauort.Id.ToString()] = [luftauslass]
            else:
                dict_MEP_Auslass[luftauslass.Einbauort.Id.ToString()].append(luftauslass)


class MEPRaum:
    def __init__(self,elemid,auslassliste):
        self.elemid = DB.ElementId(int(elemid))
        self.elem = doc.GetElement(self.elemid)
        self.nummer = self.elem.Number
        self.name = self.elem.LookupParameter('Name').AsString()
        self.List_auslass = auslassliste
        self.anl_rzu = ''
        self.anl_rab = ''
        self.anl_tzu = ''
        self.anl_tab = ''
        self.anl_24h = ''
        self.anl_lab = ''
        self.Analyse()
    
    def Analyse(self):
        liste_rzu = []
        liste_rab = []
        liste_tzu = []
        liste_tab = []
        liste_24h = []
        liste_lab = []


        for el in self.List_auslass:
            if el.Anlagennummer == None:
                continue
            if el.SystemArt_neu == 'TZU':
                if el.Anlagennummer not in liste_tzu:
                    liste_tzu.append(el.Anlagennummer)
            elif el.SystemArt_neu == 'TAB':
                if el.Anlagennummer not in liste_tab:
                    liste_tab.append(el.Anlagennummer)
            elif el.SystemArt_neu == 'RZU':
                if el.Anlagennummer not in liste_rzu:
                    liste_rzu.append(el.Anlagennummer)
            elif el.SystemArt_neu == 'RAB':
                if el.Anlagennummer not in liste_rab:
                    liste_rab.append(el.Anlagennummer)
            elif el.SystemArt_neu == 'LAB':
                if el.Anlagennummer not in liste_lab:
                    liste_lab.append(el.Anlagennummer)
            elif el.SystemArt_neu == '24h':
                if el.Anlagennummer not in liste_24h:
                    liste_24h.append(el.Anlagennummer)

        if len(liste_rzu) < 2:
            if len(liste_rzu) > 0:
                self.anl_rzu = liste_rzu[0]
        else:
            logger.error(30*'-')
            logger.error('mehr als 1 Anlagen für RaumZuluft zugewiesen')
            logger.error('Raum: {}, {}. ElementId: {}'.format(self.nummer,self.name,self.elemid.ToString()))
            logger.error('Anlagennummer: {}'.format(liste_rzu))
        
        if len(liste_rab) < 2:
            if len(liste_rab) > 0:
                self.anl_rab = liste_rab[0]
        else:
            logger.error(30*'-')
            logger.error('mehr als 1 Anlagen für RaumAbluft zugewiesen')
            logger.error('Raum: {}, {}. ElementId: {}'.format(self.nummer,self.name,self.elemid.ToString()))
            logger.error('Anlagennummer: {}'.format(liste_rab))
        
        if len(liste_tzu) < 2:
            if len(liste_tzu) > 0:
                self.anl_tzu = liste_tzu[0]
        else:
            logger.error(30*'-')
            logger.error('mehr als 1 Anlagen für TierhaltungZuluft zugewiesen')
            logger.error('Raum: {}, {}. ElementId: {}'.format(self.nummer,self.name,self.elemid.ToString()))
            logger.error('Anlagennummer: {}'.format(liste_tzu))
        
        if len(liste_tab) < 2:
            if len(liste_tab) > 0:
                self.anl_tab = liste_tab[0]
        else:
            logger.error(30*'-')
            logger.error('mehr als 1 Anlagen für TierhaltungAbluft zugewiesen')
            logger.error('Raum: {}, {}. ElementId: {}'.format(self.nummer,self.name,self.elemid.ToString()))
            logger.error('Anlagennummer: {}'.format(liste_tab))
        
        if len(liste_24h) < 2:
            if len(liste_24h) > 0:
                self.anl_24h = liste_24h[0]
        else:
            logger.error(30*'-')
            logger.error('mehr als 1 Anlagen für Abluft24h zugewiesen')
            logger.error('Raum: {}, {}. ElementId: {}'.format(self.nummer,self.name,self.elemid.ToString()))
            logger.error('Anlagennummer: {}'.format(liste_24h))
        
        if len(liste_lab) < 2:
            if len(liste_lab) > 0:
                self.anl_lab = liste_lab[0]
        else:
            logger.error(30*'-')
            logger.error('mehr als 1 Anlagen für LaborAbluft zugewiesen')
            logger.error('Raum: {}, {}. ElementId: {}'.format(self.nummer,self.name,self.elemid.ToString()))
            logger.error('Anlagennummer: {}'.format(liste_lab))


    def werte_schreiben(self):
        def wert_schreiben(param,wert):
            if wert != '':
                para = self.elem.LookupParameter(param)
                if not para:
                    logger.error("Parameter {} nicht vorhanden.".format(param))
                    return
                try:
                    para.Set(int(wert))
                except Exception as e:
                    logger.error(e)
        wert_schreiben('IGF_L_AnlagenNr_RZU',self.anl_rzu)
        wert_schreiben('IGF_L_AnlagenNr_RAB',self.anl_rab)
        wert_schreiben('IGF_L_AnlagenNr_TZU',self.anl_tzu)
        wert_schreiben('IGF_L_AnlagenNr_TAB',self.anl_tab)
        wert_schreiben('IGF_L_AnlagenNr_LAB',self.anl_lab)
        wert_schreiben('IGF_L_AnlagenNr_24h',self.anl_24h)
                
liste_MEP = dict_MEP_Auslass.keys()[:]
liste_mepraume = []   
with forms.ProgressBar(title='{value}/{max_value} MEP-Räume',cancellable=True, step=int(len(liste_MEP)/200)+1) as pb:
    for n,elemid in enumerate(liste_MEP):
        if pb.cancelled:
            script.exit()
        
        pb.update_progress(n+1,len(liste_MEP))

        mepraum = MEPRaum(elemid,dict_MEP_Auslass[elemid])
        liste_mepraume.append(mepraum)

if forms.alert('AnlagenNummer in MEP-Räume schreiben?', ok=False, yes=True, no=True):

    t = DB.Transaction(doc, "AnlagenNummer")
    t.Start()

    with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=int(len(liste_mepraume)/200)+1) as pb:
        for n, mepraum in enumerate(liste_mepraume):
            if pb.cancelled:
                t.rollBack()
                script.exit()

            pb.update_progress(n+1, len(liste_mepraume))
            mepraum.werte_schreiben()

    t.Commit()
    t.Dispose()