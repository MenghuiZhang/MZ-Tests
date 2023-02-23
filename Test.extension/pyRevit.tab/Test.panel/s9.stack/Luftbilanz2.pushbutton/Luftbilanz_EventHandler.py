# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog
import xlsxwriter
import Autodesk.Revit.DB as DB
from pyrevit import script
from System.Collections.Generic import List
from clr import GetClrType
from Luftbilanz_Forms import ABFRAGE,RaumluftbilanzExport
from Luftbilanz_Export import Raumdaten
import os

IGF_LOGO = os.path.dirname(__file__) + '\\' + 'icon.png'
logger = script.get_logger()

class ExtenalEventListe(IExternalEventHandler):
    def __init__(self):
        self.name = ''
        self.XYZAnpassen = DB.XYZ(1/3.048,1/3.048,1/3.048)
        self.GUI = None
        self.Executeapp = None
        self.Liste = []
        
    def Execute(self,uiapp):
        try:
            self.Executeapp(uiapp)
        except Exception as e:
            TaskDialog.Show('Fehler',e.ToString())

    def GetName(self):
        return self.name
    
    def ChangeTo3D(self,uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        views = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.View3D)).ToElements()
        for el in views:
            if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
                uidoc.ActiveView = el
                break
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
        del views
        
    def SelectedElements(self,uiapp,Liste):
        uiapp.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId](Liste))
    
    def SelectElements(self,uiapp):
        self.name = 'Element auswählen'
        uiapp.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId](self.Liste))
    
    def RaumAnzeigen(self,uiapp):
        self.name = 'Raum anzeigen'
        self.ChangeTo3D(uiapp)
        self.SelectedElements(uiapp,[self.GUI.mepraum.elemid])
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        t = DB.Transaction(doc,self.name)
        t.Start()
        view.IsSectionBoxActive = True
        box = view.GetSectionBox()
        MEPBox = self.GUI.mepraum.elem.get_BoundingBox(None)
        Max = box.Transform.Inverse.OfPoint(MEPBox.Max + self.XYZAnpassen)
        Min = box.Transform.Inverse.OfPoint(MEPBox.Min - self.XYZAnpassen)
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()
        
        with uidoc.GetOpenUIViews() as views:
            for v in views:
                if v.ViewId == view.Id:
                    try:v.ZoomToFit() 
                    except:pass
                    return
        
        uidoc.Dispose()
        doc.Dispose()
        view.Dispose()
        box.Dispose()
        t.Dispose()
        MEPBox.Dispose()
        del MEPBox
        del uidoc
        del doc
        del view
        del box
        del t
        del Min
        del Max

    def RaumundBauteileanzeigen(self,uiapp):
        self.name = 'Raum und alle Bauteile anzeigen'
        self.ChangeTo3D(uiapp)
        self.SelectedElements(uiapp,[self.GUI.mepraum.elemid])
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        t = DB.Transaction(doc,self.name)
        t.Start()
        view.IsSectionBoxActive = True
        box = view.GetSectionBox()
        MEPBox = self.GUI.mepraum.elem.get_BoundingBox(None)
        max_x,max_y,max_z = MEPBox.Max.X,MEPBox.Max.Y,MEPBox.Max.Z
        min_x,min_y,min_z = MEPBox.Min.X,MEPBox.Min.Y,MEPBox.Min.Z
        for el in self.GUI.mepraum.list_vsr:
            box_temp = el.elem.get_BoundingBox(None)
            max_x = max(max_x,box_temp.Max.X)
            max_y = max(max_y,box_temp.Max.Y)
            max_z = max(max_z,box_temp.Max.Z)
            min_x = min(min_x,box_temp.Min.X)
            min_y = min(min_y,box_temp.Min.Y)
            min_z = min(min_z,box_temp.Min.Z)

        Max = box.Transform.Inverse.OfPoint(DB.XYZ(max_x,max_y,max_z) + self.XYZAnpassen)
        Min = box.Transform.Inverse.OfPoint(DB.XYZ(min_x,min_y,min_z) - self.XYZAnpassen)
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()

        with uidoc.GetOpenUIViews() as views:
            for v in views:
                if v.ViewId == view.Id:
                    try:v.ZoomToFit() 
                    except:pass
                    return
            
        uidoc.Dispose()
        doc.Dispose()
        view.Dispose()
        box.Dispose()
        t.Dispose()
        MEPBox.Dispose()
        del MEPBox
        del uidoc
        del doc
        del view
        del box
        del t
        del Min
        del Max
    
    def RaumluftBerechnen(self,uiapp):
        self.name = 'Raumluftberechnen'
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,'Wert schreiben')
        t.Start()
        for el in self.GUI.mepraum_liste.values():          
            try:el.Nachtbetrieb_Berechnen()
            except Exception as e:print(e)
            try:el.Tagesbetrieb_Berechnen()
            except Exception as e:print(e)
            try:el.Druckstufe_Berechnen()
            except Exception as e:print(e)
            try:el.werte_schreiben()
            except Exception as e:print(e)

        t.Commit()
        
        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        t.Dispose()
        del doc
        del t

    def LuftmengeverteilenMEP(self,uiapp):
        self.name = 'Luftmengeverteilen ' + self.GUI.mepraum.Raumnr
        task = ABFRAGE('Luftmenge in akt. Raum gleichmäßig verteilen?',True,130)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            finally:del task
            return 
        
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        try:self.Luftmengeverteilen(self.GUI.mepraum,task)
        except Exception as e:print(e)
        t.Commit()
        

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        t.Dispose()
        del doc
        del t
        try:
            task.Dispose()
            del task
        except:pass
    
    def LuftmengeverteilenProjekt(self,uiapp):
        self.name = 'Luftmengeverteilen Projekt' 
        task = ABFRAGE('Luftmenge für das Projekt gleichmäßig verteilen?',False,160)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            finally:del task
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.mepraum_liste.values():  
         
            self.Luftmengeverteilen(mep,task)

        t.Commit()
        
        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        t.Dispose()
        del doc
        del t
        try:task.Dispose()
        except:pass
        finally:del task
    
    def Luftmengeverteilen(self,mep,task):
        zu = {}
        ab = {}
        h24 = {}
        lab = {}

        for auslass in mep.list_auslass:
            if auslass.art == 'RZU':
                if auslass.familyandtyp not in zu.keys():
                    zu[auslass.familyandtyp] = [auslass]
                else:
                    zu[auslass.familyandtyp].append(auslass)
            if auslass.art == 'RAB':
                if auslass.familyandtyp not in ab.keys():
                    ab[auslass.familyandtyp] = [auslass]
                else:
                    ab[auslass.familyandtyp].append(auslass)
            if auslass.art == '24h':
                if auslass.familyandtyp not in h24.keys():
                    h24[auslass.familyandtyp] = [auslass]
                else:
                    h24[auslass.familyandtyp].append(auslass)
            if auslass.art == 'LAB':
                if auslass.familyandtyp not in lab.keys():
                    lab[auslass.familyandtyp] = [auslass]
                else:
                    lab[auslass.familyandtyp].append(auslass)
        
        if int(mep.ab_24h.soll) != int(mep.ab_24h.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist) or int(mep.ab_lab_min.soll) != int(mep.ab_lab_min.ist):
            print(30*'-')
            print('24h-Abluft oder Laborabluft in MEP-Raum {} stimmt nicht.'.format(mep.Raumnr))
            print('24h-Abluft-soll: {} m³/h, 24h-Abluft-ist: {} m³/h'.format(mep.ab_24h.soll,mep.ab_24h.ist))
            print('Laborabluftmin-soll: {} m³/h, Laborabluftmin-ist: {} m³/h'.format(mep.ab_lab_min.soll,mep.ab_lab_min.ist))
            print('Laborabluftmax-soll: {} m³/h, Laborabluftmax-ist: {} m³/h'.format(mep.ab_lab_max.soll,mep.ab_lab_max.ist))
            print('Bitte manuell anpassen. ')

        if (int(mep.zu_min.soll) > 0 or int(mep.zu_max.soll) > 0) and len(zu.keys()) == 0:
            print(30*'-')
            print('Es fehlt Zuluftauslass in MEP-Raum {}.'.format(mep.Raumnr))
            print('Zuluft min : {} m³/h, Zuluft max: {} m³/h'.format(mep.zu_min.soll,mep.zu_max.soll))
            print('Bitte manuell anpassen. ')

        if (int(mep.ab_min.soll) > 0 or int(mep.ab_max.soll) > 0) and len(ab.keys()) == 0:
            print(30*'-')
            print('Es fehlt Abluftauslass in MEP-Raum {}.'.format(mep.Raumnr))
            print('Abluft min : {} m³/h, Abluft max: {} m³/h'.format(mep.ab_min.soll,mep.ab_max.soll))
            print('Bitte manuell anpassen. ')
                       
        if len(zu.keys()) == 1:
            for key in zu.keys():
                for auslass in zu[key]:
                    if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll * 1.0 / len(zu[key]),1)
                    if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll * 1.0 / len(zu[key]),1)
                    if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll * 1.0 / len(zu[key]),1)
                    if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll * 1.0 / len(zu[key]),1)
        
        elif len(zu.keys()) > 1:
            sum_luft = 0
            for key in zu.keys():
                for auslass in zu[key]:
                    if auslass.Luftmengenmin > 0:sum_luft += auslass.Luftmengenmin
            if sum_luft != 0:
                for key in zu.keys():
                    for auslass in zu[key]:
                        if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
            else:
                anzahlauslass = 0
                for key in zu.keys():
                    anzahlauslass += len(zu[key])
                for key in zu.keys():
                    for auslass in zu[key]:
                        if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll * 1.0 / anzahlauslass,1)
                        if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll * 1.0 / anzahlauslass,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll * 1.0 / anzahlauslass,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll * 1.0 / anzahlauslass,1)


        if len(ab.keys()) == 1:
            for key1 in ab.keys():
                for auslass in ab[key1]:
                    if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 / len(ab[key1]),1)
                    if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 / len(ab[key1]),1)
                    if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 / len(ab[key1]),1)
                    if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 / len(ab[key1]),1)
        
        elif len(ab.keys()) > 1:
            sum_luft = 0
            for key in ab.keys():
                for auslass in ab[key]:
                    if auslass.Luftmengenmin > 0:sum_luft += auslass.Luftmengenmin
            if sum_luft > 0:
                for key in ab.keys():
                    for auslass in ab[key]:
                        if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)

            else:
                anzahlauslass = 0
                for key in ab.keys():
                    anzahlauslass += len(ab[key])
                for key in ab.keys():
                    for auslass in ab[key]:
                        if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 / anzahlauslass,1)
                        if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 / anzahlauslass,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 / anzahlauslass,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 / anzahlauslass,1)

        
        for vsr in mep.list_vsr:
            if vsr.art in ['RZU','RAB','LAB','24h','RUM']:
                vsr.Luftmengenermitteln_new()
                vsr.vsrauswaelten()
                vsr.vsrueberpruefen()
        self.AederungUebernehmen(mep)

        del zu
        del ab
        del h24
        del lab
    
    def AederungUebernehmen(self,mep):
        mep.werte_schreiben()
        for auslass in mep.list_auslass:auslass.wert_schreiben()
        for vsr in mep.list_vsr:vsr.wert_schreiben()
    
    def AederungUebernehmenMEP(self,uiapp):
        self.name = 'Änderung übernehmen ' + self.GUI.mepraum.Raumnr
        task = ABFRAGE('Luftmenge akt. Raum übernehmen?',True,70,False)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            finally:del task
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        self.AederungUebernehmen(self.GUI.mepraum)
        t.Commit()
        
        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        t.Dispose()
        del doc
        del t
        try:task.Dispose()
        except:pass
        finally:del task
    
    def AederungUebernehmenProjekt(self,uiapp):
        self.name = 'Änderung übernehmen Projekt' 
        task = ABFRAGE('Alle Änderung übernehmen?',False,100,False)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            finally:del task
            return 
        
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.mepraum_liste.values():  
            self.AederungUebernehmen(mep)
        t.Commit()
        
        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        t.Dispose()
        del doc
        del t
        try:task.Dispose()
        except:pass
        finally:del task

    def ExportRaumluftbilanz(self,uiapp):
        temp = RaumluftbilanzExport(self.GUI.path,self.GUI.ListeMEP)
        temp.ShowDialog()
        if not temp.result:
            try:temp.Dispose()
            except:pass
            finally:del temp
            return
        else:
            self.GUI.path = temp.path
        
        Liste = []
        for el in self.GUI.ListeMEP:
            if el.checked:
                Liste.append(el.name)
        path = self.GUI.path + '\\Raumluftbilanz.xlsx'
        e = xlsxwriter.Workbook(path)
        worksheet = e.add_worksheet()
        rowstart = 0
        Liste_Pagebreaks = []
        for mepname in sorted(Liste):  
            mep = self.GUI.mepraum_liste[mepname]
            raumdaten = Raumdaten(mep,rowstart)
            raumdaten.book = e
            raumdaten.sheet = worksheet
            raumdaten.GetFinalExportdaten()
            rowstart = raumdaten.rowende
            Liste_Pagebreaks.extend(raumdaten.rowbreaks)
        worksheet.set_landscape()
        worksheet.set_column(0,3,20)
        worksheet.set_column(4,5,13)
        worksheet.set_column(6,7,24)
        worksheet.set_column(8,10,7)
        worksheet.set_column(11,18,11)
        worksheet.set_column(19,19,18)
        worksheet.set_column(20,20,25)
        worksheet.set_paper(9)
        worksheet.set_h_pagebreaks(Liste_Pagebreaks)   
        header2 = '&L&G'     
        worksheet.set_margins(top=1)
        worksheet.set_footer('&C&p / &N')
        worksheet.set_header(header2, {'image_left': IGF_LOGO})
        e.close() 

        try:temp.Dispose()
        except:pass
        finally:del temp

        del Liste
        del Liste_Pagebreaks
  
    def ExportLuftmengenInSchacht(self,uiapp):
        temp = RaumluftbilanzExport(self.GUI.path,self.GUI.ListeMEP)
        temp.ShowDialog()
        if not temp.result:
            try:temp.Dispose()
            except:pass
            finally:del temp
            return
        else:
            self.GUI.path = temp.path
        
        Liste = []
        for el in self.GUI.ListeMEP:
            if el.checked:
                Liste.append(self.GUI.mepraum_liste[el.name])
        path = self.GUI.path + '\\Luftmengen_Schacht.xlsx'
        e = xlsxwriter.Workbook(path)
        worksheet = e.add_worksheet()
        
        if len(Liste) > 0:
            _dict = self.luftmengen_summieren(Liste)
            raum = Liste[0]
            cell_format = e.add_format()
            cell_format.set_num_format('''0 "m³/h"''')
            for n,schacht in enumerate(sorted(raum.liste_schacht)):
                worksheet.merge_range(0,3*n+1,0,3*n+3,schacht)
                worksheet.write(1, 3*n+1, 'Raumzuluft')
                worksheet.write(1, 3*n+2, 'Raumabluft')
                worksheet.write(1, 3*n+3, '24h-Abluft')
            for n_ebene,ebene in enumerate(sorted(_dict.keys())):
                worksheet.write(n_ebene+2,0, ebene)
                for n_schahct,schacht in enumerate(sorted(_dict[ebene].keys())):
                    worksheet.write_number(n_ebene+2, 3*n_schahct+1, _dict[ebene][schacht][0],cell_format)
                    worksheet.write_number(n_ebene+2, 3*n_schahct+2, _dict[ebene][schacht][1],cell_format)
                    worksheet.write_number(n_ebene+2, 3*n_schahct+3, _dict[ebene][schacht][2],cell_format)
        e.close() 

        try:temp.Dispose()
        except:pass
        finally:del temp
        del Liste

    def luftmengen_summieren(self,Liste_raum):
        _dict = {}
        for raum in Liste_raum:
            if raum.ebene not in _dict.keys():
                _dict[raum.ebene] = {}
                for schacht in raum.liste_schacht:
                    _dict[raum.ebene][schacht] = [0,0,0]
            if raum.rzu_Schacht.nr in _dict[raum.ebene].keys():
                _dict[raum.ebene][raum.rzu_Schacht.nr][0] += raum.rzu_Schacht.menge
            if raum.rab_Schacht.nr in _dict[raum.ebene].keys():
                _dict[raum.ebene][raum.rab_Schacht.nr][1] += raum.rab_Schacht.menge
            if raum.lab_Schacht.nr in _dict[raum.ebene].keys():
                _dict[raum.ebene][raum.lab_Schacht.nr][1] += raum.lab_Schacht.menge
            if raum._24h_Schacht.nr in _dict[raum.ebene].keys():
                _dict[raum.ebene][raum._24h_Schacht.nr][2] += raum._24h_Schacht.menge
        return _dict
            
    def vsranpassen(self,uiapp):
        self.name = 'VSR anpassen' 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for vsr in self.GUI.mepraum.list_vsr0:
            try:vsr.changetype()
            except Exception as e:logger.error(e)
        for vsr in self.GUI.mepraum.list_vsr1:
            try:vsr.changetype()
            except Exception as e:logger.error(e)
        t.Commit()
        t.Dispose()
        doc.Dispose()
        del doc
        del t
