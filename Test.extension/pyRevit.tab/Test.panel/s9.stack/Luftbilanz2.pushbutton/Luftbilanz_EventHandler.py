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
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                break
        
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
        del views

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

        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                break
            
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
        del views
    
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
        except Exception as e:logger.error(e)
        t.Commit()
        
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
        
        doc.Dispose()
        t.Dispose()
        del doc
        del t
        try:task.Dispose()
        except:pass
        finally:del task
    
    def Luftmengeverteilen(self,mep,task):
        zu = mep.Dict_RZU
        ab = mep.Dict_RAB
        
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
        
        min_rab = float(mep.D_T_MIN_AbS.soll - mep.D_T_MIN_Lab.ist - mep.D_T_MIN_24h.ist - mep.D_T_MIN_AbUe.ist - mep.D_T_MIN_AbUeM.ist)
        min_rzu = float(mep.D_T_MIN_Rzu.soll)  
        if min_rab < 0:
            min_rzu -= min_rab
            min_rab = 0
        
        max_rab = float(mep.D_T_MAX_AbS.soll - mep.D_T_MAX_Lab.ist - mep.D_T_MAX_24h.ist - mep.D_T_MAX_AbUe.ist - mep.D_T_MAX_AbUeM.ist)
        max_rzu = float(mep.D_T_MAX_Rzu.soll)  
        if max_rab < 0:
            max_rzu -= max_rab
            max_rab = 0
        
        nb_rzu_sum = mep.D_T_NB_ZuS.soll
        if nb_rzu_sum == 0:
            nb_rzu = 0
            nb_rab = 0
        else:
            nb_rab = float(mep.D_T_NB_AbS.soll - mep.D_T_NB_Lab.ist - mep.D_T_NB_24h.ist - mep.D_T_NB_AbUe.ist - mep.D_T_NB_AbUeM.ist)
            nb_rzu = float(mep.D_T_NB_Rzu.soll)  
            if nb_rab < 0:
                nb_rzu -= nb_rab
                nb_rab = 0
            
        tnb_rzu_sum = mep.D_T_TNB_ZuS.soll
        if tnb_rzu_sum == 0:
            tnb_rzu = 0
            tnb_rab = 0
        else:
            tnb_rab = float(mep.D_T_NB_AbS.soll - mep.D_T_NB_Lab.ist - mep.D_T_NB_24h.ist - mep.D_T_NB_AbUe.ist - mep.D_T_NB_AbUeM.ist)
            tnb_rzu = float(mep.D_T_NB_Rzu.soll)  
            if tnb_rab < 0:
                tnb_rzu -= tnb_rab
                tnb_rab = 0

                       
        if len(zu.keys()) == 1:
            for key in zu.keys():
                anzahl = len(zu[key])
                for auslass in zu[key]:
                    if task.max:auslass.Luftmengenmax = round(max_rzu  / anzahl,1)
                    if task.nacht:auslass.Luftmengennacht = round(nb_rzu  / anzahl,1)
                    if task.tnacht:auslass.Luftmengentnacht = round(tnb_rzu / anzahl,1)
                    if task.min:auslass.Luftmengenmin = round(min_rzu / anzahl,1)

        
        elif len(zu.keys()) > 1:
            sum_luft = 0
            for key in zu.keys():
                for auslass in zu[key]:
                    if auslass.Luftmengenmin > 0:
                        sum_luft += auslass.Luftmengenmin
            if sum_luft != 0:
                for key in zu.keys():
                    for auslass in zu[key]:
                        if task.max:auslass.Luftmengenmax = round(max_rzu * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(nb_rzu * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(tnb_rzu * auslass.Luftmengenmin / sum_luft,1)
                        if task.min:auslass.Luftmengenmin = round(min_rzu * auslass.Luftmengenmin / sum_luft,1)
                       
            else:
                anzahlauslass = 0
                for key in zu.keys():
                    anzahlauslass += len(zu[key])
                for key in zu.keys():
                    for auslass in zu[key]:
                        if task.max:auslass.Luftmengenmax = round(max_rzu  / anzahlauslass,1)
                        if task.nacht:auslass.Luftmengennacht = round(nb_rzu  / anzahlauslass,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(tnb_rzu / anzahlauslass,1)
                        if task.min:auslass.Luftmengenmin = round(min_rzu / anzahlauslass,1)


        if len(ab.keys()) == 1:
            for key in ab.keys():
                anzahl = len(ab[key])
                for auslass in ab[key]:
                    if task.max:auslass.Luftmengenmax = round(max_rab / anzahl,1)
                    if task.nacht:auslass.Luftmengennacht = round(nb_rab / anzahl,1)
                    if task.tnacht:auslass.Luftmengentnacht = round(tnb_rab / anzahl,1)
                    if task.min:auslass.Luftmengenmin = round(min_rab / anzahl,1)

                    # if task.max:auslass.Luftmengenmax = round(float(mep.D_T_MAX_Rab.soll)  / anzahl,1)
                    # if task.nacht:auslass.Luftmengennacht = round(float(mep.D_T_NB_Rab.soll)  / anzahl,1)
                    # if task.tnacht:auslass.Luftmengentnacht = round(float(mep.D_T_TNB_Rab.soll)  / anzahl,1)
                    # if task.min:auslass.Luftmengenmin = round(float(mep.D_T_MIN_Rab.soll) / anzahl,1)

        
        elif len(ab.keys()) > 1:
            sum_luft = 0
            for key in ab.keys():
                for auslass in ab[key]:
                    if auslass.Luftmengenmin > 0:
                        sum_luft += auslass.Luftmengenmin
            if sum_luft > 0:
                for key in ab.keys():
                    for auslass in ab[key]:
                        if task.max:auslass.Luftmengenmax = round(max_rab * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(nb_rab * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(tnb_rab * auslass.Luftmengenmin / sum_luft,1)
                        if task.min:auslass.Luftmengenmin = round(min_rab * auslass.Luftmengenmin / sum_luft,1)

            else:
                anzahlauslass = 0
                for key in ab.keys():
                    anzahlauslass += len(ab[key])
                for key in ab.keys():
                    for auslass in ab[key]:
                        if task.max:auslass.Luftmengenmax = round(max_rab / anzahlauslass,1)
                        if task.nacht:auslass.Luftmengennacht = round(nb_rab / anzahlauslass,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(tnb_rab / anzahlauslass,1)
                        if task.min:auslass.Luftmengenmin = round(min_rab / anzahlauslass,1)

        
        for vsr in mep.list_vsr:
            if vsr.art in ['RZU','RAB','LAB','24h','RUM']:
                vsr.Luftmengenermitteln_new()
                vsr.vsrauswaelten()
                vsr.vsrueberpruefen()
        self.AederungUebernehmen(mep)

        del zu
        del ab

    
    def AederungUebernehmen(self,mep):
        mep.werte_schreiben()
        for auslass in mep.list_auslass:
            auslass.wert_schreiben()
        for auslass in mep.Liste_RaumluftUnrelevant:
            auslass.wert_schreiben()
        for vsr in mep.list_vsr:
            vsr.wert_schreiben()
        mep.update_ist_bei_verteilen()
    
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

        doc.Dispose()
        t.Dispose()
        del doc
        del t
        try:task.Dispose()
        except:pass
        finally:del task

    def ExportRaumluftbilanz(self,uiapp):
        for el in self.GUI.ListeMEP:
            el.checked = False
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
        for el in self.GUI.ListeMEP:
            el.checked = False
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

            cell_format_schacht = e.add_format()
            cell_format_schacht.set_align('center')
            cell_format_schacht.set_bold()

            cell_format_anlage = e.add_format()
            cell_format_anlage.set_font_color('red')



            anzahlebene = len(_dict['Alle'].keys()[:])
            anzahl_zeilen = anzahlebene + 5
            alle = _dict['Alle']
            worksheet.write(2,0,'Alle',cell_format_anlage)
            for n_schahct,schacht in enumerate(sorted(raum.liste_schacht)):
                worksheet.merge_range(0,4*n_schahct+1,0,4*n_schahct+4,schacht,cell_format_schacht)
                worksheet.write(0+1, 4*n_schahct+1, 'Raumzuluft')
                worksheet.write(0+1, 4*n_schahct+2, 'Raumabluft')
                worksheet.write(0+1, 4*n_schahct+3, 'Laborabluft')
                worksheet.write(0+1, 4*n_schahct+4, '24h-Abluft')

                for n_ebene,ebene in enumerate(sorted(alle.keys())):
                    worksheet.write(n_ebene+3,0, ebene)
                    for n_schahct,schacht in enumerate(sorted(raum.liste_schacht)):
                        worksheet.write_number(n_ebene+3, 4*n_schahct+1, alle[ebene][schacht][0],cell_format)
                        worksheet.write_number(n_ebene+3, 4*n_schahct+2, alle[ebene][schacht][1],cell_format)
                        worksheet.write_number(n_ebene+3, 4*n_schahct+3, alle[ebene][schacht][2],cell_format)
                        worksheet.write_number(n_ebene+3, 4*n_schahct+4, alle[ebene][schacht][3],cell_format)
                    
                
                worksheet.write(anzahlebene+3, 0, 'Summe')
                for n_schahct,schacht in enumerate(sorted(raum.liste_schacht)):
                    worksheet.write_formula(anzahlebene+3, 4*n_schahct + 1, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(3,4*n_schahct + 1),xlsxwriter.utility.xl_rowcol_to_cell(anzahlebene+2,4*n_schahct + 1)),cell_format)
                    worksheet.write_formula(anzahlebene+3, 4*n_schahct + 2, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(3,4*n_schahct + 2),xlsxwriter.utility.xl_rowcol_to_cell(anzahlebene+2,4*n_schahct + 2)),cell_format)
                    worksheet.write_formula(anzahlebene+3, 4*n_schahct + 3, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(3,4*n_schahct + 3),xlsxwriter.utility.xl_rowcol_to_cell(anzahlebene+2,4*n_schahct + 3)),cell_format)
                    worksheet.write_formula(anzahlebene+3, 4*n_schahct + 4, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(3,4*n_schahct + 4),xlsxwriter.utility.xl_rowcol_to_cell(anzahlebene+2,4*n_schahct + 4)),cell_format)
            del _dict['Alle']
 
            for anlage in sorted(_dict.keys()):
                worksheet.write(anzahl_zeilen + 2,0,'Anlage: ' + str(anlage),cell_format_anlage)

                for n_ebene,ebene in enumerate(sorted(_dict[anlage].keys())):
                    worksheet.write(anzahl_zeilen + 3 + n_ebene,0, ebene)
                    for n_schahct,schacht in enumerate(sorted(raum.liste_schacht)):
                        worksheet.write_number(anzahl_zeilen + 3 + n_ebene, 4*n_schahct+1, _dict[anlage][ebene][schacht][0],cell_format)
                        worksheet.write_number(anzahl_zeilen + 3 + n_ebene, 4*n_schahct+2, _dict[anlage][ebene][schacht][1],cell_format)
                        worksheet.write_number(anzahl_zeilen + 3 + n_ebene, 4*n_schahct+3, _dict[anlage][ebene][schacht][2],cell_format)
                        worksheet.write_number(anzahl_zeilen + 3 + n_ebene, 4*n_schahct+4, _dict[anlage][ebene][schacht][3],cell_format)

                worksheet.write(anzahl_zeilen + 4 + n_ebene, 0, 'Summe')
                for n_schahct,schacht in enumerate(sorted(raum.liste_schacht)):
                    worksheet.write_formula(anzahl_zeilen + 4 + n_ebene, 4*n_schahct + 1, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3,4*n_schahct + 1),xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3 + n_ebene,4*n_schahct + 1)),cell_format)
                    worksheet.write_formula(anzahl_zeilen + 4 + n_ebene, 4*n_schahct + 2, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3,4*n_schahct + 2),xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3 + n_ebene,4*n_schahct + 2)),cell_format)
                    worksheet.write_formula(anzahl_zeilen + 4 + n_ebene, 4*n_schahct + 3, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3,4*n_schahct + 3),xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3 + n_ebene,4*n_schahct + 3)),cell_format)
                    worksheet.write_formula(anzahl_zeilen + 4 + n_ebene, 4*n_schahct + 4, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3,4*n_schahct + 4),xlsxwriter.utility.xl_rowcol_to_cell(anzahl_zeilen + 3 + n_ebene,4*n_schahct + 4)),cell_format)
        
                anzahl_zeilen = anzahl_zeilen + n_ebene + 5
                        # worksheet.write_formula(anzahlebene+4, n_schahct + 1, '=SUM({}:{})'.format(xlsxwriter.utility.xl_rowcol_to_cell((n_anlage +1)*(anzahlebene + 4)+3,n_schahct + 1),xlsxwriter.utility.xl_rowcol_to_cell(n_ebene+(n_anlage +1)*(anzahlebene + 4)+2,n_schahct + 1)),cell_format)
        worksheet.freeze_panes(2, 1)
        e.close() 

        try:temp.Dispose()
        except:pass
        finally:del temp
        del Liste

    def luftmengen_summieren(self,Liste_raum):
        _dict = {}

        Dict_Schacht_RZU = {}
        Dict_Schacht_RAB = {}
        Dict_Schacht_LAB = {}
        Dict_Schacht_24h = {}
        
        for schacht in self.GUI.schachte:
            if schacht.rzu_Schacht.nr and schacht.rzu_Schacht.nr != schacht.schachtname:
                if schacht.rzu_Schacht.nr not in Dict_Schacht_RZU.keys():
                    Dict_Schacht_RZU[schacht.rzu_Schacht.nr] = []
                Dict_Schacht_RZU[schacht.rzu_Schacht.nr].append(schacht.schachtname)
            if schacht.rab_Schacht.nr and schacht.rab_Schacht.nr != schacht.schachtname:
                if schacht.rab_Schacht.nr not in Dict_Schacht_RAB.keys():
                    Dict_Schacht_RAB[schacht.rab_Schacht.nr] = []
                Dict_Schacht_RAB[schacht.rab_Schacht.nr].append(schacht.schachtname)
            if schacht.lab_Schacht.nr and schacht.lab_Schacht.nr != schacht.schachtname:
                if schacht.lab_Schacht.nr not in Dict_Schacht_LAB.keys():
                    Dict_Schacht_LAB[schacht.lab_Schacht.nr] = []
                Dict_Schacht_LAB[schacht.lab_Schacht.nr].append(schacht.schachtname)
            if schacht._24h_Schacht.nr and schacht._24h_Schacht.nr != schacht.schachtname:
                if schacht._24h_Schacht.nr not in Dict_Schacht_24h.keys():
                    Dict_Schacht_24h[schacht._24h_Schacht.nr] = []
                Dict_Schacht_24h[schacht._24h_Schacht.nr].append(schacht.schachtname)
        
        for el in Dict_Schacht_RZU.keys():
            for el1 in Dict_Schacht_RZU.keys():
                if el != el1:
                    if el in Dict_Schacht_RZU[el1]:
                        Dict_Schacht_RZU[el1].extend(Dict_Schacht_RZU[el])
                        Dict_Schacht_RZU[el1].remove(el)
        
        for el in Dict_Schacht_RAB.keys():
            for el1 in Dict_Schacht_RAB.keys():
                if el != el1:
                    if el in Dict_Schacht_RAB[el1]:
                        Dict_Schacht_RAB[el1].extend(Dict_Schacht_RAB[el])
                        Dict_Schacht_RAB[el1].remove(el)
        
        for el in Dict_Schacht_LAB.keys():
            for el1 in Dict_Schacht_LAB.keys():
                if el != el1:
                    if el in Dict_Schacht_LAB[el1]:
                        Dict_Schacht_LAB[el1].extend(Dict_Schacht_LAB[el])
                        Dict_Schacht_LAB[el1].remove(el)
        
        for el in Dict_Schacht_24h.keys():
            for el1 in Dict_Schacht_24h.keys():
                if el != el1:
                    if el in Dict_Schacht_24h[el1]:
                        Dict_Schacht_24h[el1].extend(Dict_Schacht_24h[el])
                        Dict_Schacht_24h[el1].remove(el)

        for raum in Liste_raum:
            anzu = None
            anab = None
            anlab = None
            an24h = None
            for anlage in raum.Anlagen_info:
                
                if anlage.name == 'RZU':anzu = anlage
                if anlage.name == 'RAB':anab = anlage
                if anlage.name == 'LAB':anlab = anlage
                if anlage.name == '24h':an24h = anlage
        
            if 'Alle' not in _dict.keys():
                _dict['Alle'] = {}
            if raum.ebene not in _dict['Alle'].keys():
                _dict['Alle'][raum.ebene] = {}
                for schacht in raum.liste_schacht:
                    _dict['Alle'][raum.ebene][schacht] = [0,0,0,0]
        
            if anzu and anzu.mep_nr:
                if anzu.mep_nr not in _dict.keys():
                    _dict[anzu.mep_nr] = {}
                if raum.ebene not in _dict[anzu.mep_nr].keys():
                    _dict[anzu.mep_nr][raum.ebene] = {}
                    for schacht in raum.liste_schacht:
                        _dict[anzu.mep_nr][raum.ebene][schacht] = [0,0,0,0]
                if raum.rzu_Schacht.nr not in _dict[anzu.mep_nr][raum.ebene].keys():
                    if raum.rzu_Schacht.menge  > 0:logger.error('Zuluftschacht {} von MEP-Raum {} existiert nicht. Bitte korrigieren.'.format(raum.rzu_Schacht.nr,raum.Raumnr))
                    continue
                _dict[anzu.mep_nr][raum.ebene][raum.rzu_Schacht.nr][0] += raum.rzu_Schacht.menge     
                _dict['Alle'][raum.ebene][raum.rzu_Schacht.nr][0] += raum.rzu_Schacht.menge    
            
            if anab and anab.mep_nr:
            
                if anab.mep_nr not in _dict.keys():
                    _dict[anab.mep_nr] = {}
                if raum.ebene not in _dict[anab.mep_nr].keys():
                    _dict[anab.mep_nr][raum.ebene] = {}
                    for schacht in raum.liste_schacht:
                        _dict[anab.mep_nr][raum.ebene][schacht] = [0,0,0,0]
                if raum.rab_Schacht.nr not in _dict[anab.mep_nr][raum.ebene].keys():
                    if raum.rab_Schacht.menge  > 0:logger.error('Raumabluft {} von MEP-Raum {} existiert nicht. Bitte korrigieren.'.format(raum.rab_Schacht.nr,raum.Raumnr))
                    continue
                _dict[anab.mep_nr][raum.ebene][raum.rab_Schacht.nr][1] += raum.rab_Schacht.menge  
                _dict['Alle'][raum.ebene][raum.rab_Schacht.nr][1] += raum.rab_Schacht.menge          
            if anlab and anlab.mep_nr:
            
                if anlab.mep_nr not in _dict.keys():
                    _dict[anlab.mep_nr] = {}
                if raum.ebene not in _dict[anlab.mep_nr].keys():
                    _dict[anlab.mep_nr][raum.ebene] = {}
                    for schacht in raum.liste_schacht:
                        _dict[anlab.mep_nr][raum.ebene][schacht] = [0,0,0,0]
                if raum.lab_Schacht.nr not in _dict[anlab.mep_nr][raum.ebene].keys():
                    if raum.lab_Schacht.menge  > 0:logger.error('Laborabluft {} von MEP-Raum {} existiert nicht. Bitte korrigieren.'.format(raum.lab_Schacht.nr,raum.Raumnr))
                    continue
                _dict[anlab.mep_nr][raum.ebene][raum.lab_Schacht.nr][2] += raum.lab_Schacht.menge   
                _dict['Alle'][raum.ebene][raum.lab_Schacht.nr][2] += raum.lab_Schacht.menge        
            if an24h and an24h.mep_nr:

                if an24h.mep_nr not in _dict.keys():
                    _dict[an24h.mep_nr] = {}
                if raum.ebene not in _dict[an24h.mep_nr].keys():
                    _dict[an24h.mep_nr][raum.ebene] = {}
                    for schacht in raum.liste_schacht:
                        _dict[an24h.mep_nr][raum.ebene][schacht] = [0,0,0,0]
                if raum._24h_Schacht.nr not in _dict[an24h.mep_nr][raum.ebene].keys():
                    if raum._24h_Schacht.menge  > 0:logger.error('24h-Abluftschacht {} von MEP-Raum {} existiert nicht. Bitte korrigieren.'.format(raum._24h_Schacht.nr,raum.Raumnr))
                    continue
                _dict[an24h.mep_nr][raum.ebene][raum._24h_Schacht.nr][3] += raum._24h_Schacht.menge 
                _dict['Alle'][raum.ebene][raum._24h_Schacht.nr][3] += raum._24h_Schacht.menge 
            
        for anlage in _dict.keys():
            for ebene in _dict[anlage].keys():
                for schacht in _dict[anlage][ebene].keys():
                    if schacht in Dict_Schacht_RZU.keys():
                        sub_schaechte = Dict_Schacht_RZU[schacht]
                        for sub in sub_schaechte:
                            if sub not in _dict[anlage][ebene].keys():
                                continue
                            _dict[anlage][ebene][schacht][0] += _dict[anlage][ebene][sub][0]
                    if schacht in Dict_Schacht_RAB.keys():
                        sub_schaechte = Dict_Schacht_RAB[schacht]
                        for sub in sub_schaechte:
                            if sub not in _dict[anlage][ebene].keys():
                                continue
                            _dict[anlage][ebene][schacht][1] += _dict[anlage][ebene][sub][1]
                    if schacht in Dict_Schacht_LAB.keys():
                        sub_schaechte = Dict_Schacht_LAB[schacht]
                        for sub in sub_schaechte:
                            if sub not in _dict[anlage][ebene].keys():
                                continue
                            _dict[anlage][ebene][schacht][2] += _dict[anlage][ebene][sub][2]
                    if schacht in Dict_Schacht_24h.keys():
                        sub_schaechte = Dict_Schacht_24h[schacht]
                        for sub in sub_schaechte:
                            if sub not in _dict[anlage][ebene].keys():
                                continue
                            _dict[anlage][ebene][schacht][3] += _dict[anlage][ebene][sub][3]
               

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
