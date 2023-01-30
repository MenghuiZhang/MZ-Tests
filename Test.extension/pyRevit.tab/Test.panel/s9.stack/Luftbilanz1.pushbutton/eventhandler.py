# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,TaskDialog,UIView,RevitCommandId,UIApplication,PostableCommand,TaskDialogResult,TaskDialogCommonButtons,ExternalEvent
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility,GridLength
import System
import xlsxwriter
from rpw import revit,DB,UI
from pyrevit import script,forms
from System.Windows.Input import ModifierKeys,Keyboard,Key
from System.Collections.Generic import List
from System.Windows.Media import Brushes
from clr import GetClrType
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext
import os
from System.Windows.Forms import FolderBrowserDialog,DialogResult

IGF_LOGO = 'C:\Users\Zhang\Documents\GitHub\MZ-Tools\Test.extension\pyRevit.tab\Test.panel\s9.stack\Luftbilanz1.pushbutton\icon.png'

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'Raumluftbilanz')

Liste_Datenbank = {}
Liste_Datenbank1 = {}
datenbankadresse = r'R:\pyRevit\03_Lüftung\02_Luftungsbilanz\IGF_RLT_Volumenstromregler_Datenbank.xlsx'

def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if not param:return ''
    if param.StorageType.ToString() == 'ElementId':
        return param.AsValueString()
    elif param.StorageType.ToString() == 'Integer':
        value = param.AsInteger()
    elif param.StorageType.ToString() == 'Double':
        value = param.AsDouble()
    elif param.StorageType.ToString() == 'String':
        value = param.AsString()
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value

class VSRHerstellerDaten:
    def __init__(self,fabrikat,typ,A,B,D,spannung,nennstrom,leistung,vmin,vmax,vnenn,bemerkung):
        self.fabrikat = fabrikat
        self.typ = typ
        self.A = A
        self.B = B
        self.D = D
        self.spannung = spannung
        self.nennstrom = nennstrom
        self.leistung = leistung
        self.vmin = vmin
        self.vmax = vmax
        self.vnenn = vnenn
        self.bemerkung = bemerkung
        if self.D:
            self.dimension = 'DN '+str(D)
        else:
            if self.A and self.B:
                self.dimension = str(self.A)+'x'+str(self.B)
            else:
                self.dimension = ''


class Luftauslassfuerexport:
    def __init__(self,art = '',slavevon = '',raumnummer = '',auslassart = '',nutzung = '', Luftmengenmin = '',Luftmengenmax = '',Luftmengennacht = '',Luftmengentnacht = '',bemerkung = '',anmerkung = '',):
        self.art = art
        self.slavevon = slavevon
        self.raumnummer = raumnummer
        self.auslassart = auslassart
        self.nutzung = nutzung
        self.Luftmengenmin = Luftmengenmin
        self.Luftmengenmax = Luftmengenmax
        self.Luftmengennacht = Luftmengennacht
        self.Luftmengentnacht = Luftmengentnacht
        self.bemerkung = bemerkung
        self.anmerkung = anmerkung

if os.path.exists(datenbankadresse):
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    fs = FileStream(datenbankadresse,FileMode.Open,FileAccess.Read)
    book = ExcelPackage(fs)
    
    for sheet in book.Workbook.Worksheets:
        try:
            fabrikat = sheet.Name
            Liste_Datenbank[fabrikat] = {}
            Liste_Datenbank1[fabrikat] = {}
            maxRowNum = sheet.Dimension.End.Row
            for row in range(3,maxRowNum+1):
                luftart = 0 # 0:Zu/Abluft,1:Zuluft,2:Abluft,3:PPS
                vsrart = 'VVR'
                art = sheet.Cells[row, 1].Value
                A = sheet.Cells[row, 2].Value
                B = sheet.Cells[row, 3].Value
                D = sheet.Cells[row, 4].Value
                Bezeichnung = sheet.Cells[row, 5].Value
                Typ = sheet.Cells[row, 7].Value
                Spannung = sheet.Cells[row, 8].Value
                Nennstrom = sheet.Cells[row, 9].Value
                Leistung = sheet.Cells[row, 10].Value
                vmin = sheet.Cells[row, 11].Value
                vmax = sheet.Cells[row, 12].Value
                vnenn = sheet.Cells[row, 13].Value
                Bemerkung = sheet.Cells[row, 16].Value
                
                if any([A,B,D]):
                    if Bemerkung and Bemerkung.upper().find('PPS') != -1:
                        luftart = 3
                    else:
                        if art:
                            if art.upper().find('ZU') != -1:
                                luftart = 1
                            elif art.upper().find('AB') != -1:
                                luftart = 2
                    if Bemerkung:
                        if Bemerkung.upper().find('VARI') == -1:
                            vsrart = 'KVR'
                    if A and B:
                        mini = min([int(A),int(B)])
                        maxi = max([int(A),int(B)])
                        daten = VSRHerstellerDaten(fabrikat,Typ,mini,maxi,None,Spannung,Nennstrom,Leistung,vmin,vmax,vnenn,Bemerkung)
                        if luftart not in Liste_Datenbank1[fabrikat].keys():
                            Liste_Datenbank1[fabrikat][luftart] = {}
                        if vsrart not in Liste_Datenbank1[fabrikat][luftart].keys():
                            Liste_Datenbank1[fabrikat][luftart][vsrart] = {}
                        if maxi not in Liste_Datenbank1[fabrikat][luftart][vsrart].keys():
                            Liste_Datenbank1[fabrikat][luftart][vsrart][mini] = {}
                        if mini not in Liste_Datenbank1[fabrikat][luftart][vsrart][maxi].keys():
                            Liste_Datenbank1[fabrikat][luftart][vsrart][maxi][mini] = []
                        Liste_Datenbank1[fabrikat][luftart][vsrart][maxi][mini].append(daten)

                    else:
                        daten = VSRHerstellerDaten(fabrikat,Typ,None,None,int(D),Spannung,Nennstrom,Leistung,vmin,vmax,vnenn,Bemerkung)
                        if luftart not in Liste_Datenbank[fabrikat].keys():
                            Liste_Datenbank[fabrikat][luftart] = {}
                        if vsrart not in Liste_Datenbank[fabrikat][luftart].keys():
                            Liste_Datenbank[fabrikat][luftart][vsrart] = {}
                        if int(D) not in Liste_Datenbank[fabrikat][luftart][vsrart].keys():
                            Liste_Datenbank[fabrikat][luftart][vsrart][int(D)] = []
                        Liste_Datenbank[fabrikat][luftart][vsrart][int(D)].append(daten)

        except Exception as e:
            logger.error(e)
            TaskDialog.Show('Fehler','Fehler beim Lesen der Excel-Datei')        
            break   

        book.Save()
        book.Dispose()
        fs.Dispose()

class Raumdaten:
    def __init__(self,raum,row = 0):
        self.bemerkung = ''
        self.anmerkung = ''
        self.row = row
        self.rowende = row + 20
        self.book = None
        self.sheet = None
        self.raum = raum
        self.raumnummer = self.raum.elem.Number
        self.raumname = self.raum.elem.LookupParameter('Name').AsString()
        self.ebene = self.raum.ebene
        self.rowbreaks = []

        self.flaeche = self.raum.flaeche
        self.personen = self.raum.personen
        self.hoehe = self.raum.hoehe/1000.0
        self.volumen = self.raum.volumen

        self.meplabmin = self.raum.ab_lab_min.soll     
        self.meplabmax = self.raum.ab_lab_max.soll
        self.mepab24h = self.raum.ab_24h.soll
        self.mepdruck = self.raum.Druckstufe.soll
        self.mepueber = self.raum.ueber_in_manuell.soll + self.raum.ueber_in.soll - self.raum.ueber_aus.soll - self.raum.ueber_aus_manuell.soll 
        self.mepnachtdruck = self.mepdruck
        self.meptnachtdruck = self.mepdruck
        self.mepnacht24h = self.mepab24h
        self.meptnacht24h = self.mepab24h

        self.bezugsname = self.raum.bezugsname
        self.faktor = self.raum.faktor
        self.einheit = self.raum.einheit
        self.mepminzu = self.raum.zu_min.soll
        self.mepmaxzu = self.raum.zu_max.soll
        self.mepminab = self.raum.ab_min.soll
        self.mepmaxab = self.raum.ab_max.soll
        self.mepnachtueber = self.mepueber
        self.meptnachtueber = self.mepueber
        # self.mepminab = self.raum.ab_min.soll + self.raum.ab_24h.soll + self.raum.ab_lab_min.soll
        # self.mepmaxab = self.raum.ab_max.soll + self.raum.ab_24h.soll + self.raum.ab_lab_max.soll

        self.isnachtbetrieb = 'Ja' if self.raum.nachtbetrieb  else 'Nein'
        self.istiefenachtbetrieb = 'Ja' if self.raum.tiefenachtbetrieb  else 'Nein'
        if self.isnachtbetrieb == 'Nein':
            self.LW_nacht = '-'
            self.mepnb_zu = '-'
            self.mepnb_ab = '-'
            self.meplabnacht = '-'
        else:
            self.LW_nacht = self.raum.NB_LW
            self.mepnb_zu = self.raum.nb_zu.soll
            self.meplabnacht = self.meplabmin
            self.mepnb_ab = self.raum.nb_ab.soll
            
        
        if self.istiefenachtbetrieb == 'Nein':
            self.tLW_nacht = '-'
            self.meptnb_zu = '-'
            self.meptnb_ab = '-'
            self.meplabtnacht = '-'
            self.meptnacht24h = '-'
        else:
            self.tLW_nacht = self.raum.T_NB_LW
            self.meptnb_zu = self.raum.tnb_zu.soll
            self.meplabtnacht = self.meplabmin
            self.meptnb_ab = self.raum.tnb_ab.soll

        # self.istiefenachtbetrieb = 'Ja' if self.raum.tiefenachtbetrieb  else 'Nein'
        # self.LW_tnacht = self.raum.T_NB_LW
        # self.meptnb_zu = self.raum.tnb_zu.soll
        # self.meptnb_ab = self.raum.tnb_ab.soll + self.raum.ab_24h.soll + self.raum.ab_lab_min.soll

        self.istminzu = self.raum.zu_min.ist
        self.istmaxzu = self.raum.zu_max.ist
        self.istnachtzu = self.raum.nb_zu.ist
        self.isttnachtzu = self.raum.tnb_zu.ist
        
        self.istminab = self.raum.ab_min.ist
        self.istmaxab = self.raum.ab_max.ist
        self.istnachtab = self.raum.nb_ab.ist
        self.isttnachtab = self.raum.tnb_ab.ist

        self.istminlab = self.raum.ab_lab_min.ist
        self.istmaxlab = self.raum.ab_lab_max.ist
        self.istnachtlab = self.raum.labnacht
        self.isttnachtlab = self.raum.labtnacht

        self.istmin24h = self.raum.ab_24h.ist
        self.istmax24h = self.raum.ab_24h.ist
        self.istnacht24h = self.raum.ab24nacht
        self.isttnacht24h = self.raum.ab24tnacht

        self.istueber = self.raum.ueber_in_manuell.ist + self.raum.ueber_in.ist - self.raum.ueber_aus.ist - self.raum.ueber_aus_manuell.ist 
        self.istminueber = self.istueber
        self.istmaxueber = self.istueber
        self.istnachtueber = self.istueber
        self.isttnachtueber = self.istueber

        self.istminsumme = self.istminzu - self.istminab - self.istminlab - self.istmin24h + self.istminueber
        self.istmaxsumme = self.istmaxzu - self.istmaxab - self.istmaxlab - self.istmax24h + self.istmaxueber
        self.istmin = 'OK' if abs(self.istminsumme-self.mepdruck) < 3 else 'Passt nicht'
        self.istmax = 'OK' if abs(self.istmaxsumme-self.mepdruck) < 3 else 'Passt nicht'

        self.mepminsumme = self.mepminzu - self.mepminab - self.meplabmin - self.mepab24h + self.mepueber
        self.mepmaxsumme = self.mepmaxzu - self.mepmaxab - self.meplabmax - self.mepab24h + self.mepueber
        self.mepmin = 'OK' if abs(self.mepminsumme-self.mepdruck) < 3 else 'Passt nicht'
        self.mepmax = 'OK' if abs(self.mepmaxsumme-self.mepdruck) < 3 else 'Passt nicht'

        if self.isnachtbetrieb == 'Nein':
            self.istnachtzu = '-'
            self.istnachtab = '-'
            self.istnachtsumme = '-'
            self.istnachtab = '-'
            self.istnachtueber = '-'
            self.istnachtlab = '-'
            self.istnacht = '-'

            self.mepnachtdruck = '-'
            
            self.mepnachtzu = '-'
            self.mepnachtab = '-'
            self.mepnachtsumme = '-'
            self.mepnachtab = '-'
            self.mepnachtueber = '-'
            self.mepnacht = '-'
            if self.istnacht24h == 0:
                self.istnacht24h = '-'
                self.mepnacht24h = '-'
            else:
                self.istnacht = 'Passt nicht'
                self.mepnacht = 'Passt nicht'
        else:
            self.istnachtsumme = self.istnachtzu - self.istnachtab - self.istnachtlab - self.istnacht24h + self.istnachtueber
            self.istnacht = 'OK' if abs(self.istnachtsumme-self.mepnachtdruck) < 3 else 'Passt nicht'
            self.mepnachtsumme = self.mepnb_zu - self.mepnb_ab - self.meplabnacht - self.mepnacht24h + self.mepnachtueber
            self.mepnacht = 'OK' if abs(self.mepnachtsumme-self.mepnachtdruck) < 3 else 'Passt nicht'


        if self.istiefenachtbetrieb == 'Nein':
            self.isttnachtzu = '-'
            self.isttnachtab = '-'
            self.isttnachtsumme = '-'
            self.isttnachtab = '-'
            self.isttnachtueber = '-'
            self.isttnacht = '-'
            self.isttnachtlab = '-'

            self.meptnachtzu = '-'
            self.meptnachtab = '-'
            self.meptnachtsumme = '-'
            self.meptnachtab = '-'
            self.meptnachtueber = '-'
            self.meptnachtdruck = '-'
            self.meptnacht = '-'
            if self.isttnacht24h == 0:
                self.isttnacht24h = '-'
                self.meptnacht24h = '-'
            else:
                self.isttnacht = 'Passt nicht'
                self.meptnacht = 'Passt nicht'
        else:
            self.isttnachtsumme = self.isttnachtzu - self.isttnachtab - self.isttnachtlab - self.isttnacht24h + self.isttnachtueber
            self.isttnacht = 'OK' if abs(self.isttnachtsumme-self.meptnachtdruck) < 3 else 'Passt nicht'
            self.meptnachtsumme = self.meptnb_zu - self.meptnb_ab - self.meplabtnacht - self.meptnacht24h + self.meptnachtueber
            self.meptnacht = 'OK' if abs(self.meptnachtsumme-self.meptnachtdruck) < 3 else 'Passt nicht'
                 
    def cellformat(self,align = 'left',Bold = True, font_size = 12, num_format = None,background = None,left=None,right=None,buttom = None,top = None,font_ground = None,textwrap=False,valigh = 3):
        '''1: Top, 2: center, 3: bottom'''
        cell_format = self.book.add_format()
        cell_format.set_text_v_align(valigh)
        cell_format.set_align(align)
        if Bold:cell_format.set_bold()
        cell_format.set_font_size(font_size)
        if num_format:cell_format.set_num_format(num_format)
        if background:
            cell_format.set_bg_color(background)
        if font_ground:
            cell_format.set_font_color(font_ground)

        if left:cell_format.set_left(left)
        if right:cell_format.set_right(right)
        if buttom:cell_format.set_bottom(buttom)
        if top:cell_format.set_top(top)
        if textwrap:cell_format.set_text_wrap()
        return cell_format

    def exportheader(self):
        if not self.sheet:
            return
        self.sheet.set_row(self.row,22)
        self.sheet.merge_range(self.row,0,self.row,5,'Projekt: ' + name,self.cellformat('center',font_size=16,left = 2,right=1,buttom=2,top=2))
        self.sheet.merge_range(self.row,6,self.row,12,'Projektnummer: '+number,self.cellformat('center',font_size=16,buttom=2,top=2,right=1))
        self.sheet.merge_range(self.row,13,self.row,19,'Gewerk: Lüftung',self.cellformat('center',font_size=16,buttom=2,top=2,right=2))
        self.sheet.write(self.row, 20, '',self.cellformat('center',False,10,buttom=2,right=2,top=2))

        self.sheet.merge_range(self.row+1,0,self.row+1,14,'Raumdaten aus den Raumbuch',self.cellformat('center',left = 2,right=2,buttom=1))
        self.sheet.merge_range(self.row+1,15,self.row+1,18,'Luftsumme der Luftauslässe',self.cellformat('center',right=2,buttom=1))
        self.sheet.write(self.row+1, 19, 'Bemerkung Raum',self.cellformat('center',left=1,buttom=1,right=2))
        self.sheet.write(self.row+1, 20, 'Anmerkung Raum',self.cellformat('center',buttom=1,right=2))

        self.sheet.merge_range(self.row+2,0,self.row+2,8,'Raumdaten',self.cellformat('center',False,10,left=2,buttom=1,right=1))
        self.sheet.merge_range(self.row+2, 9,self.row+2, 10, '',self.cellformat('center',False,10,left=1,buttom=1,right=1))
        self.sheet.merge_range(self.row+2,11,self.row+2,12,'Tagbetrieb',self.cellformat('center',False,10,left=1,buttom=1,right=1))
        self.sheet.merge_range(self.row+2, 13, self.row+3, 13,'Nacht- \nbetrieb',self.cellformat('center',False,10,left=1,buttom=1,right=1,textwrap=True,valigh=2))
        self.sheet.merge_range(self.row+2, 14, self.row+3, 14,'abgesenkter \nBetrieb',self.cellformat('center',False,10,left=1,buttom=1,right=2,textwrap=True,valigh=2))
        self.sheet.merge_range(self.row+2,15,self.row+2,16,'Tagbetrieb',self.cellformat('center',False,10,left=1,buttom=1,right=1))
        self.sheet.merge_range(self.row+2, 17,self.row+3, 17, 'Nacht- \nbetrieb',self.cellformat('center',False,10,left=1,buttom=1,right=1,textwrap=True,valigh=2))
        self.sheet.merge_range(self.row+2, 18,self.row+3, 18, 'abgesenkter \nBetrieb',self.cellformat('center',False,10,left=1,buttom=1,right=2,textwrap=True,valigh=2))
        self.sheet.merge_range(self.row+2, 19,self.row+3, 19, '',self.cellformat('center',False,10,left=1,buttom=1,right=2))
        self.sheet.merge_range(self.row+2, 20,self.row+3, 20, '',self.cellformat('center',False,10,buttom=1,right=2))

        self.sheet.merge_range(self.row+3,9,self.row+3,10,'',self.cellformat('center',left=1,buttom=1,right=1))
        self.sheet.write(self.row+3, 11, 'min',self.cellformat('center',False,10,left=1,buttom=1,right=1))
        self.sheet.write(self.row+3, 12, 'max',self.cellformat('center',False,10,left=1,buttom=1,right=1))
        self.sheet.write(self.row+3, 15, 'min',self.cellformat('center',False,10,left=1,buttom=1,right=1))
        self.sheet.write(self.row+3, 16, 'max',self.cellformat('center',False,10,left=1,buttom=1,right=1))

        self.sheet.write(self.row+3, 0,'Nummer',self.cellformat('left',False,10,right=1,left=2))
        self.sheet.write(self.row+4, 0,'Name',self.cellformat('left',False,10,right=1,left=2))
        self.sheet.write(self.row+5, 0,'Ebene',self.cellformat('left',False,10,right=1,left=2))
        self.sheet.write(self.row+6, 0,'Personen',self.cellformat('left',False,10,right=1,left=2))
        self.sheet.write(self.row+7, 0,'Fläche',self.cellformat('left',False,10,right=1,left=2))
        self.sheet.write(self.row+8, 0,'Höhe',self.cellformat('left',False,10,right=1,left=2))
        self.sheet.write(self.row+9, 0,'Volumen',self.cellformat('left',False,10,right=1,left=2,buttom=1))
        self.sheet.merge_range(self.row+10, 0,self.row+11, 4,'',self.cellformat('left',False,10,right=1,left=2,buttom=1))

        self.sheet.merge_range(self.row+3, 1, self.row+3, 4, self.raumnummer,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+4, 1, self.row+4, 4,self.raumname,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+5, 1, self.row+5, 4,self.ebene,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+6, 1, self.row+6, 4,self.personen,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+7, 1, self.row+7, 4,self.flaeche,self.cellformat('left',False,10,right=1,num_format='''0.00 "m²"'''))
        self.sheet.merge_range(self.row+8, 1, self.row+8, 4,self.hoehe,self.cellformat('left',False,10,right=1,num_format='''0.00 "m"'''))
        self.sheet.merge_range(self.row+9, 1, self.row+9, 4,self.volumen,self.cellformat('left',False,10,right=1,buttom=1,num_format='''0.00 "m³"'''))

        self.sheet.merge_range(self.row+3, 5,self.row+3, 8,'Tagbetrieb',self.cellformat('center',False,10,right=1,buttom=1))
        self.sheet.merge_range(self.row+4, 5,self.row+4, 6,'Berechnung nach:',self.cellformat('left',False,10,right=4))
        self.sheet.merge_range(self.row+4, 7,self.row+4, 8,self.bezugsname,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+5, 5,self.row+5, 6,'Faktor:',self.cellformat('left',False,10,right=4))
        self.sheet.merge_range(self.row+5, 7,self.row+5, 8,self.faktor,self.cellformat('left',False,10,right=1,num_format='''0 "{}"'''.format(self.einheit)))

        self.sheet.merge_range(self.row+6, 5,self.row+6, 8,'Nachtbetrieb',self.cellformat('center',False,10,right=1,top=1,buttom=1))
        self.sheet.merge_range(self.row+7, 5,self.row+7, 6,'ist Nachtbetrieb',self.cellformat('left',False,10,right=4))
        self.sheet.merge_range(self.row+7, 7,self.row+7, 8,self.isnachtbetrieb,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+8, 5,self.row+8, 6,'Faktor',self.cellformat('left',False,10,right=4))
        self.sheet.merge_range(self.row+8, 7,self.row+8, 8,self.LW_nacht,self.cellformat('left',False,10,right=1,num_format='''0 "-1/h"'''))

        self.sheet.merge_range(self.row+9, 5,self.row+9, 8,'abgesenkter Betrieb',self.cellformat('center',False,10,right=1,top=1,buttom=1))
        self.sheet.merge_range(self.row+10, 5,self.row+10,6,'ist angesenkter Betrieb:',self.cellformat('left',False,10,right=4))
        self.sheet.merge_range(self.row+10, 7,self.row+10, 8,self.istiefenachtbetrieb,self.cellformat('left',False,10,right=1))
        self.sheet.merge_range(self.row+11, 5,self.row+11, 6,'Faktor:',self.cellformat('left',False,10,right=4,buttom=1))
        self.sheet.merge_range(self.row+11, 7,self.row+11, 8,self.tLW_nacht,self.cellformat('left',False,10,right=1,num_format='''0 "-1/h"''',buttom=1))

        self.sheet.merge_range(self.row+4, 9, self.row+4, 10,'Raum Zuluft',self.cellformat('left',False,10,background='#A2E0F4',right=1))
        self.sheet.merge_range(self.row+5, 9, self.row+5, 10, 'Raum Abluft',self.cellformat('left',False,10,background='#F1EF8D',right=1))
        self.sheet.merge_range(self.row+6, 9, self.row+6, 10, 'Labor Abluft',self.cellformat('left',False,10,background='#F1EF8D',right=1))
        self.sheet.merge_range(self.row+7, 9, self.row+7, 10, '24h-Abluft',self.cellformat('left',False,10,background='#FFD757',right=1))
        self.sheet.merge_range(self.row+8, 9, self.row+8, 10, 'Übertrömung',self.cellformat('left',False,10,background='#BFBFBF',right=1))
        self.sheet.merge_range(self.row+9, 9, self.row+9, 10, 'Druckstufe',self.cellformat('left',False,10,background='#ADD8E6',right=1))
        self.sheet.merge_range(self.row+10, 9, self.row+10, 10, 'Luftbilanz',self.cellformat('left',False,10,right=1,buttom=1))
        self.sheet.merge_range(self.row+11, 9, self.row+11, 10, 'Überprüfung',self.cellformat('left',True,10,buttom=1,right=1))       

        self.sheet.write(self.row+4, 11, self.mepminzu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+5, 11, self.mepminab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+6, 11, self.meplabmin,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+7, 11, self.mepab24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+8, 11, self.mepueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+9, 11, self.mepdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+10, 11, self.mepminsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=1,buttom=1))
        if self.mepmin == 'OK':self.sheet.write(self.row+11, 11, self.mepmin,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=1))
        else:self.sheet.write(self.row+11, 11, self.mepmin,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=1))

        self.sheet.write(self.row+4, 12, self.mepmaxzu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+5, 12, self.mepmaxab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+6, 12, self.meplabmax,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+7, 12, self.mepab24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+8, 12, self.mepueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+9, 12, self.mepdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+10, 12, self.mepmaxsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=1,buttom=1))
        if self.mepmax == 'OK':self.sheet.write(self.row+11, 12, self.mepmax,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=1))
        else:self.sheet.write(self.row+11, 12, self.mepmax,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=1))

        self.sheet.write(self.row+4, 13, self.mepnb_zu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+5, 13, self.mepnb_ab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+6, 13, self.meplabnacht,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+7, 13, self.mepnacht24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+8, 13, self.mepnachtueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+9, 13, self.mepnachtdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+10, 13, self.mepnachtsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=1,buttom=1))
        if self.mepnacht == 'OK':self.sheet.write(self.row+11, 13, self.mepnacht,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=1))
        elif self.mepnacht == 'Passt nicht':self.sheet.write(self.row+11, 13, self.mepnacht,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=1))
        else:self.sheet.write(self.row+11, 13, self.mepnacht,self.cellformat('center',True,10,buttom=1,right=1))

        self.sheet.write(self.row+4, 14, self.meptnb_zu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+5, 14, self.meptnb_ab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+6, 14, self.meplabtnacht,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+7, 14, self.meptnacht24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+8, 14, self.meptnachtueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+9, 14, self.meptnachtdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+10, 14, self.meptnachtsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=2,buttom=1))
        if self.meptnacht == 'OK':self.sheet.write(self.row+11, 14, self.meptnacht,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=2))
        elif self.meptnacht == 'Passt nicht':self.sheet.write(self.row+11, 14, self.meptnacht,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=2))
        else:self.sheet.write(self.row+11, 14, self.meptnacht,self.cellformat('center',True,10,buttom=1,right=2))

        self.sheet.write(self.row+4, 15, self.istminzu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+5, 15, self.istminab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+6, 15, self.istminlab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+7, 15, self.istmin24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+8, 15, self.istminueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+9, 15, self.mepdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+10, 15, self.istminsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=1,buttom=1))
        if self.istmin == 'OK':self.sheet.write(self.row+11, 15, self.istmin,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=1))
        else:self.sheet.write(self.row+11, 15, self.istmin,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=1))

        self.sheet.write(self.row+4, 16, self.istmaxzu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+5, 16, self.istmaxab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+6, 16, self.istmaxlab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+7, 16, self.istmax24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+8, 16, self.istmaxueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+9, 16, self.mepdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+10, 16, self.istmaxsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=1,buttom=1))
        if self.istmax == 'OK':self.sheet.write(self.row+11, 16, self.istmax,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=1))
        else:self.sheet.write(self.row+11, 16, self.istmax,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=1))

        self.sheet.write(self.row+4, 17, self.istnachtzu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+5, 17, self.istnachtab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+6, 17, self.istnachtlab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+7, 17, self.istnacht24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+8, 17, self.istnachtueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+9, 17, self.mepnachtdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=1))
        self.sheet.write(self.row+10, 17, self.istnachtsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=1,buttom=1))
        if self.istnacht == 'OK':self.sheet.write(self.row+11, 17, self.istnacht,self.cellformat('center',True,10,buttom=1,font_ground='#17A241',right=1))
        elif self.istnacht == 'Passt nicht':self.sheet.write(self.row+11, 17, self.istnacht,self.cellformat('center',True,10,buttom=1,font_ground = '#FF0000',right=1))
        else:self.sheet.write(self.row+11, 17, self.istnacht,self.cellformat('center',True,10,buttom=1,right=1))

        self.sheet.write(self.row+4, 18, self.isttnachtzu,self.cellformat('right',False,10,background='#A2E0F4',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+5, 18, self.isttnachtab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+6, 18, self.isttnachtlab,self.cellformat('right',False,10,background='#F1EF8D',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+7, 18, self.isttnacht24h,self.cellformat('right',False,10,background='#FFD757',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+8, 18, self.isttnachtueber,self.cellformat('right',False,10,background='#BFBFBF',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+9, 18, self.meptnachtdruck,self.cellformat('right',False,10,background='#ADD8E6',num_format='''0 "m³/h"''',right=2))
        self.sheet.write(self.row+10, 18, self.isttnachtsumme,self.cellformat('right',False,10,num_format='''0 "m³/h"''',right=2,buttom=1))
        if self.isttnacht == 'OK':self.sheet.write(self.row+11, 18, self.isttnacht,self.cellformat('center',True,10,buttom=1,font_ground='#16A241',right=2))
        elif self.isttnacht == 'Passt nicht':self.sheet.write(self.row+11, 18, self.isttnacht,self.cellformat('center',True,10,buttom=1,font_ground='#FF0000',right=2))
        else:self.sheet.write(self.row+11, 18, self.isttnacht,self.cellformat('center',True,10,buttom=1,right=2))

        self.sheet.merge_range(self.row+4,19,self.row+11,19,self.bemerkung,self.cellformat('center',False,10,left=1,buttom=1,right=2,textwrap=True))
        self.sheet.merge_range(self.row+4,20,self.row+11,20,self.anmerkung,self.cellformat('center',False,10,buttom=1,right=2,textwrap=True))

    def create_art(self,row,Text,background = None):
        self.sheet.merge_range(row, 0, row, 19,Text,self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background=background))
        self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background=background))

    def create_last(self,row):
        self.sheet.merge_range(row, 0, row, 20,'',self.cellformat('left',True,10,top=2))
        
    def create_neue_ueberschrift(self,n):
        # if n >= 36:
        #     n = n-36
        #     self.row = self.row+36
        #     self.exportueberschrift1(self.row)
        #     n = n+3
        return n

    def exportueberschrift(self):
        if not self.sheet:
            return
        self.sheet.set_row(self.row+12,30)

        self.sheet.write(self.row+12, 0, 'VSR Id',self.cellformat('center',True,10,left=2,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 1, 'Slave von',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 2, 'BKS-Schild',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 3, 'Einbauort',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 4, 'Bezeichnung',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 5, 'Fabrikat',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.merge_range(self.row+12, 6,self.row+12, 7, 'Typ[Nutzung]',self.cellformat('center',True,10,buttom=1,top=6,right=1))
        self.sheet.write(self.row+12, 8, 'Spann- \nung',self.cellformat('center',True,10,buttom=1,top=6,right=4,textwrap=True))
        self.sheet.write(self.row+12, 9, 'Nenn- \nstrom',self.cellformat('center',True,10,buttom=1,top=6,right=4,textwrap=True))
        self.sheet.write(self.row+12, 10, 'Leist- \nung',self.cellformat('center',True,10,buttom=1,top=6,right=1,textwrap=True))
        self.sheet.merge_range(self.row+12, 11, self.row+12, 12,'Tagbetrieb',self.cellformat('center',True,10,buttom=1,top = 6,right=4))
        self.sheet.write(self.row+12, 13, 'Nacht- \nbetrieb',self.cellformat('center',True,10,buttom=1,top = 6,right=4,textwrap=True))
        self.sheet.write(self.row+12, 14, 'abgesenkter \nBetrieb',self.cellformat('center',True,10,buttom=1,top = 6,right=1,textwrap=True))
        self.sheet.write(self.row+12, 15, 'Regler \nmin.',self.cellformat('center',True,10,buttom=1,top = 6,right=4,textwrap=True))
        self.sheet.write(self.row+12, 16, 'max. \ngewünscht',self.cellformat('center',True,10,buttom=1,top = 6,right=4,textwrap=True))
        self.sheet.write(self.row+12, 17, 'Regler \nVnenn ',self.cellformat('center',True,10,buttom=1,top = 6,right=1,textwrap=True))
        self.sheet.merge_range(self.row+12, 18, self.row+12, 19,'Bemerkung Antrieb',self.cellformat('center',True,10,top = 6,buttom=1,right=2))
        self.sheet.write(self.row+12, 20, 'Anmerkung',self.cellformat('center',True,10,right=2,top = 6,buttom=1))

        self.sheet.merge_range(self.row+13, 0, self.row+14, 0, '',self.cellformat(left=2,right=4))
        self.sheet.merge_range(self.row+13, 1, self.row+14, 1, '',self.cellformat(right=4))
        self.sheet.merge_range(self.row+13, 2, self.row+14, 2, '',self.cellformat(right=4))
        self.sheet.merge_range(self.row+13, 3, self.row+14, 3, '',self.cellformat(right=4))
        self.sheet.merge_range(self.row+13, 4, self.row+14, 4, '',self.cellformat(right=4))
        self.sheet.merge_range(self.row+13, 5, self.row+14, 5, '',self.cellformat(right=4))
        self.sheet.merge_range(self.row+13, 6, self.row+14, 7, '',self.cellformat('center',right=1))
        self.sheet.merge_range(self.row+13, 8, self.row+14, 8, 'V',self.cellformat('center', True, 10, right=4))
        self.sheet.merge_range(self.row+13, 9, self.row+14, 9, 'A',self.cellformat('center', True, 10, right=4))
        self.sheet.merge_range(self.row+13, 10, self.row+14, 10, 'kW',self.cellformat('center', True, 10, right=1))
        self.sheet.write(self.row+13, 11, 'min',self.cellformat('center',True,10,right=4,buttom=4))
        self.sheet.write(self.row+13, 12, 'max',self.cellformat('center',True,10,right=4,buttom=4))
        self.sheet.write(self.row+14, 11, 'm³/h',self.cellformat('center',True,10,right=4,))
        self.sheet.write(self.row+14, 12, 'm³/h',self.cellformat('center',True,10,right=4,))
        self.sheet.merge_range(self.row+13, 13, self.row+14, 13, 'm³/h',self.cellformat('center',True,10,right=4))
        self.sheet.merge_range(self.row+13, 14, self.row+14, 14, 'm³/h',self.cellformat('center',True,10,right=1))
        self.sheet.merge_range(self.row+13, 15, self.row+14, 15, 'm³/h',self.cellformat('center',True,10,right=4))
        self.sheet.write(self.row+13, 16, 'bei 6m/s',self.cellformat('center',True,10,right=4,buttom=4))
        self.sheet.write(self.row+13, 17, 'bei 10V',self.cellformat('center',True,10,right=1,buttom=4))
        self.sheet.write(self.row+14, 16, 'm³/h',self.cellformat('center',True,10,right=4))
        self.sheet.write(self.row+14, 17, 'm³/h',self.cellformat('center',True,10,right=1))
        self.sheet.merge_range(self.row+13, 18, self.row+14, 19,'',self.cellformat(right=2))
        self.sheet.merge_range(self.row+13, 20, self.row+14, 20,'',self.cellformat(right=2))

    def exportueberschrift1(self,row,daten):
        if not self.sheet:
            return
        if daten != '1':
            return
        self.sheet.set_row(row,30)

        self.sheet.write(row, 0, 'VSR Id',self.cellformat('center',True,10,left=2,buttom=1,top=2,right=4))
        self.sheet.write(row, 1, 'Slave von',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 2, 'BKS-Schild',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 3, 'Einbauort',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 4, 'Bezeichnung',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 5, 'Fabrikat',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.merge_range(row, 6,row, 7, 'Typ[Nutzung]',self.cellformat('center',True,10,buttom=1,top=2,right=1))
        self.sheet.write(row, 8, 'Spann- \nung',self.cellformat('center',True,10,buttom=1,top=2,right=4,textwrap=True))
        self.sheet.write(row, 9, 'Nenn- \nstrom',self.cellformat('center',True,10,buttom=1,top=2,right=4,textwrap=True))
        self.sheet.write(row, 10, 'Leist- \nung',self.cellformat('center',True,10,buttom=1,top=2,right=1,textwrap=True))
        self.sheet.merge_range(row, 11, row, 12,'Tagbetrieb',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 13, 'Nacht- \nbetrieb',self.cellformat('center',True,10,buttom=1,top=2,right=4,textwrap=True))
        self.sheet.write(row, 14, 'abgesenkter \nBetrieb',self.cellformat('center',True,10,buttom=1,top=2,right=1,textwrap=True))
        self.sheet.write(row, 15, 'Regler \nmin.',self.cellformat('center',True,10,buttom=1,top=2,right=4,textwrap=True))
        self.sheet.write(row, 16, 'max. \ngewünscht',self.cellformat('center',True,10,buttom=1,top=2,right=4,textwrap=True))
        self.sheet.write(row, 17, 'Regler \nVnenn ',self.cellformat('center',True,10,buttom=1,top=2,right=1,textwrap=True))
        self.sheet.merge_range(row, 18, row, 19,'Bemerkung Antrieb',self.cellformat('center',True,10,top=2,buttom=1,right=2))
        self.sheet.write(row, 20, 'Anmerkung',self.cellformat('center',True,10,right=2,top=2,buttom=1))

        self.sheet.merge_range(row+1, 0, row+2, 0, '',self.cellformat(left=2,right=4,buttom=1))
        self.sheet.merge_range(row+1, 1, row+2, 1, '',self.cellformat(right=4,buttom=1))
        self.sheet.merge_range(row+1, 2, row+2, 2, '',self.cellformat(right=4,buttom=1))
        self.sheet.merge_range(row+1, 3, row+2, 3, '',self.cellformat(right=4,buttom=1))
        self.sheet.merge_range(row+1, 4, row+2, 4, '',self.cellformat(right=4,buttom=1))
        self.sheet.merge_range(row+1, 5, row+2, 5, '',self.cellformat(right=4,buttom=1))
        self.sheet.merge_range(row+1, 5, row+2, 7, '',self.cellformat('center',right=1,buttom=1))
        self.sheet.merge_range(row+1, 8, row+2, 8, 'V',self.cellformat('center', True, 10, right=4,buttom=1))
        self.sheet.merge_range(row+1, 9, row+2, 9, 'A',self.cellformat('center', True, 10, right=4,buttom=1))
        self.sheet.merge_range(row+1, 10, row+2, 10, 'kW',self.cellformat('center', True, 10, right=1,buttom=1))
        self.sheet.write(row+1, 11, 'min',self.cellformat('center',True,10,right=4,buttom=4))
        self.sheet.write(row+1, 12, 'max',self.cellformat('center',True,10,right=4,buttom=4))
        self.sheet.write(row+2, 11, 'm³/h',self.cellformat('center',True,10,right=4,buttom=1))
        self.sheet.write(row+2, 12, 'm³/h',self.cellformat('center',True,10,right=4,buttom=1))
        self.sheet.merge_range(row+1, 13, row+2, 13, 'm³/h',self.cellformat('center',True,10,right=4,buttom=1))
        self.sheet.merge_range(row+1, 14, row+2, 14, 'm³/h',self.cellformat('center',True,10,right=1,buttom=1))
        self.sheet.merge_range(row+1, 15, row+2, 15, 'm³/h',self.cellformat('center',True,10,right=4,buttom=1))
        self.sheet.write(row+1, 16, 'bei 6m/s',self.cellformat('center',True,10,right=4,buttom=4))
        self.sheet.write(row+1, 17, 'bei 10V',self.cellformat('center',True,10,right=1,buttom=4))
        self.sheet.write(row+2, 16, 'm³/h',self.cellformat('center',True,10,right=4,buttom=1))
        self.sheet.write(row+2, 17, 'm³/h',self.cellformat('center',True,10,right=1,buttom=1))
        self.sheet.merge_range(row+1, 18, row+2, 19,'',self.cellformat(right=2,buttom=1))
        self.sheet.merge_range(row+1, 20, row+2, 20,'',self.cellformat(right=2,buttom=1))
        
    def GetExportdaten(self):
        datenforexport = []
        datenforexport.append(['TRENNUNG','Raum Zuluft'])
        LAB = []
        H24 = []
        LAB_1 = []
        H24_1 = []
        RZU = []
        RAB = []
        RZU_Luftauslass = []
        RAB_Luftauslass = []
        H24_Luftauslass = []
        LAB_Luftauslass = []
        H24_1_Luftauslass = []
        LAB_1_Luftauslass = []
        
        if self.raum.list_vsr0.Count > 0:
            for vsr in self.raum.list_vsr0:
                if vsr.art == 'RZU':
                    datenforexport.append(['VSR',vsr])
                    for subvsr in vsr.List_VSR:
                        if self.raum.list_vsr1.Contains(subvsr):
                            datenforexport.append(['VSR',subvsr])
                            RZU.append(subvsr)
                            for auslass in subvsr.Auslass:RZU_Luftauslass.append(auslass)

                    for auslass in vsr.Auslass:
                        if self.raum.list_auslass.Contains(auslass) and auslass not in RZU_Luftauslass:
                            datenforexport.append(['AUSLASS',Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                'RZU',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])
                            RZU_Luftauslass.append(auslass)
    
        if self.raum.list_vsr1.Count > 0:
            for vsr in self.raum.list_vsr1:
                if vsr.art == 'RZU' and vsr not in RZU:
                    datenforexport.append(['VSR',vsr])
                    for auslass in vsr.Auslass:RZU_Luftauslass.append(auslass)

        for auslass in self.raum.list_auslass:
            if auslass not in RZU_Luftauslass and auslass.art == 'RZU':
                datenforexport.append(['AUSLASS',Luftauslassfuerexport('Durchlass Ohne VR','',self.raumnummer,\
                    'RZU',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])

                RZU_Luftauslass.append(auslass)
        
        datenforexport.append(['TRENNUNG','Raum Abluft'])

        if self.raum.list_vsr0.Count > 0:
            for vsr in self.raum.list_vsr0:
                if vsr.art == 'RAB':
                    datenforexport.append(['VSR',vsr])
                    for subvsr in vsr.List_VSR:
                        if self.raum.list_vsr1.Contains(subvsr):
                            if subvsr.art == 'LAB':
                                LAB_1.append(subvsr)
                                for auslass in subvsr.Auslass:
                                    LAB_Luftauslass.append(auslass)
                            elif subvsr.art == '24h':
                                H24_1.append(subvsr)
                                for auslass in subvsr.Auslass:
                                    H24_Luftauslass.append(auslass)
                            else:
                                datenforexport.append(['VSR',subvsr])
                                RAB.append(subvsr)
                                for auslass in subvsr.Auslass:RAB_Luftauslass.append(auslass)
                    for auslass in vsr.Auslass:
                        if self.raum.list_auslass.Contains(auslass):
                            if auslass.art == 'RAB' and auslass not in RAB_Luftauslass:
                                datenforexport.append(['AUSLASS',Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    'RAB',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])
                                RAB_Luftauslass.append(auslass)
                            
                            elif auslass.art == 'LAB' and auslass not in LAB_Luftauslass:
                                LAB_Luftauslass.append(auslass)
                                LAB_1_Luftauslass.append(Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    'LAB',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'',''))
                            
                            elif auslass.art == '24h' and auslass not in H24_Luftauslass:
                                H24_Luftauslass.append(auslass)
                                H24_1_Luftauslass.append(Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    '24h',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'',''))
                            
        if self.raum.list_vsr1.Count > 0:
            for vsr in self.raum.list_vsr1:
                if vsr.art == 'RAB' and vsr not in RAB:
                    datenforexport.append(['VSR',vsr])
                    for auslass in vsr.Auslass:
                        if auslass.art == 'RAB':RAB_Luftauslass.append(auslass)
                        elif auslass.art == 'LAB':
                            LAB_Luftauslass.append(auslass)
                            LAB_1_Luftauslass.append(Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    'LAB',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'',''))
                        elif auslass.art == '24h':
                            H24_Luftauslass.append(auslass)
                            H24_1_Luftauslass.append(Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    '24h',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'',''))
                
        
        for auslass in self.raum.list_auslass:
            if auslass not in RAB_Luftauslass and auslass.art == 'RAB':
                datenforexport.append(['AUSLASS',Luftauslassfuerexport('Durchlass Ohne VR','',self.raumnummer,\
                    'RAB',auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])
                RAB_Luftauslass.append(auslass)
        
        datenforexport.append(['TRENNUNG','Labor Abluft'])
        if len(LAB_1) > 0:
            for vsr in LAB_1:datenforexport.append(['VSR',vsr])
        
        if len(LAB_1_Luftauslass) > 0:
            for auslass in LAB_1_Luftauslass:datenforexport.append(['AUSLASS',auslass])

        if self.raum.list_vsr0.Count > 0:
            for vsr in self.raum.list_vsr0:
                if vsr.art == 'LAB':
                    datenforexport.append(['VSR',vsr])
                    for subvsr in vsr.List_VSR:
                        if self.raum.list_vsr1.Contains(subvsr):
                            datenforexport.append(['VSR',subvsr])
                            LAB.append(subvsr)
                            for auslass in subvsr.Auslass:LAB_Luftauslass.append(auslass)
                    for auslass in vsr.Auslass:
                        if self.raum.list_auslass.Contains(auslass):
                            if auslass.art == 'LAB' and auslass not in LAB_Luftauslass:
                                datenforexport.append(['AUSLASS',Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    'LAB',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])
                                LAB_Luftauslass.append(auslass)

        if self.raum.list_vsr1.Count > 0:
            for vsr in self.raum.list_vsr1:
                if vsr.art == 'LAB' and vsr not in LAB and vsr not in LAB_1:
                    datenforexport.append(['VSR',vsr])

                    for auslass in vsr.Auslass:LAB_Luftauslass.append(auslass)
        for auslass in self.raum.list_auslass:
            if auslass not in LAB_Luftauslass and auslass.art == 'LAB':
                datenforexport.append(['AUSLASS',Luftauslassfuerexport('Durchlass Ohne VR','',self.raumnummer,\
                    'LAB',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])
                LAB_Luftauslass.append(auslass)

        datenforexport.append(['TRENNUNG','24h-Abluft'])
        if len(H24_1) > 0:
            for vsr in H24_1:datenforexport.append(['VSR',vsr])

        if len(H24_1_Luftauslass) > 0:
            for auslass in H24_1_Luftauslass:datenforexport.append(['AUSLASS',auslass])
                
        if self.raum.list_vsr0.Count > 0:
            for vsr in self.raum.list_vsr0:
                if vsr.art == '24h':
                    datenforexport.append(['VSR',vsr])
                    
                    for subvsr in vsr.List_VSR:
                        if self.raum.list_vsr1.Contains(subvsr):
                            datenforexport.append(['VSR',subvsr])
                            H24.append(subvsr)
                            
                            for auslass in subvsr.Auslass: H24_Luftauslass.append(auslass)
                    for auslass in vsr.Auslass:
                        if self.raum.list_auslass.Contains(auslass):
                            if auslass.art == '24h' and auslass not in H24_Luftauslass:
                                datenforexport.append(['AUSLASS',Luftauslassfuerexport('-->Optional',vsr.vsrid,self.raumnummer,\
                                    '',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])

                                H24_Luftauslass.append(auslass)

        if self.raum.list_vsr1.Count > 0:
            for vsr in self.raum.list_vsr1:
                if vsr.art == '24h' and vsr not in H24 and vsr not in H24_1:
                    datenforexport.append(['VSR',vsr])
                    for auslass in vsr.Auslass:H24_Luftauslass.append(auslass)
        
        for auslass in self.raum.list_auslass:
            if auslass not in H24_Luftauslass and auslass.art == '24h':
                datenforexport.append(['AUSLASS',Luftauslassfuerexport('Durchlass Ohne VR','',self.raumnummer,\
                    '',auslass.nutzung,auslass.Luftmengenmin,auslass.Luftmengenmax,auslass.Luftmengennacht,auslass.Luftmengentnacht,'','')])
                H24_Luftauslass.append(auslass)
        
        
        datenforexport.append(['TRENNUNG','Überströmung in'])

        liste_ein = self.raum.list_ueber['Ein']
        liste_aus = self.raum.list_ueber['Aus']
        dict_ein = {}
        dict_aus = {}
        for el in liste_ein:
            if el.anderesraum not in dict_ein.keys():
                dict_ein[el.anderesraum] = 0
            dict_ein[el.anderesraum] += el.menge
        for el in liste_aus:
            if el.anderesraum not in dict_aus.keys():
                dict_aus[el.anderesraum] = 0
            dict_aus[el.anderesraum] += el.menge
        
        if self.raum.ueber_in_manuell.ist > 0:
            uebermanuell = self.raum.ueber_in_manuell.ist
            ueberinnacht = uebermanuell 
            ueberintnacht = uebermanuell

            if self.isnachtbetrieb == 'Nein':
                ueberinnacht = '0'

            if self.istiefenachtbetrieb == 'Nein':
                ueberintnacht = '0'
            datenforexport.append(['ÜEBERSTROM',['Manuelle Eingabe',uebermanuell,ueberinnacht,ueberintnacht]])

        for raumnummer in sorted(dict_ein.keys()):
            uebermanuell = dict_ein[raumnummer]
            ueberinnacht = uebermanuell 
            ueberintnacht = uebermanuell

            if self.isnachtbetrieb == 'Nein':
                ueberinnacht = '0'

            if self.istiefenachtbetrieb == 'Nein':
                ueberintnacht = '0'
            datenforexport.append(['ÜEBERSTROM',['aus dem Raum: ' + raumnummer,uebermanuell,ueberinnacht,ueberintnacht]])

        datenforexport.append(['TRENNUNG','Überströmung aus'])

        if self.raum.ueber_aus_manuell.ist > 0:
            uebermanuell = self.raum.ueber_aus_manuell.ist
            ueberinnacht = uebermanuell 
            ueberintnacht = uebermanuell

            if self.isnachtbetrieb == 'Nein':
                ueberinnacht = '0'

            if self.istiefenachtbetrieb == 'Nein':
                ueberintnacht = '0'
            datenforexport.append(['ÜEBERSTROM',['Manuelle Eingabe',uebermanuell,ueberinnacht,ueberintnacht]])

        for raumnummer in sorted(dict_aus.keys()):
            uebermanuell = dict_aus[raumnummer]
            ueberinnacht = uebermanuell 
            ueberintnacht = uebermanuell

            if self.isnachtbetrieb == 'Nein':
                ueberinnacht = '0'

            if self.istiefenachtbetrieb == 'Nein':
                ueberintnacht = '0'
            datenforexport.append(['ÜEBERSTROM',['in dem Raum: ' + raumnummer,uebermanuell,ueberinnacht,ueberintnacht]])
        return datenforexport      
    
    def GetFinalExportdaten(self):
        datenforexport_temp = self.GetExportdaten()
        anzahl = len(datenforexport_temp)+15
        if anzahl <= 45:
            self.exportgesamtdaten(datenforexport_temp)
            return
        else:
            Liste = []
            for n, daten in enumerate(datenforexport_temp):
                if daten[0] == 'TRENNUNG':
                    Liste.append(n)
            for n1,nummer in enumerate(Liste):
                if nummer >= 30:
                    break
            nummer = Liste[n1-1]
            datenforexport_temp.insert(nummer,['Überschrift',''])
            datenforexport_temp.insert(nummer,['Überschrift',''])
            datenforexport_temp.insert(nummer,['Überschrift','1'])
            self.exportgesamtdaten(datenforexport_temp)
            return
    
    def exportgesamtdaten(self,datengesamt):
        self.exportheader()
        self.exportueberschrift()
        for n, daten in enumerate(datengesamt):
            if daten[0] == 'TRENNUNG':
                self.exporttrennung(self.row+n+15,daten[1])
            elif daten[0] == 'VSR':
                self.exportvsr(self.row+n+15,daten[1])
            elif daten[0] == 'AUSLASS':
                self.exportluftauslass(self.row+n+15,daten[1])
            elif daten[0] == 'ÜEBERSTROM':
                self.exportueberstromeinzeln(self.row+n+15,daten[1])
            else:
                if daten[1] == '1':self.rowbreaks.append(self.row+n+15)
                self.exportueberschrift1(self.row+n+15,daten[1])
        self.rowbreaks.append(self.row+len(datengesamt)+16)
        self.create_last(self.row+len(datengesamt)+15)
        self.rowende = self.row+len(datengesamt)+16
    
    def exporttrennung(self,row,daten):
        bold = self.book.add_format()
        bold.set_bold()
        normal = self.book.add_format()
        if daten == 'Raum Zuluft':
            self.sheet.merge_range(row, 0, row, 19,daten,self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#A2E0F4'))
            self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#A2E0F4'))
        elif daten == 'Raum Abluft':
            self.sheet.merge_range(row, 0, row, 19,daten,self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#F1EF8D'))
            self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#F1EF8D'))
        elif daten == 'Labor Abluft':
            self.sheet.merge_range(row, 0, row, 19,daten,self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#F1EF8D'))
            self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#F1EF8D'))
        elif daten == '24h-Abluft':
            self.sheet.merge_range(row, 0, row, 19,daten,self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#FFD757'))
            self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#FFD757'))
        elif daten == 'Überströmung in':
            self.sheet.merge_range(row, 0, row, 19,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#BFBFBF'))
            self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#BFBFBF'))
            self.sheet.write_rich_string(row, 0,normal,'Überströmung ', bold, 'in', ' dem Raum',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#BFBFBF'))
  
        elif daten == 'Überströmung aus':
            self.sheet.merge_range(row, 0, row, 19,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#BFBFBF'))
            self.sheet.write(row, 20,'',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#BFBFBF'))
            self.sheet.write_rich_string(row, 0,normal,'Überströmung ', bold, 'aus', ' dem Raum',self.cellformat('left',True,10,top=1,buttom=1,left=2,right=2,background='#BFBFBF'))

    def exportvsr(self,row,vsr):
        if not self.sheet:
            return
        self.sheet.set_row(row,26)
        vsr.vsrauswaelten()
        self.sheet.write(row, 0, vsr.vsrid,self.cellformat('left',False,10,left=2,right=4))
        self.sheet.write(row, 1, vsr.slavevon,self.cellformat('left',False,10,right=4))
        self.sheet.write(row, 2, vsr.BKSschild,self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 3, vsr.raumnummer,self.cellformat('left',False,10,right=4))
        
        self.sheet.write(row, 4, vsr.vsrart,self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:self.sheet.write(row, 5, vsr.VSR_Hersteller.fabrikat,self.cellformat('center',False,10,right=4))
        else:self.sheet.write(row, 5, '',self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:
            if vsr.nutzung:
                self.sheet.merge_range(row, 6, row, 7, vsr.VSR_Hersteller.typ+'[{}]'.format(vsr.nutzung),self.cellformat('left',False,10,right=1,textwrap=True))
            else:
                self.sheet.merge_range(row, 6, row, 7, vsr.VSR_Hersteller.typ,self.cellformat('left',False,10,right=1,textwrap=True))
        else:
            if vsr.ispps:
                self.sheet.merge_range(row, 6, row, 7, '{} [PPs]'.format(vsr.size),self.cellformat('left',False,10,right=1,textwrap=True))
            else:
                self.sheet.merge_range(row, 6, row, 7, '{} [Blech]'.format(vsr.size),self.cellformat('left',False,10,right=1,textwrap=True))
   
            
        if vsr.VSR_Hersteller:self.sheet.write(row, 8, vsr.VSR_Hersteller.spannung,self.cellformat('center',False,10,right=4))
        else:self.sheet.write(row, 8, '-',self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:self.sheet.write(row, 9, vsr.VSR_Hersteller.nennstrom,self.cellformat('center',False,10,right=4))
        else:self.sheet.write(row, 9, '-',self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:self.sheet.write(row, 10, vsr.VSR_Hersteller.leistung,self.cellformat('center',False,10,right=1))
        else:self.sheet.write(row, 10, '-',self.cellformat('center',False,10,right=1))

        if vsr.Luftmengenmin > vsr.Luftmengenmax:
            try:self.sheet.write(row, 11, round(vsr.Luftmengenmax,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 11, vsr.Luftmengenmax,self.cellformat('center',False,10,right=4))
            try:self.sheet.write(row, 12, round(vsr.Luftmengenmin,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 12, vsr.Luftmengenmin,self.cellformat('center',False,10,right=4))

        else:
            try:self.sheet.write(row, 11, round(vsr.Luftmengenmin,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 11, vsr.Luftmengenmin,self.cellformat('center',False,10,right=4))
            try:self.sheet.write(row, 12, round(vsr.Luftmengenmax,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 12, vsr.Luftmengenmax,self.cellformat('center',False,10,right=4))
        
        try:self.sheet.write(row, 13, round(vsr.Luftmengennacht,2),self.cellformat('center',False,10,right=4))
        except:self.sheet.write(row, 13, vsr.Luftmengennacht,self.cellformat('center',False,10,right=4))
        try:self.sheet.write(row, 14, round(vsr.Luftmengentnacht,2),self.cellformat('center',False,10,right=1))
        except:self.sheet.write(row, 14, vsr.Luftmengentnacht,self.cellformat('center',False,10,right=1))
        if vsr.VSR_Hersteller:self.sheet.write(row, 15, vsr.VSR_Hersteller.vmin,self.cellformat('center',False,10,right=4))
        else:self.sheet.write(row, 15, '',self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:self.sheet.write(row, 16, vsr.VSR_Hersteller.vmax,self.cellformat('center',False,10,right=4))
        else:self.sheet.write(row, 16, '',self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:self.sheet.write(row, 17, vsr.VSR_Hersteller.vnenn,self.cellformat('center',False,10,right=1))
        else:self.sheet.write(row, 17, '',self.cellformat('center',False,10,right=1))
        if vsr.VSR_Hersteller:
            self.sheet.merge_range(row, 18, row, 19,vsr.VSR_Hersteller.bemerkung,self.cellformat('center',False,8,right=2,textwrap=True))
        else:self.sheet.merge_range(row, 18,row, 19, '',self.cellformat('center',False,8,right=2,textwrap=True))
        if vsr.VSR_Hersteller:
            if vsr.VSR_Hersteller.dimension != vsr.size:
                self.sheet.write(row, 20, vsr.anmerkung,self.cellformat('center',False,10,right=2,font_ground='#FF0000'))
            else:
                self.sheet.write(row, 20, '',self.cellformat('center',False,10,right=2))
        else:self.sheet.write(row, 20, vsr.anmerkung,self.cellformat('center',False,10,right=2,font_ground='#FF0000'))
    
    def exportluftauslass(self,row,luftauslass):
        if not self.sheet:
            return
        self.sheet.set_row(row,24)
        self.sheet.write(row, 0, luftauslass.art,self.cellformat('left',False,10,left=2,right=4))
        self.sheet.write(row, 1, luftauslass.slavevon,self.cellformat('left',False,10,right=4))
        self.sheet.write(row, 2, '-',self.cellformat('left',False,10,right=4))
        self.sheet.write(row, 3, luftauslass.raumnummer,self.cellformat('left',False,10,right=4))
        self.sheet.write(row, 4, luftauslass.auslassart,self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 5, '-',self.cellformat('center',False,10,right=4))
        self.sheet.merge_range(row, 6, row, 7, luftauslass.nutzung,self.cellformat('left',False,8,right=1,textwrap=True))
        self.sheet.write(row, 8, '-',self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 9, '-',self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 10, '-',self.cellformat('center',False,10,right=1))

        if luftauslass.Luftmengenmin > luftauslass.Luftmengenmax:
            try:self.sheet.write(row, 11, round(luftauslass.Luftmengenmax,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 11, luftauslass.Luftmengenmax,self.cellformat('center',False,10,right=4))
            try:self.sheet.write(row, 12, round(luftauslass.Luftmengenmin,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 12, luftauslass.Luftmengenmin,self.cellformat('center',False,10,right=4))

        else:
            try:self.sheet.write(row, 11, round(luftauslass.Luftmengenmin,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 11, luftauslass.Luftmengenmin,self.cellformat('center',False,10,right=4))
            try:self.sheet.write(row, 12, round(luftauslass.Luftmengenmax,2),self.cellformat('center',False,10,right=4))
            except:self.sheet.write(row, 12, luftauslass.Luftmengenmax,self.cellformat('center',False,10,right=4))

        try:self.sheet.write(row, 13, round(luftauslass.Luftmengennacht,2),self.cellformat('center',False,10,right=4))
        except:self.sheet.write(row, 13, luftauslass.Luftmengennacht,self.cellformat('center',False,10,right=4))
        try:self.sheet.write(row, 14, round(luftauslass.Luftmengentnacht,2),self.cellformat('center',False,10,right=1))
        except:self.sheet.write(row, 14, luftauslass.Luftmengentnacht,self.cellformat('center',False,10,right=1))
        self.sheet.write(row, 15, '-',self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 16, '-',self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 17, '-',self.cellformat('center',False,10,right=1))
        self.sheet.merge_range(row, 18,row, 19, '',self.cellformat('center',False,8,right=2,textwrap=True))
        self.sheet.write(row, 20, '',self.cellformat('center',False,10,right=2))

    def exportueberstromeinzeln(self,row,daten):
        Raumnummer,lufttag,luftnacht,lufttnacht = daten
        self.sheet.merge_range(row, 0,row, 10, Raumnummer,self.cellformat('right',False,10,left=2,right=1))  
        try:self.sheet.write(row, 11, round(lufttag),self.cellformat('center',False,10,left=1,right=4))
        except:self.sheet.write(row, 11, lufttag,self.cellformat('center',False,10,left=1,right=4))
        try:self.sheet.write(row, 12, round(lufttag),self.cellformat('center',False,10,right=4))
        except:self.sheet.write(row, 12, lufttag,self.cellformat('center',False,10,right=4))
        try:self.sheet.write(row, 13, round(luftnacht),self.cellformat('center',False,10,right=4))
        except:self.sheet.write(row, 13, luftnacht,self.cellformat('center',False,10,right=4))
        try:self.sheet.write(row, 14, round(lufttnacht),self.cellformat('center',False,10,right=1))
        except:self.sheet.write(row, 14, lufttnacht,self.cellformat('center',False,10,right=1))
        self.sheet.merge_range(row, 15,row, 17, '',self.cellformat('center',False,10,right=1))  
        self.sheet.merge_range(row, 18,row, 19, '',self.cellformat('center',False,10,right=2))
        self.sheet.write(row, 20, '',self.cellformat('center',False,10,right=2))

    def exportuberstrom(self,row):
        if not self.sheet:
            return
            
RED = Brushes.Red
BLACK = Brushes.Black
BLUE = Brushes.Blue
BLUEVIOLET = Brushes.BlueViolet
HIDDEN = Visibility.Hidden
VISIBLE = Visibility.Visible

DICT_MEP_NUMMER_NAME = {} # 100.103 : 100.103 - Büro
DICT_MEP_AUSLASS = {} ## MEPID: OB(Auslässe)
DICT_MEP_VSR = {} ## MEPID: [VSRID]
DICT_MEP_UEBERSTROM = {} ## MEPID: {'Ein': ..., 'Aus':...}

ELEMID_UEBER = [] ## ElememntId Überstrom
LISTE_SCHACHT = []

DICT_VSR_MEP_NUR_NUMMER = {} ## VSRID: [MEPNr,MEPNr]
DICT_VSR_MEP = {} ## VSRID: [MEPNr-Name,MEPNr-Name]
DICT_VSR_VSRLISTE = {} ## VSRID: [VSRID,VSRID]
DICT_VSR_AUSLASS = {} ## VSRID: OB(Auslässe)
DICT_VSR = {} ## VSRID: Class VSR
LISTE_VSR = []
LISTE_IRIS = []

views = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.View3D)).ToElements()
for el in views:
    if el.Name == 'Berechnungsmodell_L_KG4xx_IGF':
        uidoc.ActiveView = el
        break

active_view = uidoc.ActiveView
if active_view.Name != 'Berechnungsmodell_L_KG4xx_IGF':
    logger.error('die Ansicht "Berechnungsmodell_L_KG4xx_IGF "nicht gefunden!')
    script.exit()

class ABFRAGE(forms.WPFWindow):
    def __init__(self,maininfo = '',checked=True,height = None,minmax = True):
        forms.WPFWindow.__init__(self,'abfrage.xaml')
        self.maininfo.Text = maininfo
        self.gridlenge = GridLength(0.0)
        self.hoehe = height
        self.minmax = minmax
        if self.minmax == False:
            self.maingrid.RowDefinitions[1].Height = self.gridlenge
        if self.hoehe != None:
            self.Height = self.hoehe
        self.checked = checked
        self.result = False
        if self.checked == True:
            self.bestaetigen.IsChecked = True
            self.maingrid.RowDefinitions[2].Height = self.gridlenge
        else:
            self.bestaetigen.IsChecked = False
            self.ja.IsEnabled = False
    
    def movewindow(self, sender, args):
        self.DragMove()
        
    def checkedchanged(self,sender,e):
        if sender.IsChecked == True:
            self.ja.IsEnabled = True
        else:
            self.ja.IsEnabled = False
    
    def yes(self,sender,e):
        self.result = True
        self.Close()
    
    def no(self,sender,e):
        self.result = False
        self.Close()

    @staticmethod
    def show(maininfo = '',checked=True):
        abfrage = ABFRAGE(maininfo,checked)
        abfrage.ShowDialog()
        return abfrage.result

class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

## Familien Einstellungen
class Itemtemplate(TemplateItemBase):
    def __init__(self,name,checked = False):
        TemplateItemBase.__init__(self)
        self.name = name
        self._checked = checked
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

IS_AUSLASS = ObservableCollection[Itemtemplate]()
IS_AUSLASS_D = ObservableCollection[Itemtemplate]()
IS_AUSLASS_LAB = ObservableCollection[Itemtemplate]()
IS_AUSLASS_24H = ObservableCollection[Itemtemplate]()
IS_VSR = ObservableCollection[Itemtemplate]()
IS_DOSSEL = ObservableCollection[Itemtemplate]()
def get_IS():
    Families = DB.FilteredElementCollector(doc).OfClass(GetClrType(DB.Family)).ToElements()
    Liste_Auslass = []
    Liste_Auslass_Lab = []
    Liste_Zubehoer = []
    for el in Families:
        FamName = el.Name
        if el.FamilyCategoryId.IntegerValue == -2008013:
            if FamName not in Liste_Auslass:
                Liste_Auslass.append(FamName)
            for typid in el.GetFamilySymbolIds():
                typ = doc.GetElement(typid)
                typname = typ.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if typname not in Liste_Auslass_Lab:
                    Liste_Auslass_Lab.append(typname)
        elif el.FamilyCategoryId.IntegerValue == -2008016:
            if FamName not in Liste_Auslass:
                Liste_Zubehoer.append(FamName)

    #     FamName = el.Symbol.FamilyName
    #     famundtyp = FamName + ': ' + el.Name
    #     if FamName not in Liste_Auslass:
    #         Liste_Auslass.append(FamName)
    #     if famundtyp not in Liste_Auslass_Lab:
    #         Liste_Auslass_Lab.append(famundtyp)
        
    # for el in Ductaccessorys:
    #     if el.Symbol.FamilyName not in Liste_Zubehoer:
    #         Liste_Zubehoer.append(el.Symbol.FamilyName)

    for el in sorted(Liste_Auslass):
        IS_AUSLASS_D.Add(Itemtemplate(el))
    for el in sorted(Liste_Zubehoer):
        IS_VSR.Add(Itemtemplate(el))
        IS_DOSSEL.Add(Itemtemplate(el))
    for el in sorted(Liste_Auslass_Lab):
        IS_AUSLASS.Add(Itemtemplate(el,True))
        IS_AUSLASS_LAB.Add(Itemtemplate(el))
        IS_AUSLASS_24H.Add(Itemtemplate(el))

get_IS()

class Familien(forms.WPFWindow):
    def __init__(self,IS_Auslass0,IS_Auslass1,IS_Auslass2,IS_Auslass3,IS_VSR,IS_Drossel):
        forms.WPFWindow.__init__(self,'Familien.xaml')
        self.IS_Auslass0 = IS_Auslass0
        self.IS_Auslass1 = IS_Auslass1
        self.IS_Auslass2 = IS_Auslass2
        self.IS_Auslass3 = IS_Auslass3
        self.IS_VSR = IS_VSR
        self.IS_Drossel = IS_Drossel
        self.read_config()
        self.auslass.ItemsSource = IS_Auslass0
        self.auslass_d.ItemsSource = IS_Auslass1
        self.auslass_lab.ItemsSource = IS_Auslass2
        self.auslass_24h.ItemsSource = IS_Auslass3
        self.vsr.ItemsSource = IS_VSR
        self.klappe.ItemsSource = IS_Drossel
        self.temp0 = ObservableCollection[Itemtemplate]()
        self.temp1 = ObservableCollection[Itemtemplate]()
        self.temp2 = ObservableCollection[Itemtemplate]()
        self.temp3 = ObservableCollection[Itemtemplate]()
        self.temp4 = ObservableCollection[Itemtemplate]()
        self.temp5 = ObservableCollection[Itemtemplate]()
        
    def read_config(self):
        def readconfig_intern(liste,liste1,checked):
            for el in liste:
                if el.name in liste1:
                    el.checked = checked
        try:readconfig_intern(self.IS_Auslass0,config.auslass,False)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Auslass1,config.auslassd,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_VSR,config.vsr,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Drossel,config.drossel,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Auslass2,config.auslasslab,True)
        except Exception as e:print(e)
        try:readconfig_intern(self.IS_Auslass3,config.auslass24h,True)
        except Exception as e:print(e)
    
    def write_config(self):
        def write_config_intern(liste0,checked):
            return [el.name for el in liste0 if el.checked == checked]
        config.auslass = write_config_intern(self.IS_Auslass0,False)
        config.auslassd = write_config_intern(self.IS_Auslass1,True)
        config.vsr = write_config_intern(self.IS_VSR,True)
        config.drossel = write_config_intern(self.IS_Drossel,True)
        config.auslasslab = write_config_intern(self.IS_Auslass2,True)
        config.auslass24h = write_config_intern(self.IS_Auslass3,True)
        script.save_config()
    
    def checkedchanged(self,sender,liste):
        Checked = sender.IsChecked
        if liste.SelectedItem is not None:
            try:
                if sender.DataContext in liste.SelectedItems:
                    for item in liste.SelectedItems:
                        try:item.checked = Checked
                        except:pass
                else:pass
            except:pass

    def checkedchanged_auslass(self,sender,e):
        self.checkedchanged(sender,self.auslass)

    def checkedchanged_auslass_lab(self,sender,e):
        self.checkedchanged(sender,self.auslass_lab)
    
    def checkedchanged_auslass_24h(self,sender,e):
        self.checkedchanged(sender,self.auslass_24h)
    
    def checkedchanged_auslass_d(self,sender,e):
        self.checkedchanged(sender,self.auslass_d)
    
    def checkedchanged_vsr(self,sender,e):
        self.checkedchanged(sender,self.vsr)
    
    def checkedchanged_klappe(self,sender,e):
        self.checkedchanged(sender,self.klappe)
    
    def suche_text_changed(self,sender,liste0,liste1,liste2):
        liste1.Clear()
        text = sender.Text
        if not text:liste2.ItemsSource = liste0
        else:
            for el in liste0:
                if el.name.upper().find(text.upper()) != -1:
                    liste1.Add(el)
            liste2.ItemsSource = liste1
        liste2.Items.Refresh()
    
    def suche_textchanged_Lab(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass2,self.temp0,self.auslass_lab)

    def suche_textchanged_24h(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass3,self.temp1,self.auslass_24h)
    
    def suche_textchanged_auslass(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass0,self.temp2,self.auslass)
    
    def suche_textchanged_auslass_d(self,sender,e):
        self.suche_text_changed(sender,self.IS_Auslass1,self.temp3,self.auslass_d)
    
    def suche_textchanged_vsr(self,sender,e):
        self.suche_text_changed(sender,self.IS_VSR,self.temp4,self.vsr)
    
    def suche_textchanged_klappe(self,sender,e):
        self.suche_text_changed(sender,self.IS_Drossel,self.temp5,self.klappe)
    
    def ok(self,sender,e):
        self.write_config()
        self.Close()
    
    def schliessen(self,sender,e):
        self.Close()
        script.exit()

        
familien = Familien(IS_AUSLASS,IS_AUSLASS_D,IS_AUSLASS_LAB,IS_AUSLASS_24H,IS_VSR,IS_DOSSEL)
try:familien.ShowDialog()
except Exception as e:
    logger.error(e)
    familien.Close()
    script.exit()

class MEPFOREXPORTITEM(Itemtemplate):
    def __init__(self, name, berechnungnach,checked=False):
        Itemtemplate.__init__(self,name, checked)
        self.berechnng = berechnungnach

class RaumluftbilanzExport(forms.WPFWindow):
    def __init__(self,path,Liste_MEP):
        self.path = path
        forms.WPFWindow.__init__(self,'einstellung.xaml')
        self.Liste_MEP = Liste_MEP
        self.liste_temp = ObservableCollection[MEPFOREXPORTITEM]()
        if os.path.exists(self.path):
            self.ordner.Text = self.path
        else:
            self.path = ''
        self.result = False
        self.LV_MEP.ItemsSource = self.Liste_MEP
    def suchetextchanged(self,sender,e):
        text = sender.Text
        self.liste_temp.Clear()
        if not text:
            self.LV_MEP.ItemsSource = self.Liste_MEP
            return
        else:
            for el in self.Liste_MEP:
                if el.name.upper().find(text.upper()) != -1:
                    self.liste_temp.Add(el)
            self.LV_MEP.ItemsSource = self.liste_temp
            return    
    def checkedchanged(self,sender,e):
        Checked = sender.IsChecked
        if self.LV_MEP.SelectedItem is not None:
            try:
                if sender.DataContext in self.LV_MEP.SelectedItems:
                    for item in self.LV_MEP.SelectedItems:
                        try:item.checked = Checked
                        except:pass
                else:pass
            except:pass

    def alle(self,sender,e):
        for el in self.LV_MEP.Items:
            el.checked = True

    def keine(self,sender,e):
        for el in self.LV_MEP.Items:
            el.checked = False
    
    def nur(self,sender,e):
        for el in self.LV_MEP.Items:
            if el.berechnng != 'keine':
                el.checked = True
    
    def ok(self,sender,e):
        if not self.ordner.Text:
            UI.TaskDialog.Show('Info.','Kein Ordner ausgewählt')
            return
        elif os.path.exists(self.ordner.Text) is False:
            UI.TaskDialog.Show('Info.','Ordner nicht vorhanden')
            self.ordner.Text = ''
            return
        self.result = True
        self.Close()
        return
    def schliessen(self,sender,e):

        self.result = False
        self.Close()
        return
    
    def movewindow(self, sender, args):
        self.DragMove()
    
    def durchsuchen(self,sender,args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Ordner auswählen"
        dialog.ShowNewFolderButton = True
        if dialog.ShowDialog() == DialogResult.OK:
            folder = dialog.SelectedPath
            self.ordner.Text = folder
            self.path = folder
        
VSR_AUSLASS_LISTE = config.auslass
DRO_AUSLASS_LISTE = config.auslassd
LAB_AUSLASS_LISTE = config.auslasslab
H24_AUSLASS_LISTE = config.auslass24h
VSR_FAMILIE_LISTE = config.vsr
DRO_FAMILIE_LISTE = config.drossel

## DICT_MEP_NUMMER_NAME = {} # 100.103 : 100.103 - Büro
def get_MEP_NUMMER_NAMR():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
    for el in spaces:
        DICT_MEP_NUMMER_NAME[el.Number] = [el.Number + ' - ' + el.LookupParameter('Name').AsString(),el.Id.ToString()]
        schacht = el.LookupParameter('TGA_RLT_InstallationsSchacht').AsInteger()
        if schacht == 1:
            name = el.LookupParameter('TGA_RLT_InstallationsSchachtName').AsString()
            if name not in LISTE_SCHACHT:
                LISTE_SCHACHT.append(name)

get_MEP_NUMMER_NAMR() 
LISTE_SCHACHT.sort()
## NoMEPSpace
class NoMEPSpace(Exception):
    def __init__(self,elemid,typ,family,art):
        self.elemid = elemid
        self.typ = typ
        self.family = family
        self.art = art
        self.errorinfo = '{}: Einbauort konnte nicht ermittelt werden, FamilieName: {}, TypName: {}, ElementId: {}'.format(self.art,self.family,self.typ,self.elemid.ToString())
    
    def __str__(self):
        return self.errorinfo

## Grundklasse für alle Familienexamplar
class FamilieExemplar(TemplateItemBase):
    def __init__(self,elemid,art):
        TemplateItemBase.__init__(self)
        self.elemid = elemid
        self.logger = logger
        self.art_temp = art
        self._art = ''
        self._Luftmengenmin = 0
        self._Luftmengennacht = 0
        self._Luftmengenmax = 0
        self._Luftmengentnacht = 0
        self._size = ''
        self._Anmerkung = ''
        self.elem = doc.GetElement(self.elemid)
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self._familyandtyp = self.familyname + ': ' + self.typname
        self.ismuster =  self.Muster_Pruefen()
        self.phase = doc.GetElement(self.elem.CreatedPhaseId)
        self.raum = ''
        self.raumnummer = ''
        self.raumid = ''
        self.GetRaum()
    
    @property
    def art(self):
        return self._art
    @art.setter
    def art(self,value):
        if value != self._art:
            self._art = value
            self.RaisePropertyChanged('art')
    
    @property
    def familyandtyp(self):
        return self._familyandtyp
    @familyandtyp.setter
    def familyandtyp(self,value):
        if value != self._familyandtyp:
            self._familyandtyp = value
            self.RaisePropertyChanged('familyandtyp')
    
    @property
    def size(self):
        return self._size
    @size.setter
    def size(self,value):
        if value != self._size:
            self._size = value
            self.RaisePropertyChanged('size')
    
    @property
    def Luftmengenmin(self):
        return self._Luftmengenmin
    @Luftmengenmin.setter
    def Luftmengenmin(self,value):
        if value != self._Luftmengenmin:
            self._Luftmengenmin = value
            self.RaisePropertyChanged('Luftmengenmin')
    
    @property
    def Luftmengenmax(self):
        return self._Luftmengenmax
    @Luftmengenmax.setter
    def Luftmengenmax(self,value):
        if value != self._Luftmengenmax:
            self._Luftmengenmax = value
            self.RaisePropertyChanged('Luftmengenmax')
    
    @property
    def Luftmengennacht(self):
        return self._Luftmengennacht
    @Luftmengennacht.setter
    def Luftmengennacht(self,value):
        if value != self._Luftmengennacht:
            self._Luftmengennacht = value
            self.RaisePropertyChanged('Luftmengennacht')
    
    @property
    def Luftmengentnacht(self):
        return self._Luftmengentnacht
    @Luftmengentnacht.setter
    def Luftmengentnacht(self,value):
        if value != self._Luftmengentnacht:
            self._Luftmengentnacht = value
            self.RaisePropertyChanged('Luftmengentnacht')
    
    @property
    def Anmerkung(self):
        return self._Anmerkung
    @Anmerkung.setter
    def Anmerkung(self,value):
        if value != self._Anmerkung:
            self._Anmerkung = value
            self.RaisePropertyChanged('Anmerkung')
    
    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
    def GetRaum(self):
        try:
            self.raumid = self.elem.Space[self.phase].Id.ToString()
            self.raumnummer = self.elem.Space[self.phase].Number
            self.raum = self.elem.Space[self.phase].Number + ' - ' + self.elem.Space[self.phase].LookupParameter('Name').AsString()
        except:
            if not self.ismuster:
                try:
                    param_einbauort = get_value(self.elem.LookupParameter('IGF_X_Einbauort'))
                    if param_einbauort not in DICT_MEP_NUMMER_NAME.keys():
                        raise NoMEPSpace(self.elemid,self.typname,self.familyname,self.art_temp)
                    else:
                        self.raum = DICT_MEP_NUMMER_NAME[param_einbauort][0]
                        self.raumid = DICT_MEP_NUMMER_NAME[param_einbauort][1]
                except NoMEPSpace as e:
                    self.logger.error(str(e))
            return

    def set_up(self):
        self.familyname = self.elem.Symbol.FamilyName
        self.typname = self.elem.Name
        self.familyandtyp = self.familyname + ': ' + self.typname
    
    def get_Size(self):
        try:
            conns = self.elem.MEPModel.ConnectorManager.Connectors

            for conn in conns:
                try:
                    self.lufttyp = conn.DuctSystemType.ToString()
                except:pass
                try:
                    return 'DN ' + str(int(round(conn.Radius*609.6)))
                except:
                    Height = int(round(conn.Height*304.8))
                    Width = int(round(conn.Width*304.8))
                    return str(max([Width,Height])) + 'x' + str(min([Width,Height]))
        except:return
    
    def wert_schreiben(self):
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(self.Luftmengenmin).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(self.Luftmengennacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(self.Luftmengenmax).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(self.Luftmengentnacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raumnummer)
        except Exception as e:pass

## Klasse für Überstromsgruppe
class Baugruppe:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        if self.Pruefen():
            self.Volumen = get_value(self.elem.LookupParameter('IGF_RLT_Überströmung'))
            self.Eingang = ''
            self.Ausgang = ''
            try:self.Analyse()
            except Exception as e:logger.error(e)

    @property
    def auslaesse(self):
        auslass_liste = []
        for elemid in self.elem.GetMemberIds():
            elem = doc.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008013' and elem.Symbol.FamilyName.upper().find("ÜBER") != -1:
                auslass_liste.append(elem.Id)
        return auslass_liste
    
    def Pruefen(self):
        return len(self.auslaesse) == 2
   
    def Analyse(self):
        for elemid in self.auslaesse:
            auslass = UeberStromAuslass(elemid,self.Volumen)
            if auslass.typ == 'Aus':
                self.Ausgang = auslass
            elif auslass.typ == 'Ein':
                self.Eingang = auslass
        self.Ausgang.anderesraum = self.Eingang.raum
        self.Eingang.anderesraum = self.Ausgang.raum

## Klasse für Überstrom
class UeberStromAuslass(FamilieExemplar):
    def __init__(self,elemid,vol):
        FamilieExemplar.__init__(self,elemid,'Überstrom')
        self._menge = vol
        self._anderesraum = ""
        self.typ = self.get_typ()
    
    @property
    def menge(self):
        return self._menge
    @menge.setter
    def menge(self,value):
        if value != self._menge:
            self._menge = value
            self.RaisePropertyChanged('menge')
    @property
    def anderesraum(self):
        return self._anderesraum
    @anderesraum.setter
    def anderesraum(self,value):
        if value != self._anderesraum:
            self._anderesraum = value
            self.RaisePropertyChanged('anderesraum')
        
    @property
    def conns(self):
        return list(self.elem.MEPModel.ConnectorManager.Connectors)

    def get_typ(self):
        conn = self.conns[0]
        if conn.Direction.ToString() == 'Out':
            return 'Aus'
        elif conn.Direction.ToString() == 'In':
            return 'Ein'

# ## Klasse für Auslässe
class Luftauslass(FamilieExemplar):
    """
    iris: '', -1, str(elemid)
    vsr = elemid
    vsrliste: []
    vsr kann nicht in vsrliste sein.
    wenn vsr in vsrliste ist, vsr ist falsch platziert.
    vsr:[]
    oder vsr: [vsr1,vsr2]
    """
    def __init__(self,elemid):
        FamilieExemplar.__init__(self,elemid,'Luftauslass')
        self.Liste_LAB_Auslass = LAB_AUSLASS_LISTE
        self.Liste_24h_Auslass = H24_AUSLASS_LISTE
        self.enabledmin = True
        self.enabledmax = True
        self.enablednacht = True
        self.enabledtnacht = True
        self.System = ''
        self.AnlNr = ''
        self.vsr = ''
        self.RoutingListe = []
        self.Einbauteile = []
        self.VSR_Liste = []
        self.lufttyp = ''
        self.iris = ''
        self.VSR_Class = None
        self.Haupt_Class = None
        self.IRIS_Class = None
        self.nutzung = ''
        self.slavevon = ''
        try:
            self.nutzung = self.elem.Symbol.LookupParameter('Beschreibung').AsString() + ', ' + self.elem.Symbol.LookupParameter('Typenkommentare').AsString()
        except:
            pass
        self.Luftmengenermitteln()
        self.size = self.get_Size()
        try:self.System = self.elem.LookupParameter('Systemtyp').AsValueString()
        except:self.System = ''
        try:
            self.AnlNr = doc.GetElement(self.elem.LookupParameter('Systemtyp').AsElementId()).LookupParameter('IGF_X_AnlagenNr').AsValueString()
        except:self.AnlNr = ''
      
        if self.familyandtyp not in VSR_AUSLASS_LISTE:
            self.get_RountingListe(self.elem)   

        self.get_Art()
        if self.familyname in DRO_AUSLASS_LISTE:
            if self.iris in [-1,'']:
                self.Anmerkung = 'Der Auslass ist von Drosselklappe geregelt aber keine ist damit angeschlossen.'

        else:
            if self.iris not in [-1,'']:
                self.Anmerkung = 'Der Auslass ist nicht von Drosselklappe geregelt aber eine ist damit angeschlossen.'

        if self.vsr in self.VSR_Liste:
            self.Anmerkung = 'VSR falsch angeschlossen. Grund kann sein, dass das Eingang des VSRs ist mit ein Zuluftgitter verbunden.'
        
        if not self.vsr:
            self.Anmerkung = 'Kein VSR.'
        else:
            try:self.slavevon = doc.GetElement(DB.ElementId(int(self.vsr))).LookupParameter('SBI_Bauteilnummerierung').AsString()
            except:pass
    
    def set_up(self):
        FamilieExemplar.set_up(self)
        self.Luftmengenermitteln()
        self.size = self.get_Size()
        self.get_Art()

    def Luftmengenermitteln(self):
        try:self.Luftmengenmin = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin'))),1)
        except:pass
        try:self.Luftmengennacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht'))),1)
        except:pass
        try:self.Luftmengenmax = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax'))),1)
        except:pass
        try:self.Luftmengentnacht = round(float(get_value(self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht'))),1)
        except:pass

    def get_RountingListe(self,element):
        if self.RoutingListe.Count > 500:return
        elemid = element.Id.ToString()
        self.RoutingListe.append(elemid)
        cate = element.Category.Name
        if not cate in ['Luftkanal Systeme', 'Rohr Systeme', 'Luftkanaldämmung außen', 'Luftkanaldämmung innen', 'Rohre', 'Rohrformteile', 'Rohrzubehör','Rohrdämmung']:
            conns = None
            try:conns = element.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = element.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2 and self.iris == '':self.iris = -1
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.RoutingListe:
                            if owner.Category.Name == 'Luftkanalzubehör':
                                faminame = owner.Symbol.FamilyName
                                typname = owner.Name
                                if faminame in VSR_FAMILIE_LISTE:
                                    conns_temp = owner.MEPModel.ConnectorManager.Connectors
                                    for conn_temp in conns_temp:
                                        if self.lufttyp == 'ReturnAir':
                                            if conn_temp.Direction.ToString() == 'In' or conn_temp.Description == 'Haupt':#???
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                        if self.lufttyp == 'SupplyAir':
                                            if conn_temp.Direction.ToString() == 'Out' or conn_temp.Description == 'Haupt':#???
                                                if conn.IsConnectedTo(conn_temp):
                                                    self.vsr = owner.Id.ToString()
                                                    return
                                    if not self.vsr:
                                        self.vsr = owner.Id.ToString()
                                    self.VSR_Liste.append(owner.Id.ToString())
                                    return
                                elif faminame in DRO_FAMILIE_LISTE:
                                    if not self.iris:self.iris = owner.Id.ToString()
                                    
                            self.get_RountingListe(owner)
    
    def get_Art(self):
        systyp = self.System
        if systyp:
            if systyp.upper().find('TIERHALTUNG') != -1:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB' 
            else:
                if systyp.upper().find('ZULUFT') != -1:
                    self.art = 'RZU'
                elif systyp.upper().find('ABLUFT') != -1:
                    self.art = 'RAB'
                else:
                    systemklassifizierung = self.elem.LookupParameter('Systemklassifizierung').AsString()
                    if systemklassifizierung.upper().find('ZULUFT') != -1:
                        self.art = 'RZU'
                    elif systemklassifizierung.upper().find('ABLUFT') != -1:
                        self.art = 'RAB'
        else:
            self.art = 'XXX'
        try:
            if self.familyandtyp in self.Liste_LAB_Auslass:
                self.art = 'LAB'
            elif self.familyandtyp in self.Liste_24h_Auslass:
                self.art = '24h'
                self.enabledmax = False
                self.enablednacht = False
                self.enabledtnacht = False
        except:
            pass

## Klasse für Volumenstromregler
class VSR(FamilieExemplar):
    def __init__(self,elemid):
        FamilieExemplar.__init__(self,elemid,'VSR')
        self.checked = False
        self.Auslass = ObservableCollection[Luftauslass]()
        self.liste_Raum = []
        self.slavevon = '-'
        self.liste_Raum_nurNummer = []
        self.size = self.get_Size()
        self.IsHaupt = False
        self.IsIris = False
        self.List_Iris = ObservableCollection[VSR]()
        self.List_VSR = ObservableCollection[VSR]()
        self.List_Haupt = ObservableCollection[VSR]()
        self._Hersteller = 'Wildeboer'
        self.Datenbank_rund = Liste_Datenbank 
        self.Datenbank_eck = Liste_Datenbank1 
        self.VSR_Hersteller = None
        self._VSR_HerstellerTyp = ''
        self._anmerkung = 'kein passender Typ gefunden'
        self.material = self.elem.LookupParameter('IGF_X_Material_Text')
        self._dict = {
            'VR1-80':DB.ElementId(40884949),
            'VR1-100':DB.ElementId(30817223),
            'VR1-125':DB.ElementId(30817226),
            'VR1-160':DB.ElementId(30817239),
            'VR1-200':DB.ElementId(30817229),
            'VR1-250':DB.ElementId(30817232),
            'VR1-315':DB.ElementId(30817235),
            'VRE1-100':DB.ElementId(30778826),
            'VRE1-125':DB.ElementId(30781828),
            'VRE1-160':DB.ElementId(30784831),
            'VRE1-200':DB.ElementId(29973863),
            'VRE1-250':DB.ElementId(29970867),
            'VRE1-315':DB.ElementId(29960042),
            'VRL1-80':DB.ElementId(40884949),
            'VRL1-100':DB.ElementId(30817223),
            'VRL1-125':DB.ElementId(30817226),
            'VRL1-160':DB.ElementId(30817239),
            'VRL1-200':DB.ElementId(30817229),
            'VRL1-250':DB.ElementId(30817232),
        }
        try:
            self.vsrid = self.elem.LookupParameter('SBI_Bauteilnummerierung').AsString()
        except:
            self.vsrid = ''
        try:
            self.BKSschild = self.elem.LookupParameter('IGF_GA_BKS-Schild').AsString()
        except:
            self.BKSschild = ''
        
        self.ispps = False
        
        
        # try:
        #     self.spannung = get_value(self.elem.Symbol.LookupParameter('IGF_GA_Spannung'))
        # except:
        #     self.spannung = '24'

        self.vsrart = 'VVR'
        self.istppspruefen()
        self.istkvr()

        self.nutzung = ''
        
    
    @property
    def Hersteller(self):
        return self._Hersteller
    @Hersteller.setter
    def Hersteller(self,value):
        if value != self._Hersteller:
            self._Hersteller = value
            self.RaisePropertyChanged('Hersteller')
    @property
    def VSR_HerstellerTyp(self):
        return self._VSR_HerstellerTyp
    @VSR_HerstellerTyp.setter
    def VSR_HerstellerTyp(self,value):
        if value != self._VSR_HerstellerTyp:
            self._VSR_HerstellerTyp = value
            self.RaisePropertyChanged('VSR_HerstellerTyp')
    @property
    def anmerkung(self):
        return self._anmerkung
    @anmerkung.setter
    def anmerkung(self,value):
        if value != self._anmerkung:
            self._anmerkung = value
            self.RaisePropertyChanged('anmerkung')
    
    def istppspruefen(self):
        if self.material:
            try:
                # material = self.familyandtyp
                if self.material.AsString().upper().find('PPS') != -1:
                    self.ispps = True
                else:
                    self.ispps = False
            except:
                self.ispps = False
    
    def istkvr(self):
        if self.familyandtyp.upper().find('KONST') != -1 or self.familyandtyp.upper().find('KVR') != -1:
            self.vsrart = 'KVR'
        else:
            self.vsrart = 'VVR'
    
    def changetype(self):
        # if self.art != 'RZU':
        # if self.art == '24h':
        #     try:self.material.Set('PPs')
        #     except:pass
        # elif self.art == 'RZU':
        #     try:self.material.Set('Blech')
        #     except:pass
        # elif self.art == 'RAB':
        #     try:self.material.Set('Blech')
        #     except:pass
        # else:
        #     if self.ispps:
        #         try:self.material.Set('PPs')
        #         except:pass
        #     else:
        #         try:self.material.Set('Blech')
        #         except:pass
        if self.VSR_HerstellerTyp in self._dict.keys():
            try:self.elem.ChangeTypeId(self._dict[self.VSR_HerstellerTyp])
            except:pass
        try:
            self.set_up()
        except:
            pass
    def vsrueberpruefen(self):
        if self.VSR_Hersteller:
            self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
            if self.VSR_Hersteller.dimension != self.size:
                self.anmerkung = 'In Projekt {} modelliert'.format(self.size)
            else:
                self.anmerkung = ''
        else:self.anmerkung = 'kein passender Typ gefunden'
            
    def vsrauswaelten(self):
        liste_volumen = [self.Luftmengenmax,self.Luftmengenmin,self.Luftmengennacht,self.Luftmengentnacht]
        while( 0 in liste_volumen):
            liste_volumen.remove(0)
        if not liste_volumen:
            self.VSR_Hersteller = None
            self.VSR_HerstellerTyp = ''
            return

        minvol = min(liste_volumen)
        maxvol = max(liste_volumen)

        if self.size.find('DN') != -1:
            if self.Hersteller in self.Datenbank_rund.keys():liste = self.Datenbank_rund[self.Hersteller]
            else:
                self.VSR_Hersteller = None
                self.VSR_HerstellerTyp = ''
                return
        else:
            if self.Hersteller in self.Datenbank_eck.keys():liste = self.Datenbank_eck[self.Hersteller]
            else:
                self.VSR_Hersteller = None
                self.VSR_HerstellerTyp = ''
                return

        if self.art == 'RZU':
            for art in [1,0]:
                if art in liste.keys():
                    listetemp = liste[art]
                    if self.vsrart in listetemp.keys():
                        listetemp_vsr = listetemp[self.vsrart]
                        if self.size.find('DN') != -1:
                            for dimension in sorted(listetemp_vsr.keys()):
                                for vsr_temp in listetemp_vsr[dimension]:
                                    if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                        self.VSR_Hersteller = vsr_temp
                                        self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
                                        return
                        else:
                            for max_a in sorted(listetemp_vsr.keys()):
                                for min_a in sorted(listetemp_vsr[max_a].keys()):
                                    for vsr_temp in listetemp_vsr[max_a][min_a]:
                                        if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                            self.VSR_Hersteller = vsr_temp
                                            self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
                                            return
        elif self.art in ['RAB','LAB','24h']:
            if self.ispps:
                if 3 in liste.keys():
                    listetemp = liste[3]
                    if self.vsrart in listetemp.keys():
                        listetemp_vsr = listetemp[self.vsrart]
                        if self.size.find('DN') != -1:
                            for dimension in sorted(listetemp_vsr.keys()):
                                for vsr_temp in listetemp_vsr[dimension]:
                                    if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                        self.VSR_Hersteller = vsr_temp
                                        self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
                                        return
                        else:
                            for max_a in sorted(listetemp_vsr.keys()):
                                for min_a in sorted(listetemp_vsr[max_a].keys()):
                                    for vsr_temp in listetemp_vsr[max_a][min_a]:
                                        if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                            self.VSR_Hersteller = vsr_temp
                                            self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
                                            return

            else:
                for art in [2,0]:
                    if art in liste.keys():
                        listetemp = liste[art]
                        if self.vsrart in listetemp.keys():
                            listetemp_vsr = listetemp[self.vsrart]
                            if self.size.find('DN') != -1:
                                for dimension in sorted(listetemp_vsr.keys()):
                                    for vsr_temp in listetemp_vsr[dimension]:
                                        if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                            self.VSR_Hersteller = vsr_temp
                                            self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
                                            return
                            else:
                                for max_a in sorted(listetemp_vsr.keys()):
                                    for min_a in sorted(listetemp_vsr[max_a].keys()):
                                        for vsr_temp in listetemp_vsr[max_a][min_a]:
                                            if vsr_temp.vmin <= minvol and vsr_temp.vmax >= maxvol:
                                                self.VSR_Hersteller = vsr_temp
                                                self.VSR_HerstellerTyp = self.VSR_Hersteller.typ
                                                return
        self.VSR_Hersteller = None
        self.VSR_HerstellerTyp = ''
        return
        
    def set_up(self):
        FamilieExemplar.set_up(self)
        self.istkvr()
        self.istppspruefen()
        
        self.size = self.get_Size()
        self.get_Art()
        self.vsrauswaelten()
        self.vsrueberpruefen()

    def Luftmengenermitteln(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(auslass.Luftmengenmin)
            Luftmengennacht += float(auslass.Luftmengennacht)
            Luftmengenmax += float(auslass.Luftmengenmax)
            Luftmengentnacht += float(auslass.Luftmengentnacht)
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht))
    
    def Luftmengenermitteln_new(self):
        if self.art == 'RUM':
            self.Luftmengenermitteln_um()
            return
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0
        for auslass in self.Auslass:
            Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
            Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
            Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
            Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def Luftmengenermitteln_um(self):
        Luftmengenmin = 0
        Luftmengennacht = 0
        Luftmengenmax = 0
        Luftmengentnacht = 0

        for auslass in self.Auslass:
            if auslass.art == 'UMZU':
                Luftmengenmin += float(str(auslass.Luftmengenmin).replace(',', '.'))
                Luftmengennacht += float(str(auslass.Luftmengennacht).replace(',', '.'))
                Luftmengenmax += float(str(auslass.Luftmengenmax).replace(',', '.'))
                Luftmengentnacht += float(str(auslass.Luftmengentnacht).replace(',', '.'))
        self.Luftmengenmin = int(round(Luftmengenmin))
        self.Luftmengennacht = int(round(Luftmengennacht))   
        self.Luftmengenmax = int(round(Luftmengenmax))
        self.Luftmengentnacht = int(round(Luftmengentnacht)) 

    def get_Art(self):
        art_liste = []
        for auslass in self.Auslass:
            if auslass.art:
                if auslass.art not in art_liste:
                    art_liste.append(auslass.art)
        
        if art_liste.Count == 0:return
        if art_liste.Count == 1:
            self.art = art_liste[0]
        elif art_liste.Count == 2:
            if 'RZU' in art_liste and 'RAB' in art_liste:
                self.art = 'RUM'
                for auslass in self.Auslass:
                    if auslass.art == 'RZU':
                        auslass.art = 'UMZU'
                    elif auslass.art == 'RAB':
                        auslass.art = 'UMAB'
            elif 'RAB' in art_liste and 'LAB' in art_liste:
                self.art = 'RAB'
            else:
                print(self.elemid.ToString(),art_liste)
        elif art_liste.Count == 3 and '24h' in art_liste and 'LAB' in art_liste and 'RAB' in art_liste:self.art = 'RAB'
        else:print(self.elemid.ToString(),art_liste)

        if self.art == 'LAB' or self.art == '24h':
            for auslass in self.Auslass:
                self.nutzung = auslass.nutzung
                break

    
    def wert_schreiben(self):
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMin').SetValueString(str(self.Luftmengenmin).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromNacht').SetValueString(str(self.Luftmengennacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromMax').SetValueString(str(self.Luftmengenmax).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht').SetValueString(str(self.Luftmengentnacht).replace(',','.'))
        except Exception as e:pass
        try:self.elem.LookupParameter('IGF_X_Einbauort').Set(self.raumnummer)
        except Exception as e:pass
        try:
            numm = ''
            for nummer in sorted(self.liste_Raum_nurNummer):
                numm += nummer + ', '
            self.elem.LookupParameter('IGF_X_Wirkungsort').Set(numm[:-2])
        except:pass

def Get_Ueberstrom_Info():
    """
    DICT_MEP_UEBERSTROM:  {[raumid:{'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}}
    ELEMID_UEBER:   [lemid.ToString()]
    """
    Baugruppen = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title = "{value}/{max_value} Überströmungsbaugruppen",cancellable=True,step=10) as pb:
        for n, BGid in enumerate(Baugruppen):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(Baugruppen))
            baugruppe = Baugruppe(BGid)
            if not baugruppe.Pruefen():
                continue
            if not baugruppe.Eingang.raumid in DICT_MEP_UEBERSTROM.keys():
                DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
            if not baugruppe.Ausgang.raumid in DICT_MEP_UEBERSTROM.keys():
                DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid] = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
            DICT_MEP_UEBERSTROM[baugruppe.Eingang.raumid][baugruppe.Eingang.typ].Add(baugruppe.Eingang)
            DICT_MEP_UEBERSTROM[baugruppe.Ausgang.raumid][baugruppe.Ausgang.typ].Add(baugruppe.Ausgang)
            ELEMID_UEBER.append(baugruppe.Eingang.elemid.ToString())
            ELEMID_UEBER.append(baugruppe.Ausgang.elemid.ToString())

Get_Ueberstrom_Info()

def Get_Auslass_Info():
    """
    DICT_MEP_AUSLASS: {raumid:{auslass.art:{auslass.familyandtyp:[auslass(Klasse)]}}}
    DICT_VSR_MEP: {auslass.vsr: ["auslass.raumnr-auslass.raumname"],auslass.iris: ["auslass.raumnr-auslass.raumname"]}
    DICT_VSR_MEP_NUR_NUMMER: {auslass.vsr: [auslass.raumnummer],auslass.iris: [auslass.raumnummer]}
    DICT_VSR_VSRLISTE: {auslass.vsr: [auslass.VSR_Liste]} elemid: [elemid]
    DICT_MEP_VSR: {auslass.raumid: [auslass.iris,auslass.vsr]}
    DICT_VSR_AUSLASS: {auslass.vsr:{auslass.art:{auslass.familyandtyp:[auslass]}}}
    DICT_VSR: {vsrid:vsr}
    LISTE_VSR: [elemid]
    LISTE_IRIS: [elemid]

    """
    Ductterminalids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title = "{value}/{max_value} Luftauslässe", cancellable=True,step=10) as pb:
        for n, ductid in enumerate(Ductterminalids):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(Ductterminalids))
            if ductid.ToString() in ELEMID_UEBER:continue
            auslass = Luftauslass(ductid)
            
            if not auslass.raumid in DICT_MEP_AUSLASS.keys():
                 DICT_MEP_AUSLASS[auslass.raumid] = {}
            if not auslass.art in DICT_MEP_AUSLASS[auslass.raumid].keys():
                DICT_MEP_AUSLASS[auslass.raumid][auslass.art] = {}
            if not auslass.familyandtyp in DICT_MEP_AUSLASS[auslass.raumid][auslass.art].keys():
                DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp] = []
            DICT_MEP_AUSLASS[auslass.raumid][auslass.art][auslass.familyandtyp].append(auslass)

            if auslass.familyandtyp in VSR_AUSLASS_LISTE:
                continue
            if not auslass.vsr:
                if auslass.Muster_Pruefen() != True:
                    logger.error('Kein VSR mit Auslass {} angeschlossen. Raum {}, Familie: {}'.format(auslass.elemid,auslass.raumnummer,auslass.familyandtyp))
                continue
            
            if not auslass.vsr in DICT_VSR_MEP.keys():
                DICT_VSR_MEP[auslass.vsr] = [auslass.raum]
                DICT_VSR_MEP_NUR_NUMMER[auslass.vsr] = [auslass.raumnummer]
            else:
                if auslass.raum not in DICT_VSR_MEP[auslass.vsr]:
                    DICT_VSR_MEP[auslass.vsr].append(auslass.raum)
                    DICT_VSR_MEP_NUR_NUMMER[auslass.vsr].append(auslass.raumnummer)
            if auslass.iris not in [-1,'']:

                if not auslass.iris in DICT_VSR_MEP.keys():
                    DICT_VSR_MEP[auslass.iris] = [auslass.raum]
                    DICT_VSR_MEP_NUR_NUMMER[auslass.iris] = [auslass.raumnummer]
                else:
                    if auslass.raum not in DICT_VSR_MEP[auslass.iris]:
                        DICT_VSR_MEP[auslass.iris].append(auslass.raum)
                        DICT_VSR_MEP_NUR_NUMMER[auslass.iris].append(auslass.raumnummer)
            
            if len(auslass.VSR_Liste) != 0 and auslass.vsr not in auslass.VSR_Liste:
                if auslass.vsr not in DICT_VSR_VSRLISTE.keys():
                    DICT_VSR_VSRLISTE[auslass.vsr] = auslass.VSR_Liste
                else:
                    DICT_VSR_VSRLISTE[auslass.vsr].extend(auslass.VSR_Liste)
            else:
                LISTE_VSR.append(auslass.vsr)
            
            if auslass.iris not in [-1,'']:
                LISTE_IRIS.append(auslass.iris)

            if not auslass.raumid in DICT_MEP_VSR.keys():
                DICT_MEP_VSR[auslass.raumid] = []
                DICT_MEP_VSR[auslass.raumid].append(auslass.vsr) 
                if auslass.iris not in [-1,'']:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.iris)
            else:
                if auslass.vsr not in DICT_MEP_VSR[auslass.raumid]:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.vsr)
                if auslass.iris not in [-1,'']:
                    DICT_MEP_VSR[auslass.raumid].append(auslass.iris)

            if not auslass.vsr in DICT_VSR_AUSLASS.keys():
                DICT_VSR_AUSLASS[auslass.vsr] = {}
            if not auslass.art in DICT_VSR_AUSLASS[auslass.vsr].keys():
                DICT_VSR_AUSLASS[auslass.vsr][auslass.art] = {}
            if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.vsr][auslass.art].keys():
                DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp] = []
            DICT_VSR_AUSLASS[auslass.vsr][auslass.art][auslass.familyandtyp].append(auslass)
            if auslass.iris not in [-1,'']:
                if not auslass.iris in DICT_VSR_AUSLASS.keys():
                    DICT_VSR_AUSLASS[auslass.iris] = {}
                if not auslass.art in DICT_VSR_AUSLASS[auslass.iris].keys():
                    DICT_VSR_AUSLASS[auslass.iris][auslass.art] = {}
                if not auslass.familyandtyp in DICT_VSR_AUSLASS[auslass.iris][auslass.art].keys():
                    DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp] = []
                DICT_VSR_AUSLASS[auslass.iris][auslass.art][auslass.familyandtyp].append(auslass)  


    liste_temp = DICT_VSR_AUSLASS.keys()[:]
    liste_temp1 = LISTE_VSR[:]
    liste_temp2 = LISTE_IRIS[:]
    liste_temp.extend(liste_temp1)
    liste_temp.extend(liste_temp2)
    liste_temp = set(liste_temp)
    liste_temp = list(liste_temp)
   

    with forms.ProgressBar(title = "{value}/{max_value} Volumenstromregler", cancellable=True, step=10) as pb:
        for n, vsrid in enumerate(liste_temp):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n + 1, len(liste_temp))
            if vsrid in DICT_VSR_VSRLISTE.keys():
                vsrliste = set(DICT_VSR_VSRLISTE[vsrid])
                vsrliste = list(vsrliste)                
                for vsrid_neu in vsrliste:
                    if vsrid != vsrid_neu:
                        if vsrid_neu not in DICT_VSR_AUSLASS.keys():
                            logger.error('Kein Auslass mit VSR {} verbunden.'.format(vsrid_neu))
                            continue
                        for key0 in DICT_VSR_AUSLASS[vsrid_neu].keys():
                            for key1 in DICT_VSR_AUSLASS[vsrid_neu][key0].keys():
                                for auslass in DICT_VSR_AUSLASS[vsrid_neu][key0][key1]:
                                    if key0 not in DICT_VSR_AUSLASS[vsrid].keys():
                                        DICT_VSR_AUSLASS[vsrid][key0] = {}
                                    if key1 not in DICT_VSR_AUSLASS[vsrid][key0].keys():
                                        DICT_VSR_AUSLASS[vsrid][key0][key1] = []
                                    if auslass not in DICT_VSR_AUSLASS[vsrid][key0][key1]:
                                        DICT_VSR_AUSLASS[vsrid][key0][key1].append(auslass)

         
            vsr = VSR(DB.ElementId(int(vsrid)))
            if vsr.elemid.ToString() in LISTE_IRIS:
                vsr.IsIris = True
            elif vsrid in DICT_VSR_VSRLISTE.keys():
                vsr.IsHaupt = True

            vsr.Auslass = ObservableCollection[Luftauslass]()
            for art in sorted(DICT_VSR_AUSLASS[vsrid].keys()):
                for fam in sorted(DICT_VSR_AUSLASS[vsrid][art].keys()):
                    for terminal in DICT_VSR_AUSLASS[vsrid][art][fam]:
                        vsr.Auslass.Add(terminal)
            vsr.get_Art()
            vsr.Luftmengenermitteln()
            vsr.vsrauswaelten()
            vsr.vsrueberpruefen()
            vsr.liste_Raum = DICT_VSR_MEP[vsrid]
            vsr.liste_Raum_nurNummer = DICT_VSR_MEP_NUR_NUMMER[vsrid]
            
            DICT_VSR[vsrid] = vsr
            for auslass in vsr.Auslass:
                if vsr.IsIris:
                    auslass.IRIS_Class = vsr
                elif vsr.IsHaupt:
                    auslass.Haupt_Class = vsr
                else:
                    auslass.VSR_Class = vsr

Get_Auslass_Info()

class SchachtGrundinfo(object):
    def __init__(self,name):
        self.name = name

class MEPGrundInfo(TemplateItemBase):
    def __init__(self,name,soll,tooltip):
        TemplateItemBase.__init__(self)
        self.name = name
        self._soll = soll
        self._ist = ''
        self.tooltip = tooltip
    @property
    def soll(self):
        return self._soll
    @soll.setter
    def soll(self,value):
        if value != self._soll:
            self._soll = value
            self.RaisePropertyChanged('soll')
    @property
    def ist(self):
        return self._ist
    @ist.setter
    def ist(self,value):
        if value != self._ist:
            self._ist = value
            self.RaisePropertyChanged('ist')

class MEPSchachtInfo(TemplateItemBase):
    def __init__(self,name,nr,menge,SchachtListe = []):
        TemplateItemBase.__init__(self)
        self.name = name
        self.nr = nr
        self._menge = menge
        self.SchachtListe = SchachtListe
        self._schachtindex = -1
        self.get_index()
    
    @property
    def menge(self):
        return self._menge
    @menge.setter
    def menge(self,value):
        if value != self._menge:
            self._menge = value
            self.RaisePropertyChanged('menge')
    
    @property
    def schachtindex(self):
        return self._schachtindex
    @schachtindex.setter
    def schachtindex(self,value):
        if value != self._schachtindex:
            self._schachtindex = value
            self.RaisePropertyChanged('schachtindex')

    def get_index(self):
        if self.nr in self.SchachtListe:

            try:
                self.schachtindex = self.SchachtListe.index(self.nr)
            except:
                self.schachtindex = -1
        else:
            self.schachtindex = -1

class MEPAnlagenInfo(TemplateItemBase):
    def __init__(self,name,mep_nr,mep_mengen,sys_nr,sys_mengen):
        TemplateItemBase.__init__(self)
        self.name = name
        self._mep_nr = mep_nr
        self._mep_mengen = mep_mengen
        self._sys_nr = sys_nr
        self._sys_mengen = sys_mengen 
    @property
    def mep_nr(self):
        return self._mep_nr
    @mep_nr.setter
    def mep_nr(self,value):
        if value != self._mep_nr:
            self._mep_nr = value
            self.RaisePropertyChanged('mep_nr')
    @property
    def mep_mengen(self):
        return self._mep_mengen
    @mep_mengen.setter
    def mep_mengen(self,value):
        if value != self._mep_mengen:
            self._mep_mengen = value
            self.RaisePropertyChanged('mep_mengen')
    @property
    def sys_nr(self):
        return self._sys_nr
    @sys_nr.setter
    def sys_nr(self,value):
        if value != self._sys_nr:
            self._sys_nr = value
            self.RaisePropertyChanged('sys_nr')
    @property
    def sys_mengen(self):
        return self._sys_mengen
    @sys_mengen.setter
    def sys_mengen(self,value):
        if value != self._sys_mengen:
            self._sys_mengen = value
            self.RaisePropertyChanged('sys_mengen')

spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
Dict_Ueber_Manuell = {}
for ele in spaces_collector:
    raum = get_value(ele.LookupParameter("TGA_RLT_RaumÜberströmungAus"))
    if raum:
        summe2 = get_value(ele.LookupParameter('TGA_RLT_RaumÜberströmungMenge'))
        if raum not in Dict_Ueber_Manuell.keys():
            Dict_Ueber_Manuell[raum] = summe2
        else:Dict_Ueber_Manuell[raum] += summe2

class MEPRaum(object):
    def __init__(self, elemid,list_vsr):
        self.logger = logger
        self.berechnung_nach = {
                        '1': "Fläche",
                        '2': "Luftwechsel",
                        '3': "Person",
                        '4': "manuell",
                        '5': "nurZUMa",
                        '6': "nurABMa",
                        '5.1': "nurZU_Fläche",
                        '5.2': "nurZU_Luftwechsel",
                        '5.3': "nurZU_Person",
                        '6.1': "nurAB_Fläche",
                        '6.2': "nurAB_Luftwechsel",
                        '6.3': "nurAB_Person",
                        '9': "keine",
                    }
        self.einheit_liste = {
                    '1': 'm³/h pro m²',
                    '2': '-1/h',
                    '3': 'm3/h pro P',
                    '4': 'm³/h ',
                    '5': 'm³/h ',
                    '6': 'm³/h' ,
                    '5.1': "m³/h pro m²",
                    '5.2': '-1/h',
                    '5.3': 'm3/h pro P',
                    '6.1': "m³/h pro m²",
                    '6.2': '-1/h',
                    '6.3': 'm3/h pro P',
                    '9': '',
                }
        self.liste_schacht = LISTE_SCHACHT
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.box = self.elem.get_BoundingBox(None)
        self.Raumnr = self.elem.Number + ' - ' + self.elem.LookupParameter('Name').AsString()
        self.list_vsr = list_vsr
        self.list_vsr0 = ObservableCollection[VSR]() # Haupt
        self.list_vsr1 = ObservableCollection[VSR]() # VSR
        self.list_vsr2 = ObservableCollection[VSR]() # Iris
        for el in self.list_vsr:
            if el.IsIris == True:
                self.list_vsr2.Add(el)
            elif el.IsHaupt == True:
                self.list_vsr0.Add(el) 
            else:
                self.list_vsr1.Add(el)

        if self.list_vsr0.Count > 0:
            for el in self.list_vsr0:
                for auslass in el.Auslass:
                    iris = auslass.IRIS_Class
                    vsr = auslass.VSR_Class
                    if iris is not None:
                        if el.List_Iris.Contains(iris) == False:
                            el.List_Iris.Add(iris)
                            iris.List_Haupt.Add(el)
                    if vsr is not None:
                        if el.List_VSR.Contains(vsr) == False:
                            el.List_VSR.Add(vsr)
                            vsr.List_Haupt.Add(el)
                        if iris is not None:
                            if vsr.List_Iris.Contains(iris) == False:
                                vsr.List_Iris.Add(iris)
                                iris.List_VSR.Add(el)

        else:
            if self.list_vsr1.Count > 0:
                for el in self.list_vsr1:
                    for auslass in el.Auslass:
                        iris = auslass.IRIS_Class
                        if iris is not None:
                            if el.List_Iris.Contains(iris) == False:
                                el.List_Iris.Add(iris)
                                iris.List_VSR.Add(el)
        if self.list_vsr0.Count > 0:
            for mainvsr in self.list_vsr0:
                for subvsr in mainvsr.List_VSR:
                    if self.list_vsr1.Contains(subvsr):
                        if subvsr.slavevon != '-':
                            if subvsr.slavevon != mainvsr.vsrid:
                                logger.error('VSR {} hat zwei Haupt VSR {}, {}. Bitte überprüfen.'.format(subvsr.vsrid,mainvsr.vsrid,subvsr.slavevon))
                        subvsr.slavevon = mainvsr.vsrid
                        if subvsr.vsrid.find('-->') == -1:
                            subvsr.vsrid = '--> ' + subvsr.vsrid
                        

        # self.angezuluft = 0
        # self.angeabluft = 0

        self.IsTierRaum = (self.get_element('IGF_RLT_Tierkäfig_raumunabhängig') != 0)
        self.IsSchacht = (self.get_element('TGA_RLT_InstallationsSchacht') != 0)
        self.list_auslass = ObservableCollection[Luftauslass]()
        if self.elemid.ToString() in DICT_MEP_AUSLASS.keys():
             for art in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()].keys()):
                for fam in sorted(DICT_MEP_AUSLASS[self.elemid.ToString()][art].keys()):
                    for terminal in DICT_MEP_AUSLASS[self.elemid.ToString()][art][fam]:
                        self.list_auslass.Add(terminal)

        # try:self.list_auslass = DICT_MEP_AUSLASS[self.elemid.ToString()]
        # except:self.list_auslass = ObservableCollection[Luftauslass]()
        
        try:self.list_ueber = DICT_MEP_UEBERSTROM[self.elemid.ToString()]
        except:self.list_ueber = {'Ein':ObservableCollection[UeberStromAuslass](),'Aus':ObservableCollection[UeberStromAuslass]()}
        
        # Grundinfo.
        self.bezugsnummer = self.elem.LookupParameter('TGA_RLT_VolumenstromProNummer').AsValueString()
        try:self.bezugsname = self.berechnung_nach[self.bezugsnummer]
        except:self.bezugsname = 'keine'
        try:self.ebene = self.elem.LookupParameter('Ebene').AsValueString()
        except:self.ebene = ''
        self.flaeche = round(self.get_element('Fläche'), 2)
        # self.volumen = round(self.get_element('IGF_A_Volumen'),2)
        self.personen = round(self.get_element('Personenzahl'),1)
        self.faktor = self.get_element('TGA_RLT_VolumenstromProFaktor')
        try:self.einheit = self.einheit_liste[self.bezugsnummer]
        except:self.einheit = ''
        self.hoehe = int(self.get_element('IGF_A_Lichte_Höhe'))
        self.volumen = round((self.hoehe*self.flaeche/1000),2)
        # self.hoehe = round((self.volumen/self.flaeche),2)
        self.nachtbetrieb = self.get_element('IGF_RLT_Nachtbetrieb')
        self.tiefenachtbetrieb = self.get_element('IGF_RLT_TieferNachtbetrieb')
        self.NB_LW = self.get_element('IGF_RLT_NachtbetriebLW')
        self.T_NB_LW = self.get_element('IGF_RLT_TieferNachtbetriebLW')

        
        # Übersicht
        self.Uebersicht = ObservableCollection[MEPGrundInfo]()
        self.angezuluft = MEPGrundInfo('Ange.Zuluft',self.get_element('Angegebener Zuluftstrom'),'Angegebener Zuluftstrom')
        self.Uebersicht.Add(self.angezuluft)
        self.angeabluft = MEPGrundInfo('Ange.Abluft',self.get_element('Angegebener Abluftluftstrom'),'Angegebener Abluftluftstrom')
        self.Uebersicht.Add(self.angeabluft)

        self.ab_24h = MEPGrundInfo('24h Ab',self.get_element('IGF_RLT_AbluftSumme24h'),'IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_24h)

        self.ab_minsum = MEPGrundInfo('min.AB_SUM',0,'RAB,LAB,24h,Über,Über_M')
        self.Uebersicht.Add(self.ab_minsum)
        self.ab_min = MEPGrundInfo('min.RAB',self.get_element('IGF_RLT_AbluftminRaum'),'IGF_RLT_AbluftminRaum-IGF_RLT_AbluftminSummeLabor-IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_min)
        self.ab_lab_min = MEPGrundInfo('min.LAB',self.get_element('IGF_RLT_AbluftminSummeLabor'),'IGF_RLT_AbluftminSummeLabor')
        self.Uebersicht.Add(self.ab_lab_min)
        self.zu_minsum = MEPGrundInfo('min.ZU_SUM',0,'RZU,Über,Über_M')
        self.Uebersicht.Add(self.zu_minsum)
        self.zu_min = MEPGrundInfo('min.RZU',self.get_element('IGF_RLT_ZuluftminRaum'),'IGF_RLT_ZuluftminRaum')
        self.Uebersicht.Add(self.zu_min)
        
        self.ab_maxsum = MEPGrundInfo('max.AB_SUM',0,'RAB,LAB,24h,Über,Über_M')
        self.Uebersicht.Add(self.ab_maxsum)
        self.ab_max = MEPGrundInfo('max.RAB',self.get_element('IGF_RLT_AbluftmaxRaum'),'IGF_RLT_AbluftmaxRaum-IGF_RLT_AbluftmaxSummeLabor-IGF_RLT_AbluftSumme24h')
        self.Uebersicht.Add(self.ab_max)
        self.ab_lab_max = MEPGrundInfo('max.LAB',self.get_element('IGF_RLT_AbluftmaxSummeLabor'),'IGF_RLT_AbluftmaxSummeLabor')
        self.Uebersicht.Add(self.ab_lab_max)
        self.zu_maxsum = MEPGrundInfo('max.ZU_SUM',0,'RZU,Über,Über_M')
        self.Uebersicht.Add(self.zu_maxsum)
        self.zu_max = MEPGrundInfo('max.RZU',self.get_element('IGF_RLT_ZuluftmaxRaum'),'IGF_RLT_ZuluftmaxRaum')
        self.Uebersicht.Add(self.zu_max)

        self.nb_von = MEPGrundInfo('NB von',self.get_element('IGF_RLT_NachtbetriebVon'),'IGF_RLT_NachtbetriebVon')
        self.Uebersicht.Add(self.nb_von)
        self.nb_bis = MEPGrundInfo('NB bis',self.get_element('IGF_RLT_NachtbetriebBis'),'IGF_RLT_NachtbetriebBis')
        self.Uebersicht.Add(self.nb_bis)
        self.nb_dauer = MEPGrundInfo('NB Dauer',self.get_element('IGF_RLT_NachtbetriebDauer'),'IGF_RLT_NachtbetriebDauer')
        self.Uebersicht.Add(self.nb_dauer)
        self.nb_zu = MEPGrundInfo('NB Zu',self.get_element('IGF_RLT_ZuluftNachtRaum'),'IGF_RLT_ZuluftNachtRaum')
        self.Uebersicht.Add(self.nb_zu)
        self.nb_ab = MEPGrundInfo('NB Ab',self.get_element('IGF_RLT_AbluftNachtRaum'),'IGF_RLT_AbluftNachtRaum')
        self.Uebersicht.Add(self.nb_ab)

        self.tnb_von = MEPGrundInfo('TNB von',self.get_element('IGF_RLT_TieferNachtbetriebVon'),'IGF_RLT_TieferNachtbetriebVon')
        self.Uebersicht.Add(self.tnb_von)
        self.tnb_bis = MEPGrundInfo('TNB bis',self.get_element('IGF_RLT_TieferNachtbetriebBis'),'IGF_RLT_TieferNachtbetriebBis')
        self.Uebersicht.Add(self.tnb_bis)
        self.tnb_dauer = MEPGrundInfo('TNB Dauer',self.get_element('IGF_RLT_TieferNachtbetriebDauer'),'IGF_RLT_TieferNachtbetriebDauer')
        self.Uebersicht.Add(self.tnb_dauer)
        self.tnb_zu = MEPGrundInfo('TNB Zu',self.get_element('IGF_RLT_ZuluftTieferNachtRaum'),'IGF_RLT_ZuluftTieferNachtRaum')
        self.Uebersicht.Add(self.tnb_zu)
        self.tnb_ab = MEPGrundInfo('TNB Ab',self.get_element('IGF_RLT_AbluftTieferNachtRaum'),'IGF_RLT_AbluftTieferNachtRaum')
        self.Uebersicht.Add(self.tnb_ab)

        if self.IsTierRaum != 0:
            self.tier_zu_min = MEPGrundInfo('min.TZU',self.get_element('IGF_RLT_Luftmenge_min_TZU'),'IGF_RLT_Luftmenge_min_TZU')
            self.Uebersicht.Add(self.tier_zu_min)
            self.tier_ab_min = MEPGrundInfo('min.TAB',self.get_element('IGF_RLT_Luftmenge_min_TAB'),'IGF_RLT_Luftmenge_min_TAB')
            self.Uebersicht.Add(self.tier_ab_min)
            self.tier_zu_max = MEPGrundInfo('max.TZU',self.get_element('IGF_RLT_Luftmenge_max_TZU'),'IGF_RLT_Luftmenge_max_TZU')
            self.Uebersicht.Add(self.tier_zu_max)
            self.tier_ab_max = MEPGrundInfo('max.TAB',self.get_element('IGF_RLT_Luftmenge_max_TAB'),'IGF_RLT_Luftmenge_max_TAB')
            self.Uebersicht.Add(self.tier_ab_max)
        else:
            self.tier_zu_min = MEPGrundInfo('min.TZU',0,'IGF_RLT_Luftmenge_min_TZU')
            self.Uebersicht.Add(self.tier_zu_min)
            self.tier_ab_min = MEPGrundInfo('min.TAB',0,'IGF_RLT_Luftmenge_min_TAB')
            self.Uebersicht.Add(self.tier_ab_min)
            self.tier_zu_max = MEPGrundInfo('max.TZU',0,'IGF_RLT_Luftmenge_max_TZU')
            self.Uebersicht.Add(self.tier_zu_max)
            self.tier_ab_max = MEPGrundInfo('max.TAB',0,'IGF_RLT_Luftmenge_max_TAB')
            self.Uebersicht.Add(self.tier_ab_max)
        

        self.ueber_sum = MEPGrundInfo('Überstrom',self.get_element('IGF_RLT_ÜberströmungRaum'),'IGF_RLT_ÜberströmungRaum')
        self.Uebersicht.Add(self.ueber_sum)
        self.ueber_in = MEPGrundInfo('Über. Ein.',self.get_element('IGF_RLT_ÜberströmungSummeIn'),'IGF_RLT_ÜberströmungSummeIn')
        self.Uebersicht.Add(self.ueber_in)
        self.ueber_aus = MEPGrundInfo('Über. Aus.',self.get_element('IGF_RLT_ÜberströmungSummeAus'),'IGF_RLT_ÜberströmungSummeAus')
        self.Uebersicht.Add(self.ueber_aus)
        self.ueber_in_manuell = MEPGrundInfo('Über.Ein.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        self.Uebersicht.Add(self.ueber_in_manuell)
        self.ueber_aus_manuell = MEPGrundInfo('Über.Aus.M.',self.get_element('TGA_RLT_RaumÜberströmungMenge'),'TGA_RLT_RaumÜberströmungMenge')
        try:self.ueber_aus_manuell.soll = Dict_Ueber_Manuell[self.elem.Number]
        except:self.ueber_aus_manuell.soll = 0
        self.Uebersicht.Add(self.ueber_aus_manuell)
        self.ueber_aus_manuell.ist =self.ueber_aus_manuell.soll
        self.ueber_in_manuell.ist = self.ueber_in_manuell.soll

        self.ab_min.soll = self.ab_min.soll - self.ab_24h.soll - self.ab_lab_min.soll
        self.ab_max.soll = self.ab_max.soll - self.ab_24h.soll - self.ab_lab_max.soll
        self.labnacht = 0
        self.labtnacht = 0
        self.ab24nacht = 0
        self.ab24tnacht = 0
        self.sum_update()

        
        
        self.IGF_Legende = ''     
            
        self.Druckstufe = MEPGrundInfo('Druckstufe',self.get_element('IGF_RLT_RaumDruckstufeEingabe'),'IGF_RLT_RaumDruckstufeEingabe')
        self.Uebersicht.Add(self.Druckstufe)
        # Anlagen
        self.Anlagen_info = ObservableCollection[MEPAnlagenInfo]()

        # Schacht
        self.Schacht_info = ObservableCollection[MEPSchachtInfo]()

        self.Tagesbetrieb_Berechnen()
        self.Nachtbetrieb_Berechnen()
        self.Druckstufe_Berechnen()
    
        self.Analyse()
        self.get_Anlagen_info()
        self.get_Schacht_info()
    
    def update(self):        
        self.ab_24h.soll = self.get_element('IGF_RLT_AbluftSumme24h')
        self.ab_lab_min.soll = self.get_element('IGF_RLT_AbluftminSummeLabor') 
        self.Druckstufe.soll = self.get_element('IGF_RLT_RaumDruckstufeEingabe')
    
    def sum_update(self):
        self.ab_minsum.soll = str(int(round(float(self.ab_min.soll)+float(self.ab_lab_min.soll)+float(self.ab_24h.soll)+float(self.ueber_aus.soll)+float(self.ueber_aus_manuell.soll)))) + \
            ' (' + str(int(self.ab_min.soll)) + ', ' + str(int(self.ab_lab_min.soll)) + ', ' + str(int(self.ab_24h.soll)) + ', ' + str(int(self.ueber_aus.soll)) + ', ' + str(int(self.ueber_aus_manuell.soll))+ ')'
        
        self.ab_maxsum.soll = str(int(round(float(self.ab_max.soll)+float(self.ab_lab_max.soll)+float(self.ab_24h.soll)+float(self.ueber_aus.soll)+float(self.ueber_aus_manuell.soll)))) + \
            ' (' + str(int(self.ab_max.soll)) + ', ' + str(int(self.ab_lab_max.soll)) + ', ' + str(int(self.ab_24h.soll)) + ', ' + str(int(self.ueber_aus.soll)) + ', ' + str(int(self.ueber_aus_manuell.soll))+ ')'
        
        self.zu_minsum.soll = str(int(round(float(self.zu_min.soll)+float(self.ueber_in.soll)+float(self.ueber_in_manuell.soll)))) + \
            ' (' + str(int(self.zu_min.soll)) + ', ' + str(int(self.ueber_in.soll)) + ', ' + str(int(self.ueber_in_manuell.soll))+ ')'
        
        self.zu_maxsum.soll = str(int(round(float(self.zu_max.soll)+float(self.ueber_in.soll)+float(self.ueber_in_manuell.soll)))) + \
            ' (' + str(int(self.zu_max.soll)) + ', ' + str(int(self.ueber_in.soll)) + ', ' + str(int(self.ueber_in_manuell.soll)) + ')'
        
        # self.ab_minsum.ist = str(int(round(float(self.ab_min.ist)+float(self.ab_lab_min.ist)+float(self.ab_24h.ist)+float(self.ueber_aus.ist)+float(self.ueber_aus_manuell.ist)))) + \
        #     ' (' + str(int(self.ab_min.ist)) + ', ' + str(int(self.ab_lab_min.ist)) + ', ' + str(int(self.ab_24h.ist)) + ', ' + str(int(self.ueber_aus.ist)) + ', ' + str(int(self.ueber_aus_manuell.ist))+ ')'
        
        # self.ab_maxsum.ist = str(int(round(float(self.ab_max.ist)+float(self.ab_lab_max.ist)+float(self.ab_24h.ist)+float(self.ueber_aus.ist)+float(self.ueber_aus_manuell.ist)))) + \
        #     ' (' + str(int(self.ab_max.ist)) + ', ' + str(int(self.ab_lab_max.ist)) + ', ' + str(int(self.ab_24h.ist)) + ', ' + str(int(self.ueber_aus.ist)) + ', ' + str(int(self.ueber_aus_manuell.ist))+ ')'
        
        # self.zu_minsum.ist = str(int(round(float(self.zu_min.ist)+float(self.ueber_in.ist)+float(self.ueber_in_manuell.ist)))) + \
        #     ' (' + str(int(self.zu_min.ist)) + ', ' + str(int(self.ueber_in.ist)) + ', ' + str(int(self.ueber_in_manuell.ist))+ ')'
        
        # self.zu_maxsum.ist = str(int(round(float(self.zu_max.ist)+float(self.ueber_in.ist)+float(self.ueber_in_manuell.ist)))) + \
        #     ' (' + str(int(self.zu_max.ist)) + ', ' + str(int(self.ueber_in.ist)) + ', ' + str(int(self.ueber_in_manuell.ist)) + ')'

    def luft_round(self,luft):
        zahl = luft%5
        if zahl != 0:return(int(luft/5)+1) * 5
        else:return luft
    
    def Druckstufe_Berechnen(self):
        n = abs(int(self.Druckstufe.soll/10)) if abs(int(self.Druckstufe.soll/10)) < 6 else 5
        if self.Druckstufe.soll > 0:self.IGF_Legende = n*'+'
        else:self.IGF_Legende = n*'-'
          
    def Labmin_24h_Druckstufe_Pruefen(self,zuluftmin):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll
        
        
        # if self.Druckstufe.soll < 0:
        #     if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll + self.Druckstufe.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #         zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.Druckstufe.soll + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # else:
        # if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #     zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # if self.Druckstufe.soll > 0:
        #     zuluftmin = zuluftmin + self.Druckstufe.soll 

        return zuluftmin
    
    def Labmin_24h_Druckstufe_Nacht_Pruefen(self,zuluftmin):

        zuluftmin =  zuluftmin - self.ueber_in_manuell.soll - self.ueber_in.soll
        # if self.Druckstufe.soll < 0:
        #     if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll + self.Druckstufe.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #         zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.Druckstufe.soll + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # else:
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll > zuluftmin + self.ueber_in_manuell.soll + self.ueber_in.soll:
            zuluftmin = self.ueber_aus.soll + self.ueber_aus_manuell.soll  + self.ab_lab_min.soll + self.ab_24h.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        if zuluftmin < 0:
            zuluftmin = 0
        # if self.Druckstufe.soll > 0:
        #     zuluftmin = zuluftmin + self.Druckstufe.soll 

        return zuluftmin

    def Labmax_24h_Druckstufe_Pruefen(self,zuluftmax):
        if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax :
            zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll
        
        # if self.Druckstufe.soll < 0:
        #     if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll + self.Druckstufe.soll > zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #         zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.Druckstufe.soll  - self.ueber_in_manuell.soll - self.ueber_in.soll
        # else:
        # if self.ueber_aus.soll + self.ueber_aus_manuell.soll + self.ab_lab_max.soll + self.ab_24h.soll > zuluftmax + self.ueber_in_manuell.soll + self.ueber_in.soll:
        #     zuluftmax =  self.ab_lab_max.soll + self.ab_24h.soll + self.ueber_aus.soll + self.ueber_aus_manuell.soll - self.ueber_in_manuell.soll - self.ueber_in.soll
        # if self.Druckstufe.soll > 0:
        #     zuluftmax = zuluftmax + self.Druckstufe.soll 

        return zuluftmax

    def Nachtbetrieb_Berechnen(self):
        if self.bezugsname in ['Fläche',"Luftwechsel","Person","manuell"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = self.luft_round(self.T_NB_LW * self.volumen)
                
                    self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll + 24.00
                self.nb_zu.soll = self.luft_round(self.NB_LW * self.volumen)
                self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_zu.soll = 0   
        elif self.bezugsname in ['nurZU_Fläche',"nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = 0-(self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll )#- self.Druckstufe.soll)
                    # self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = 0

                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0
                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll + 24.00
                self.nb_zu.soll = 0-(self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll )#- self.Druckstufe.soll)
                # self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = 0

            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_zu.soll = 0   
        elif self.bezugsname in ['nurAB_Fläche',"nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
            if self.nachtbetrieb:
                if self.tiefenachtbetrieb:
                    self.tnb_dauer.soll = self.tnb_bis.soll - self.tnb_von.soll
                    if self.tnb_dauer.soll <= 0:
                        self.tnb_dauer.soll += 24.00
                    self.tnb_zu.soll = 0
                    # self.tnb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.tnb_zu.soll)
                    self.tnb_ab.soll = self.tnb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll# - self.Druckstufe.soll
                else:
                    self.tnb_dauer.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_zu.soll = 0
                    self.tnb_ab.soll = 0

                self.nb_dauer.soll = self.nb_bis.soll - self.nb_von.soll - self.tnb_dauer.soll + 24.00
                self.nb_zu.soll = 0
                # self.nb_zu.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.nb_zu.soll)
                self.nb_ab.soll = self.nb_zu.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            else:
                self.nb_dauer.soll = 0
                self.nb_zu.soll = 0
                self.nb_ab.soll = 0
                self.tnb_dauer.soll = 0
                self.tnb_ab.soll = 0
                self.tnb_zu.soll = 0   
        else:
            self.nb_dauer.soll = 0
            self.nb_zu.soll = 0
            self.nb_ab.soll = 0
            self.tnb_dauer.soll = 0
            self.tnb_ab.soll = 0
            self.tnb_zu.soll = 0   

    def Tagesbetrieb_Berechnen(self):
        if self.flaeche == 0:
            return
        if self.bezugsname in ['nurZU_Fläche',"nurZU_Luftwechsel","nurZU_Person","nurZUMa"]:
            if (self.ab_lab_min.soll + self.ab_24h.soll> 0) or (self.ab_lab_max.soll + self.ab_24h.soll> 0):
                self.logger.error("Berechnungsprinzip von Raum {} ist Falsch. Der Raum ist nur über Überströmung ausströmt aber hat Laborabluft min: {}, max: {} m³/h und 24h-Abluft: {} m³/h".format(self.Raumnr,self.ab_lab_min.soll,self.ab_lab_max.soll,self.ab_24h.soll))
                return
            if self.bezugsname == "nurZU_Fläche":
                self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor))
            elif self.bezugsname == "nurZU_Luftwechsel":
                self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor))
            
            elif self.bezugsname == "nurZU_Person":
                self.zu_min.soll = self.luft_round(self.personen * float(self.faktor))
            elif self.bezugsname == "nurZUMa":
                self.zu_min.soll = self.luft_round(float(self.faktor))
            
            self.ab_min.soll = self.ab_max.soll = 0
            self.zu_max.soll = self.zu_min.soll

            self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen(self.zu_min.soll)
            self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen(self.zu_max.soll)
            self.angeabluft.soll = self.angezuluft.soll = self.zu_min.soll
            
            abweichung = self.ueber_aus_manuell.soll + self.ueber_aus.soll - self.zu_min.soll - self.ueber_in.soll - self.ueber_in_manuell.soll #+ self.Druckstufe.soll
            
            if abweichung >= 0:
                self.zu_min.soll += abweichung
                self.zu_max.soll += abweichung


            else:
                self.hinweis = "Achtung: Bitte Überströmung-Aus um {} m³/h erhöhen".format(int(0 - abweichung))
                self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))

        elif self.bezugsname in ['nurAB_Fläche',"nurAB_Luftwechsel","nurAB_Person","nurABMa"]:
            if self.bezugsname == "nurAB_Fläche":
                self.angezuluft.soll = self.luft_round(self.flaeche * float(self.faktor))
            elif self.bezugsname == "nurAB_Luftwechsel":
                self.angezuluft.soll = self.luft_round(self.volumen * float(self.faktor))
            elif self.bezugsname == "nurAB_Person":
                self.angezuluft.soll = self.luft_round(self.personen * float(self.faktor))
            elif self.bezugsname == "nurABMa":
                self.angezuluft.soll = self.luft_round(float(self.faktor))
                        
            self.angeabluft.soll = self.angezuluft.soll
            self.zu_min.soll = self.zu_max.soll = 0
            
            abweichung_max = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll #- self.Druckstufe.soll
            abweichung_min = self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            abweichung = self.ueber_in.soll + self.ueber_in_manuell.soll - self.angezuluft.soll
            if abweichung_max >= 0:
                self.ab_min.soll = abweichung_min
                self.ab_max.soll = abweichung_max
                if abweichung < 0:
                    self.hinweis = "Achtung: Bitte Überströmung-Ein um {} m³/h erhöhen".format(int(0 - abweichung))
                    self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))
            else:
                self.ab_min.soll = 0
                self.ab_max.soll = 0

  
                if abweichung < 0:
                    self.hinweis = "Achtung: Bitte Überströmung-Ein um {} m³/h erhöhen".format(max(int(0 - abweichung),int(0 - abweichung_max)))
                    self.logger.error("Raum {}: {}".format(self.Raumnr,self.hinweis))
               
        elif self.bezugsname in ['Fläche',"Luftwechsel","Person","manuell"]:
            if self.bezugsname == "Fläche":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.flaeche * float(self.faktor))
            elif self.bezugsname == "Luftwechsel":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.volumen * float(self.faktor))
            elif self.bezugsname == "Person":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(self.personen * float(self.faktor))
            elif self.bezugsname == "manuell":
                self.zu_max.soll = self.zu_min.soll = self.luft_round(float(self.faktor))
        
            
            if self.bezugsname == 'manuell':
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.zu_min.soll)
                self.zu_max.soll = self.Labmin_24h_Druckstufe_Nacht_Pruefen(self.zu_max.soll)
            else:
                self.zu_min.soll = self.Labmin_24h_Druckstufe_Pruefen(self.zu_min.soll)
                self.zu_max.soll = self.Labmax_24h_Druckstufe_Pruefen(self.zu_max.soll)

            self.angeabluft.soll = self.angezuluft.soll = self.zu_max.soll
            

            self.ab_max.soll = self.zu_max.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_max.soll - self.ab_24h.soll #- self.Druckstufe.soll
            self.ab_min.soll = self.zu_min.soll + self.ueber_in.soll + self.ueber_in_manuell.soll - self.ueber_aus.soll - self.ueber_aus_manuell.soll - self.ab_lab_min.soll - self.ab_24h.soll #- self.Druckstufe.soll
            
            if self.ab_max.soll <= 0:
                self.zu_max.soll -= self.ab_max.soll
                self.ab_max.soll = 0
            if self.ab_min.soll <= 0:
                self.zu_max.soll -= self.ab_min.soll
                self.ab_min.soll = 0
                       
        elif self.bezugsname == 'keine':
            self.zu_min.soll = 0
            self.angezuluft.soll = 0
            self.angeabluft.soll = 0
            self.ab_min.soll = 0
            self.zu_max.soll = 0
            self.ab_max.soll = 0
            
        self.get_Schacht_info()
        self.get_Anlagen_info()
        self.sum_update()

    def Analyse(self):
        min_zu = 0
        min_ab = 0
        max_zu = 0
        max_ab = 0
        ab24h = 0
        ab24h_nacht = 0
        ab24h_tnacht = 0
        lab_min = 0
        lab_max = 0
        lab_nacht = 0
        lab_tnacht = 0
        nb_zu = 0
        nb_ab = 0
        tnb_zu = 0
        tnb_ab = 0
        uber_in = 0
        uber_aus = 0


        for uber in self.list_ueber["Ein"]:
            uber_in += uber.menge
        for uber in self.list_ueber["Aus"]:
            uber_aus += uber.menge

        for auslass in self.list_auslass:
            if auslass.art == '24h':
                ab24h += auslass.Luftmengenmin
                ab24h_nacht += auslass.Luftmengennacht
                ab24h_tnacht += auslass.Luftmengentnacht
            elif auslass.art == 'LAB':
                lab_min += auslass.Luftmengenmin
                lab_max += auslass.Luftmengenmax
                lab_nacht += auslass.Luftmengennacht
                lab_tnacht += auslass.Luftmengentnacht
            elif auslass.art == 'RZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentnacht
            elif auslass.art == 'RAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
            elif auslass.art == 'TZU':
                min_zu += auslass.Luftmengenmin
                max_zu += auslass.Luftmengenmax
                nb_zu += auslass.Luftmengennacht
                tnb_zu += auslass.Luftmengentnacht
            elif auslass.art == 'TAB':
                min_ab += auslass.Luftmengenmin
                max_ab += auslass.Luftmengenmax
                nb_ab += auslass.Luftmengennacht
                tnb_ab += auslass.Luftmengentnacht
        
        self.zu_min.ist = int(round(min_zu))
        self.ab_min.ist = int(round(min_ab))
        self.zu_max.ist = int(round(max_zu))
        self.ab_max.ist = int(round(max_ab))
        self.ab_24h.ist = int(round(ab24h))
        self.ab_lab_min.ist = int(round(lab_min))
        self.ab_lab_max.ist = int(round(lab_max))
        self.nb_zu.ist = int(round(nb_zu))
        self.nb_ab.ist = int(round(nb_ab))
        self.tnb_zu.ist = int(round(tnb_zu))
        self.tnb_ab.ist = int(round(tnb_ab))
        self.ueber_in.ist = int(round(uber_in))
        self.ueber_aus.ist = int(round(uber_aus))
        self.ueber_sum.ist = int(round(uber_in-uber_aus))

        self.labnacht = int(round(lab_nacht))
        self.labtnacht = int(round(lab_tnacht))
        self.ab24nacht = int(round(ab24h_nacht))
        self.ab24tnacht = int(round(ab24h_tnacht))
        
    def get_Anlagen_info(self):
        self.Anlagen_info.Clear()
        Dict = {}
        for el in self.list_auslass:
            if not el.art in Dict.keys():
                Dict[el.art] = {}
            if not el.AnlNr in Dict[el.art].keys():
                Dict[el.art][el.AnlNr] = float(el.Luftmengenmax)
            else:
                Dict[el.art][el.AnlNr] += float(el.Luftmengenmax)
        
        if 'RZU' in Dict.keys():
            for anl in sorted(Dict['RZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                    self.zu_max.soll,anl,Dict['RZU'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_RZU') or self.get_element('IGF_RLT_Luftmenge_RZU'):
            self.Anlagen_info.Add(MEPAnlagenInfo('RZU',self.get_element('IGF_RLT_AnlagenNr_RZU'),\
                self.zu_max.soll,'',''))
        if 'RAB' in Dict.keys():
            for anl in sorted(Dict['RAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                    self.ab_max.soll,anl,Dict['RAB'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_RAB') or self.get_element('IGF_RLT_Luftmenge_RAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('RAB',self.get_element('IGF_RLT_AnlagenNr_RAB'),\
                self.ab_max.soll,'',''))
        if 'TZU' in Dict.keys():
            for anl in sorted(Dict['TZU'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                    self.tier_zu_max.soll,anl,Dict['TZU'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_TZU') or self.get_element('IGF_RLT_Luftmenge_TZU'):
            self.Anlagen_info.Add(MEPAnlagenInfo('TZU',self.get_element('IGF_RLT_AnlagenNr_TZU'),\
                self.tier_zu_max.soll,'',''))
        
        if 'TAB' in Dict.keys():
            for anl in sorted(Dict['TAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                    self.tier_ab_max.soll,anl,Dict['TAB'][anl]))
        else:
            # if self.get_element('IGF_RLT_AnlagenNr_TAB') or self.get_element('IGF_RLT_Luftmenge_TAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('TAB',self.get_element('IGF_RLT_AnlagenNr_TAB'),\
                self.tier_ab_max.soll,'',''))
        
        if '24h' in Dict.keys():
            for anl in sorted(Dict['24h'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                    self.ab_24h.soll,anl,Dict['24h'][anl]))
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_24h') or self.get_element('IGF_RLT_Luftmenge_24h'):
            self.Anlagen_info.Add(MEPAnlagenInfo('24h',self.get_element('IGF_RLT_AnlagenNr_24h'),\
                self.ab_24h.soll,'',''))
           
        if 'LAB' in Dict.keys():
            for anl in sorted(Dict['LAB'].keys()):
                self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                    self.ab_lab_max.soll,anl,Dict['LAB'][anl]))  
        else:
            #if self.get_element('IGF_RLT_AnlagenNr_LAB') or self.get_element('IGF_RLT_Luftmenge_LAB'):
            self.Anlagen_info.Add(MEPAnlagenInfo('LAB',self.get_element('IGF_RLT_AnlagenNr_LAB'),\
                self.ab_lab_max.soll,'',''))
        
    def get_Schacht_info(self):
        self.Schacht_info.Clear()
        self.rzu_Schacht = MEPSchachtInfo('RZU',self.get_element('TGA_RLT_SchachtZuluft'),self.zu_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.rzu_Schacht)
        self.rab_Schacht = MEPSchachtInfo('RAB',self.get_element('TGA_RLT_SchachtAbluft'),self.ab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.rab_Schacht)
        
        self.tzu_Schacht = MEPSchachtInfo('TZU',self.get_element('IGF_RLT_Schacht_TZU'),self.tier_zu_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.tzu_Schacht)
        self.tab_Schacht = MEPSchachtInfo('TAB',self.get_element('IGF_RLT_Schacht_TAB'),self.tier_ab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.tab_Schacht)
        self._24h_Schacht = MEPSchachtInfo('24h',self.get_element('TGA_RLT_Schacht24hAbluft'),self.ab_24h.soll,self.liste_schacht)
        self.Schacht_info.Add(self._24h_Schacht)
        self.lab_Schacht = MEPSchachtInfo('LAB',self.get_element('IGF_RLT_Schacht_LAB'),self.ab_lab_max.soll,self.liste_schacht)
        self.Schacht_info.Add(self.lab_Schacht)
        
    def get_element(self, param_name):
        param = self.elem.LookupParameter(param_name)
        if not param:
            self.logger.info("Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return get_value(param)
    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            try:
                if not wert is None:
                #  logger.info(
                #      "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                    if self.elem.LookupParameter(param_name):
                        if self.elem.LookupParameter(param_name).IsReadOnly is True:
                            self.logger.error(self.elemid)
                            self.logger.error(param_name)
                            return
                        self.elem.LookupParameter(param_name).SetValueString(str(wert))
                    else:
                        print(param_name)
            except Exception as e:print(e,0)
        def wert_schreiben2(param_name, wert):
            try:
                if self.elem.LookupParameter(param_name):
                    if self.elem.LookupParameter(param_name).IsReadOnly is True:
                        self.logger.error(self.elemid)
                        self.logger.error(param_name)
                        return
                    self.elem.LookupParameter(param_name).Set(wert)
                else:
                    print(param_name)
            except Exception as e:print(e,1)  
        def wert_schreiben3(param_name, wert):
            '''für Schacht'''
            try:
                if self.elem.LookupParameter(param_name):
                    if wert.schachtindex != -1:
                        self.elem.LookupParameter(param_name).Set(wert.SchachtListe[wert.schachtindex])
                else:
                    print(param_name)
                  
            except Exception as e:pass#print(e)

        wert_schreiben("Angegebener Zuluftstrom", self.angezuluft.soll)
        wert_schreiben("Angegebener Abluftluftstrom", self.angeabluft.soll)
        wert_schreiben("IGF_RLT_AbluftminRaum", self.ab_min.soll+self.ab_lab_min.soll+self.ab_24h.soll+self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_AbluftmaxRaum", self.ab_max.soll+self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_ZuluftminRaum", self.zu_min.soll)
        wert_schreiben("IGF_RLT_ZuluftmaxRaum", self.zu_max.soll)
        wert_schreiben("IGF_A_Lichte_Höhe", self.hoehe)
        wert_schreiben("IGF_A_Volumen", self.volumen)

        wert_schreiben2("TGA_RLT_VolumenstromProName", self.bezugsname)
        wert_schreiben("TGA_RLT_VolumenstromProEinheit", self.einheit)
        wert_schreiben("TGA_RLT_VolumenstromProNummer", self.bezugsnummer)
        wert_schreiben("TGA_RLT_VolumenstromProFaktor", float(self.faktor))

        wert_schreiben2("IGF_RLT_Nachtbetrieb", self.nachtbetrieb)
        wert_schreiben("IGF_RLT_NachtbetriebLW", self.NB_LW)
        wert_schreiben("IGF_RLT_NachtbetriebDauer", self.nb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftNachtRaum", self.nb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftNachtRaum", self.nb_ab.soll +self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_NachtbetriebVon", self.nb_von.soll)
        wert_schreiben("IGF_RLT_NachtbetriebBis", self.nb_bis.soll)

        wert_schreiben2("IGF_RLT_TieferNachtbetrieb", self.tiefenachtbetrieb)
        wert_schreiben("IGF_RLT_TieferNachtbetriebLW", self.T_NB_LW)
        wert_schreiben("IGF_RLT_TieferNachtbetriebDauer", self.tnb_dauer.soll)
        wert_schreiben("IGF_RLT_ZuluftTieferNachtRaum", self.tnb_zu.soll)
        wert_schreiben("IGF_RLT_AbluftTieferNachtRaum", self.tnb_ab.soll +self.ab_lab_max.soll+self.ab_24h.soll+self.tier_ab_max.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebVon", self.tnb_von.soll)
        wert_schreiben("IGF_RLT_TieferNachtbetriebBis", self.tnb_bis.soll)

        wert_schreiben("IGF_RLT_AbluftSumme24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_AbluftminRaumL24h", self.ab_24h.soll)    
        wert_schreiben("IGF_RLT_AbluftminSummeLabor", self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftminSummeLabor24h", self.ab_24h.soll + self.ab_lab_min.soll)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor24h", self.ab_24h.soll + self.ab_lab_max.soll)
        wert_schreiben("IGF_RLT_AbluftmaxSummeLabor", self.ab_lab_max.soll)

        wert_schreiben("IGF_RLT_RaumDruckstufeEingabe", self.Druckstufe.soll)
        wert_schreiben2("IGF_RLT_RaumDruckstufeLegende", self.IGF_Legende)
        
        # wert_schreiben("IGF_RLT_AnlagenRaumAbluft", self.ab_min.soll+self.ab_lab_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaumZuluft", self.zu_min.soll)
        # wert_schreiben("IGF_RLT_AnlagenRaum24hAbluft", self.ab_24h.soll)
        
        wert_schreiben("IGF_RLT_Luftmenge_RAB", self.ab_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_RZU", self.zu_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_24h", self.ab_24h.soll)
        wert_schreiben("IGF_RLT_Luftmenge_LAB", self.ab_lab_max.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TAB", self.tier_ab_min.soll)
        wert_schreiben("IGF_RLT_Luftmenge_min_TZU", self.tier_zu_min.soll)
        
        wert_schreiben3('TGA_RLT_SchachtZuluft',self.rzu_Schacht)
        wert_schreiben3('TGA_RLT_SchachtAbluft',self.rab_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TZU',self.tzu_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_TAB',self.tab_Schacht)
        wert_schreiben3('TGA_RLT_Schacht24hAbluft',self._24h_Schacht)
        wert_schreiben3('IGF_RLT_Schacht_LAB',self.lab_Schacht)

        for el in self.Anlagen_info:
            if el.name == 'RZU':
                wert_schreiben2('IGF_RLT_AnlagenNr_RZU',int(el.mep_nr))
            elif el.name == 'RAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_RAB',int(el.mep_nr))
            elif el.name == 'TZU':
                wert_schreiben2('IGF_RLT_AnlagenNr_TZU',int(el.mep_nr))
            elif el.name == 'TAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_TAB',int(el.mep_nr))
            elif el.name == 'LAB':
                wert_schreiben2('IGF_RLT_AnlagenNr_LAB',int(el.mep_nr))
            elif el.name == '24h':
                wert_schreiben2('IGF_RLT_AnlagenNr_24h',int(el.mep_nr))
       
DICT_MEP_ITEMSSOIRCE = {}
LISTE_MEP = ObservableCollection[MEPFOREXPORTITEM]()

def Get_MEP_Info():
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElementIds()
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb:
        for n,space_id in enumerate(spaces):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n+1, len(spaces))
            
            if space_id.ToString() in DICT_MEP_VSR.keys():
                list_vsr = ObservableCollection[VSR]()
                dict_vsr = {}
                for e in DICT_MEP_VSR[space_id.ToString()]:
                    temp = DICT_VSR[e]
                    if temp.art not in dict_vsr.keys():
                        dict_vsr[temp.art] = {}
                    if temp.familyandtyp not in dict_vsr[temp.art].keys():
                        dict_vsr[temp.art][temp.familyandtyp] = []
                    dict_vsr[temp.art][temp.familyandtyp].append(temp)
                for art in sorted(dict_vsr.keys()):
                    for fam in sorted(dict_vsr[art].keys()):
                        for kla in dict_vsr[art][fam]:
                            try:list_vsr.Add(kla)
                            except:pass
            else:list_vsr = ObservableCollection[VSR]()

            mepraum = MEPRaum(space_id,list_vsr)
            # if not mepraum.IsSchacht:
            DICT_MEP_ITEMSSOIRCE[mepraum.Raumnr] = mepraum
    for Raumnr in sorted(DICT_MEP_ITEMSSOIRCE.keys()):
        LISTE_MEP.Add(MEPFOREXPORTITEM(Raumnr,DICT_MEP_ITEMSSOIRCE[Raumnr].bezugsname))
    

Get_MEP_Info()

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
        Max = box.Transform.Inverse.OfPoint(self.GUI.mepraum.box.Max + self.XYZAnpassen)
        Min = box.Transform.Inverse.OfPoint(self.GUI.mepraum.box.Min - self.XYZAnpassen)
        box.Min = Min
        box.Max = Max
        view.SetSectionBox(box)
        doc.Regenerate()
        t.Commit()
        t.Dispose()
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                return
        uidoc.Dispose()
        doc.Dispose()
        view.Dispose()

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
        max_x,max_y,max_z = self.GUI.mepraum.box.Max.X,self.GUI.mepraum.box.Max.Y,self.GUI.mepraum.box.Max.Z
        min_x,min_y,min_z = self.GUI.mepraum.box.Min.X,self.GUI.mepraum.box.Min.Y,self.GUI.mepraum.box.Min.Z
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
        t.Dispose()
        views = uidoc.GetOpenUIViews()
        for v in views:
            if v.ViewId == view.Id:
                try:v.ZoomToFit() 
                except:pass
                return
        uidoc.Dispose()
        doc.Dispose()
        view.Dispose()
    
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
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()

    def LuftmengeverteilenMEP(self,uiapp):
        self.name = 'Luftmengeverteilen ' + self.GUI.mepraum.Raumnr
        task = ABFRAGE('Luftmenge in akt. Raum gleichmäßig verteilen?',True,130)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        try:self.Luftmengeverteilen(self.GUI.mepraum,task)
        except Exception as e:print(e)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        try:task.Dispose()
        except:pass
    
    def LuftmengeverteilenProjekt(self,uiapp):
        self.name = 'Luftmengeverteilen Projekt' 
        task = ABFRAGE('Luftmenge für das Projekt gleichmäßig verteilen?',False,160)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.mepraum_liste.values():  
         
            self.Luftmengeverteilen(mep,task)

        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        try:task.Dispose()
        except:pass
    
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
            print('24h-Abluft oder Laborabluft in MEP-Raum {} stimmt nicht übereinandern.'.format(mep.Raumnr))
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
   
        if len(h24.keys()) > 0:
            for key in h24.keys():
                for auslass in h24[key]:
                    auslass.Luftmengennacht = auslass.Luftmengenmin
                    auslass.Luftmengenmax = auslass.Luftmengenmin
                    if mep.tiefenachtbetrieb:auslass.Luftmengentnacht = auslass.Luftmengenmin
                    else:auslass.Luftmengentnacht = 0
                    
        if len(zu.keys()) == 1:
            for key in zu.keys():
                for auslass in zu[key]:
                    if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll * 1.0 / len(zu[key]),1)
                    if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll * 1.0 / len(zu[key]),1)
                    if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll * 1.0 / len(zu[key]),1)
                    if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll * 1.0 / len(zu[key]),1)
        
        elif len(zu.keys()) > 1:
            sum_luft = 0
            for key in zu.keys():
                for auslass in zu[key]:
                    if auslass.Luftmengenmin > 0:sum_luft += auslass.Luftmengenmin
                    else:sum_luft += 0.01
            for key in zu.keys():
                for auslass in zu[key]:
                    if auslass.Luftmengenmin > 0:
                        if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                    else:
                        if task.min:auslass.Luftmengenmin = round(mep.zu_min.soll *1.0 * 0.01 / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_zu.soll *1.0 * 0.01 / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.zu_max.soll *1.0 * 0.01 / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_zu.soll *1.0 * 0.01 / sum_luft,1)

        if len(ab.keys()) == 1:
            for key1 in ab.keys():
                for auslass in ab[key1]:
                    if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 / len(ab[key1]),1)
                    if task.nacht:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 / len(ab[key1]),1)
                    if task.max:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 / len(ab[key1]),1)
                    if task.tnacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 / len(ab[key1]),1)
        elif len(ab.keys()) > 1:
            sum_luft = 0
            for key in ab.keys():
                for auslass in ab[key]:
                    if auslass.Luftmengenmin > 0:sum_luft += auslass.Luftmengenmin
                    else:sum_luft += 0.01
            for key in ab.keys():
                for auslass in ab[key]:
                    if auslass.Luftmengenmin > 0:
                        if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 * auslass.Luftmengenmin / sum_luft,1)
                    else:
                        if task.min:auslass.Luftmengenmin = round(mep.ab_min.soll *1.0 * 0.01 / sum_luft,1)
                        if task.nacht:auslass.Luftmengennacht = round(mep.nb_ab.soll *1.0 * 0.01 / sum_luft,1)
                        if task.max:auslass.Luftmengenmax = round(mep.ab_max.soll *1.0 * 0.01 / sum_luft,1)
                        if task.tnacht:auslass.Luftmengentnacht = round(mep.tnb_ab.soll *1.0 * 0.01 / sum_luft,1)
        
        for vsr in mep.list_vsr:
            if vsr.art in ['RZU','RAB','LAB','24h','RUM']:
                vsr.Luftmengenermitteln_new()
                vsr.vsrauswaelten()
                vsr.vsrueberpruefen()
        self.AederungUebernehmen(mep)
    
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
            return 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        self.AederungUebernehmen(self.GUI.mepraum)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        try:task.Dispose()
        except:pass
    
    def AederungUebernehmenProjekt(self,uiapp):
        self.name = 'Änderung übernehmen Projekt' 
        task = ABFRAGE('Alle Änderung übernehmen?',False,100,False)
        task.ShowDialog()
        if task.result == False:
            try:task.Dispose()
            except:pass
            return 
        
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        for mep in self.GUI.mepraum_liste.values():  
            self.AederungUebernehmen(mep)
        t.Commit()
        t.Dispose()

        self.GUI.Auswertung_System()
        self.GUI.Auswertung_MEP()
        doc.Dispose()
        try:task.Dispose()
        except:pass

    def ExportRaumluftbilanz(self,uiapp):
        temp = RaumluftbilanzExport(self.GUI.path,self.GUI.ListeMEP)
        temp.ShowDialog()
        if not temp.result:
            try:temp.Dispose()
            except:pass
            return
        else:
            self.GUI.path = temp.path
        try:temp.Dispose()
        except:pass
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
        worksheet.set_column(4,6,13)
        worksheet.set_column(7,10,7)
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

    def vsranpassen(self,uiapp):
        self.name = 'VSR anpassen' 
        doc = uiapp.ActiveUIDocument.Document
        t = DB.Transaction(doc,self.name)
        t.Start()
        # for mep in self.GUI.mepraum_liste.keys():  
        #     if mep.upper().find('130.') != -1 and mep.upper().find('TIER') == -1:
        #         mepraum = self.GUI.mepraum_liste[mep]
        #         for vsr in mepraum.list_vsr:
        #             try:vsr.changetype()
        #             except:pass
        for vsr in self.GUI.mepraum.list_vsr:
            try:vsr.changetype()
            except:pass
        t.Commit()
        t.Dispose()
        doc.Dispose()
