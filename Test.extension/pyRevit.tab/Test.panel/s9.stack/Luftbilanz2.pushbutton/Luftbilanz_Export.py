# coding: utf8
from Luftbilanz_Config import number,name

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
                self.dimension = str(int(self.B))+'x'+str(int(self.A))
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
        self.mepueber = self.raum.ueber_sum.soll
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
        
    def exportueberschrift(self):
        if not self.sheet:
            return
        self.sheet.set_row(self.row+12,30)

        self.sheet.write(self.row+12, 0, 'VSR Id',self.cellformat('center',True,10,left=2,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 1, 'Slave von',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 2, 'Einbauort',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 3, 'Bezeichnung',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.write(self.row+12, 4, 'Fabrikat',self.cellformat('center',True,10,buttom=1,top=6,right=4))
        self.sheet.merge_range(self.row+12, 5,self.row+12, 7, 'Typ[Nutzung]',self.cellformat('center',True,10,buttom=1,top=6,right=1))
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
        self.sheet.merge_range(self.row+13, 5, self.row+14, 7, '',self.cellformat('center',right=1))
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
        self.sheet.write(row, 2, 'Einbauort',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 3, 'Bezeichnung',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.write(row, 4, 'Fabrikat',self.cellformat('center',True,10,buttom=1,top=2,right=4))
        self.sheet.merge_range(row, 5,row, 7, 'Typ[Nutzung]',self.cellformat('center',True,10,buttom=1,top=2,right=1))
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
                if auslass.familyandtyp.upper().find('VORHALTUNG') == -1:
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
                if auslass.familyandtyp.upper().find('VORHALTUNG') == -1:
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
                if auslass.familyandtyp.upper().find('VORHALTUNG') == -1:
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
        self.sheet.write(row, 2, vsr.raumnummer,self.cellformat('left',False,10,right=4))
        
        self.sheet.write(row, 3, vsr.vsrart,self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:self.sheet.write(row, 4, vsr.VSR_Hersteller.fabrikat,self.cellformat('center',False,10,right=4))
        else:self.sheet.write(row, 4, '',self.cellformat('center',False,10,right=4))
        if vsr.VSR_Hersteller:
            if vsr.nutzung:
                self.sheet.merge_range(row, 5, row, 7, vsr.VSR_Hersteller.typ+'\n[{}]'.format(vsr.nutzung),self.cellformat('left',False,10,right=1,textwrap=True))
            else:
                self.sheet.merge_range(row, 5, row, 7, vsr.VSR_Hersteller.typ,self.cellformat('left',False,10,right=1,textwrap=True))
        else:
            if vsr.ispps:
                self.sheet.merge_range(row, 5, row, 7, '{} [PPs]'.format(vsr.size),self.cellformat('left',False,10,right=1,textwrap=True))
            else:
                self.sheet.merge_range(row, 5, row, 7, '{} [Blech]'.format(vsr.size),self.cellformat('left',False,10,right=1,textwrap=True))
   
            
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
            self.sheet.merge_range(row, 18, row, 19,vsr.VSR_Hersteller.bemerkung,self.cellformat('center',False,10,right=2,textwrap=True))
        else:self.sheet.merge_range(row, 18,row, 19, '',self.cellformat('center',False,10,right=2,textwrap=True))
        if vsr.VSR_Hersteller:
            if vsr.VSR_Hersteller.dimension != vsr.size:
                self.sheet.write(row, 20, vsr.anmerkung,self.cellformat('center',False,10,right=2,font_ground='#FF0000'))
            else:
                self.sheet.write(row, 20, vsr.anmerkung,self.cellformat('center',False,10,right=2))
        else:self.sheet.write(row, 20, vsr.anmerkung,self.cellformat('center',False,10,right=2,font_ground='#FF0000'))
    
    def exportluftauslass(self,row,luftauslass):
        if not self.sheet:
            return
        self.sheet.set_row(row,24)
        self.sheet.write(row, 0, luftauslass.art,self.cellformat('left',False,10,left=2,right=4))
        self.sheet.write(row, 1, luftauslass.slavevon,self.cellformat('left',False,10,right=4))
        self.sheet.write(row, 2, luftauslass.raumnummer,self.cellformat('left',False,10,right=4))
        self.sheet.write(row, 3, luftauslass.auslassart,self.cellformat('center',False,10,right=4))
        self.sheet.write(row, 4, '-',self.cellformat('center',False,10,right=4))
        self.sheet.merge_range(row, 5, row, 7, luftauslass.nutzung,self.cellformat('left',False,10,right=1,textwrap=True))
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
 