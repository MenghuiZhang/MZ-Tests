# coding: utf8
import xlsxwriter
from rpw import revit,DB,UI
from pyrevit import script,forms


__title__ = "8.30 AKS Nummer(für NA)"
__doc__ = """
AKS-Nummer ins Modell schreiben.
RaumNummer in Beiteile schreiben.
Schema:720081CB-DA99-40DC-9415-E53F280AA1F1 in Familietyp
Form: KG + '-'+ SystemABK+ '.' + AnlagenNr.+'-'+BauteilABK+'.'+Fortlaufende Nr.+'_'+Gebäudename+'.'+Ebene+'.'+Raumnr.

Parameter:
IGF_X_KG_Exemplar
IGF_X_SystemKürzel_Exemplar
IGF_X_AnlagenNr_Exemplar
IGF_X_Bauteilnummerierung
IGF_X_Einbauort


[2022.06.29]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

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

doc = revit.doc
uidoc = revit.uidoc

projektname = doc.ProjectInformation.Name
projektnummer = doc.ProjectInformation.Number
IGF_LOGO = None

class Raumdaten:
    def __init__(self,raum,row = 0):
        self.row = row
        self.book = None
        self.sheet = None
        self.raum = raum
        self.raumnummer = self.raum.elem.Number
        self.raumname = self.raum.elem.LookupParameter('Name').AsString()
        self.ebene = self.raum.ebene

        self.flaeche = self.raum.flaeche
        self.personen = self.raum.personne
        self.hoehe = self.raum.hoehe
        self.volumen = self.raum.volumen

        self.meplabmin = self.raum.ab_lab_min.soll     
        self.meplabmax = self.raum.ab_lab_max.soll
        self.mepab24h = self.raum.ab_24h.soll
        self.mepdruck = self.raum.Druckstufe.soll
        self.mepueber = self.raum.ueber_in_manuell.soll + self.raum.ueber_in.soll - self.raum.ueber_aus.soll - self.raum.ueber_aus_manuell.soll 

        self.bezugsname = self.raum.bezugsname
        self.faktor = self.raum.faktor
        self.einheit = self.raum.einheit
        self.mepminzu = self.raum.zu_min.soll
        self.mepmaxzu = self.raum.zu_max.soll
        self.mepminab = self.raum.ab_min.soll + self.raum.ab_24h.soll + self.raum.ab_lab_min.soll
        self.mepmaxab = self.raum.ab_max.soll + self.raum.ab_24h.soll + self.raum.ab_lab_max.soll

        self.isnachtbetrieb = 'Ja' if self.raum.nachtbetrieb  else 'Nein'
        self.LW_nacht = self.raum.NB_LW
        self.mepnb_zu = self.raum.nb_zu.soll
        self.mepnb_ab = self.raum.nb_ab.soll + self.raum.ab_24h.soll + self.raum.ab_lab_min.soll

        self.istiefenachtbetrieb = 'Ja' if self.raum.tiefenachtbetrieb  else 'Nein'
        self.LW_tnacht = self.raum.T_NB_LW
        self.meptnb_zu = self.raum.tnb_zu.soll
        self.meptnb_ab = self.raum.tnb_ab.soll + self.raum.ab_24h.soll + self.raum.ab_lab_min.soll

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
        self.istnachtlab = self.raum.ab_lab_min.ist
        self.isttnachtlab = self.raum.ab_lab_min.ist

        self.istmin24h = self.raum.ab_24h.ist
        self.istmax24h = self.raum.ab_24h.ist
        self.istnacht24h = self.raum.ab_24h.ist
        self.isttnacht24h = self.raum.ab_24h.ist

        self.istueber = self.raum.ueber_in_manuell.ist + self.raum.ueber_in.ist - self.raum.ueber_aus.ist - self.raum.ueber_aus_manuell.ist 
        self.istminueber = self.istueber
        self.istmaxueber = self.istueber
        self.istnachtueber = self.istueber
        self.isttnachtueber = self.istueber

        self.istminsumme = self.istminzu - self.istminab - self.istminlab - self.istmin24h + self.istueber - self.mepdruck
        self.istmaxsumme = self.istmaxzu - self.istmaxab - self.istmaxlab - self.istmax24h + self.istueber - self.mepdruck
        self.istnachtsumme = self.istnachtzu - self.istnachtab - self.istminlab - self.istmin24h + self.istueber - self.mepdruck
        self.isttnachtsumme = self.isttnachtzu - self.isttnachtab - self.istminlab - self.istmin24h + self.istueber - self.mepdruck

        self.istmin = 'OK' if abs(self.istminsumme) < 3 else 'Passt nicht'
        self.istmax = 'OK' if abs(self.istmaxsumme) < 3 else 'Passt nicht'
        self.istnacht = 'OK' if abs(self.istnachtsumme) < 3 else 'Passt nicht'
        self.isttnacht = 'OK' if abs(self.isttnachtsumme) < 3 else 'Passt nicht'

    def exportheader(self):
        if not self.sheet:
            return
        
        self.sheet.merge_range(self.row,1,self.row,4,projektname,self.book.add_format({'align': 'left', 'bold': True, 'font_size': 16}))
        self.sheet.merge_range(self.row,7,self.row,9,projektnummer,self.book.add_format({'align': 'center', 'bold': True, 'font_size': 16}))
        self.sheet.merge_range(self.row,10,self.row+1,14,'',self.book.add_format())
        self.sheet.merge_range(self.row+1,2,self.row+1,9,'Gewerk: Lüftung',self.book.add_format({'align': 'left', 'bold': True, 'font_size': 12}))
        self.sheet.merge_range(self.row+2,2,self.row+2,9,'Raumdaten aus den Raumbuch',self.book.add_format({'align': 'center', 'bold': True, 'font_size': 12}))
        self.sheet.write(self.row, 0, 'Projekt:',self.book.add_format({'align': 'left', 'bold': True, 'font_size': 16}))
        self.sheet.merge_range(self.row,5,self.row,6,'Projektnummer:',self.book.add_format({'align': 'center', 'bold': True, 'font_size': 16}))
        self.sheet.insert_image('K'+str(self.row), IGF_LOGO)
        self.sheet.write(self.row+1, 0, 'Raumdaten',self.book.add_format({'align': 'left', 'bold': True, 'font_size': 10}))
        self.sheet.write(self.row+2, 0, 'Nummer',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+3, 0, 'Name',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+4, 0, 'Ebene',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 0, 'Personen',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 0, 'Fläche',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 0, 'Höhe',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+8, 0, 'Volumen',self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.write(self.row+2, 1, self.raumnummer,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+3, 1, self.raumname,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+4, 1, self.ebene,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 1, self.personen,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 1, self.flaeche,self.book.add_format({'align': 'left', 'font_size': 10, 'num_format': '''" m²"'''}))
        self.sheet.write(self.row+7, 1, self.hoehe,self.book.add_format({'align': 'left', 'font_size': 10, 'num_format': '''" mm"'''}))
        self.sheet.write(self.row+8, 1, self.volumen,self.book.add_format({'align': 'left', 'font_size': 10, 'num_format': '''" m³"'''}))

        self.sheet.merge_range(self.row+3, 2, self.row+3, 3,'Grundinfos',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+4, 2, 'Druckstufe',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 2, 'Labmin',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 2, 'Labmax',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 2, '24h-Abluft',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+8, 2, 'Überstrom',self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.write(self.row+4, 3, self.mepdruck,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 3, self.meplabmin,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 3, self.meplabmax,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 3, self.mepab24h,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+8, 3, self.mepueber,self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.merge_range(self.row+3, 4, self.row+3, 5,'Tagbetrieb',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+4, 4, 'Berechnung nach',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 4, 'Faktor',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 4, 'min Zu',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 4, 'min Ab',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+8, 4, 'max Zu',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+9, 4, 'max Ab',self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.write(self.row+4, 5, self.bezugsname,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 5, self.faktor,self.book.add_format({'align': 'left', 'font_size': 10, 'num_format': '''" {}"'''.format(self.einheit)}))
        self.sheet.write(self.row+6, 5, self.mepminzu,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 5, self.mepminab,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+8, 5, self.mepmaxzu,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+9, 5, self.mepmaxab,self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.merge_range(self.row+3, 6, self.row+3, 7,'Nachtbetrieb',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+4, 6, 'Nachtbetrieb',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 6, 'Faktor',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 6, 'Zuluft',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 6, 'Abluft',self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.write(self.row+4, 7, self.isnachtbetrieb,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 7, self.LW_nacht,self.book.add_format({'align': 'left', 'font_size': 10, 'num_format': '''" {}"'''.format(self.einheit)}))
        self.sheet.write(self.row+6, 7, self.mepnb_zu,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 7, self.mepnb_ab,self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.merge_range(self.row+3, 8, self.row+3, 9,'TiefeNachtbetrieb',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+4, 8, 'TiefeNachtbetrieb',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 8, 'Faktor',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+6, 8, 'Zuluft',self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 8, 'Abluft',self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.write(self.row+4, 9, self.istiefenachtbetrieb,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+5, 9, self.LW_tnacht,self.book.add_format({'align': 'left', 'font_size': 10, 'num_format': '''" {}"'''.format(self.einheit)}))
        self.sheet.write(self.row+6, 9, self.meptnb_zu,self.book.add_format({'align': 'left', 'font_size': 10}))
        self.sheet.write(self.row+7, 9, self.meptnb_ab,self.book.add_format({'align': 'left', 'font_size': 10}))

        self.sheet.write(self.row+2, 10, 'IST-Werte',self.book.add_format({'align': 'center', 'font_size': 10}))
        self.sheet.write(self.row+2, 11, 'min',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+2, 12, 'max',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+2, 13, 'nacht',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+2, 14, 'tnacht',self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+3, 10, 'RZU',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+4, 10, 'RAB',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+5, 10, 'LAB',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+6, 10, '24H',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+7, 10, 'Über',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+8, 10, 'Bilanz',self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+9, 10, 'Überprüfung',self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+3, 11, self.istminzu,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+3, 12, self.istmaxzu,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+3, 13, self.istnachtzu,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+3, 14, self.isttnachtzu,self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+4, 11, self.istminab,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+4, 12, self.istmaxab,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+4, 13, self.istnachtab,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+4, 14, self.isttnachtab,self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+5, 11, self.istminlab,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+5, 12, self.istmaxlab,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+5, 13, self.istnachtlab,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+5, 14, self.isttnachtlab,self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+6, 11, self.istmin24h,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+6, 12, self.istmax24h,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+6, 13, self.istnacht24h,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+6, 14, self.isttnacht24h,self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+7, 11, self.istminueber,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+7, 12, self.istmaxueber,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+7, 13, self.istnachtueber,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+7, 14, self.isttnachtueber,self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+8, 11, self.istminsumme,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+8, 12, self.istmaxsumme,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+8, 13, self.istnachtsumme,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+8, 14, self.isttnachtsumme,self.book.add_format({'align': 'right', 'font_size': 10}))

        self.sheet.write(self.row+9, 11, self.istmin,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+9, 12, self.istmax,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+9, 13, self.istnacht,self.book.add_format({'align': 'right', 'font_size': 10}))
        self.sheet.write(self.row+9, 14, self.isttnacht,self.book.add_format({'align': 'right', 'font_size': 10}))




