# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
import clr
import System
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles, Visibility
from System.Windows.Media import Brushes, BrushConverter
from System.Windows.Forms import OpenFileDialog,DialogResult
from System.Windows.Controls import *
from pyrevit.forms import WPFWindow
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal
from IGF_forms import abfrage
from IGF_lib import get_value
import os

__title__ = "8.12 BIM-ID und Bearbeitsungsbereiche (Systeme-Tabelle)"
__doc__ = """BIM-ID und Bearbeitunsbereich aus excel in Modell schreiben
Systemtyp erstellen/anpassen
Daten exportieren

[2022.03.25]
Version: 3.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

logger = script.get_logger()
output = script.get_output()

name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number

bimid_config = script.get_config(name+number+'BIM-ID und Bearbeitsungsbereiche')

muster_bb = ['KG4xx_Musterbereich']

uebergreifend = []

#Farbe
BLACK = Brushes.Black
RED = Brushes.Red

def resize(text,anzahl):
    return (anzahl-len(text)) * '0' + text

def get_bb_dict():
    '''return dict von Bearbeitungsbereich'''
    worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
    Workset_dict = {}
    for el in worksets:
        Workset_dict[el.Name] = el.Id.ToString()
    return Workset_dict

def get_line_dict():
    _DICT_LINE = {}
    LineInternal = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement).ToElements()
    for el in LineInternal:
        _DICT_LINE[el.Name] = el.Id
    return _DICT_LINE

def get_dict_systemtyp():
    dict_systemtyp = {}   
    luftsystyp = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsElementType().ToElements()
    rohrsystyp = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()
         
    for el in luftsystyp:
        name = el.LookupParameter('Typname').AsString()
        f = el.LineColor 
        try:rgb = resize(str(f.Red),3)+','+resize(str(f.Green),3)+','+resize(str(f.Blue),3)
        except:rgb = 'xxx,xxx,xxx'
        klasse = el.LookupParameter('Systemklassifizierung').AsString()
        try:line = doc.GetElement(el.LinePatternId).Name
        except:line = 'keine Überschreibung'
        try:
            GK = el.LookupParameter('IGF_X_Gewerkkürzel').AsString()
            if GK is None:GK = 'R'
        except:GK = 'R'
        try:
            KG = el.LookupParameter('IGF_X_Kostengruppe').AsValueString()
            if KG is None:KG = '000'
            elif len(KG) != 3:KG = '000'
        except:KG = '000'
        try:
            KN01 = el.LookupParameter('IGF_X_Kennnummer_1').AsValueString()
            if KN01 is None:KN01 = '00'
            elif len(KN01) != 2:KN01 = resize(KN01,2)
        except:KN01 = '00'
        try:
            KN02 = el.LookupParameter('IGF_X_Kennnummer_2').AsValueString()
            if KN02 is None:KN02 = '00'
            elif len(KN02) != 2:KN02 = resize(KN02,2)
        except:KN02 = '00'
        try:BIMID = GK +"_"+KG+"_"+KN01+"-"+KN02
        except:BIMID = ''
        try:
            Sysname = el.LookupParameter('IGF_X_SystemName').AsString()
            if Sysname is None:Sysname = ''
        except:Sysname = ''
        try:
            SysKZ = el.LookupParameter('IGF_X_SystemKürzel').AsString()
            if SysKZ is None:SysKZ = ''
        except:SysKZ = ''
        try:
            AnGeAn = el.LookupParameter('IGF_X_AnlagenGeräteAnzahl').AsValueString()
            if AnGeAn is None:AnGeAn = '1'
        except:AnGeAn = '1'
        try:
            AnNr = el.LookupParameter('IGF_X_AnlagenNr').AsValueString()
            if AnGeAn is None:AnGeAn = ''
        except:AnGeAn = ''
        try:
            PzAT = el.LookupParameter('IGF_X_AnlagenProzentualAnteil').AsValueString()
            if PzAT is None:PzAT = ''
            else:
                try:PzAT = str(float(PzAT)/100)+'%'
                except:pass
        except:PzAT = ''
        try:
            PzAZ = el.LookupParameter('IGF_X_AnlagenProzentualAnzahl').AsValueString()
            if PzAZ is None:PzAZ = ''            
        except:PzAZ = ''
        try:
            PzAZ = el.LookupParameter('IGF_X_AnlagenProzentualAnzahl').AsValueString()
            if PzAZ is None:PzAZ = ''            
        except:PzAZ = ''
        try:
            PzAZ = el.LookupParameter('IGF_X_AnlagenProzentualAnzahl').AsValueString()
            if PzAZ is None:PzAZ = '' 
        except:PzAZ = ''
        try:
            TempS = get_value(el.LookupParameter('IGF_RLT_ZuluftTemperaturSo'))
            if TempS is None or str(int(TempS)) == '-273':TempS = '' 
             
                      
        except:TempS = ''
        try:
            TempW = get_value(el.LookupParameter('IGF_RLT_ZuluftTemperaturWi'))
            if TempW is None or str(int(TempW)) == '-273':TempW = ''      
                  
        except:TempW = ''
        try:iso_dicke = get_value(el.LookupParameter('IGF_X_Vorgabe_ISO_Dicke'))
        except:iso_dicke = ''
        try:iso_art = get_value(el.LookupParameter('IGF_X_Vorgabe_ISO_Art'))                  
        except:iso_art = ''

        dict_systemtyp[name] = [klasse,rgb,line,el.Id,GK,KG,KN01,KN02,BIMID,Sysname,SysKZ,AnGeAn,AnNr,PzAT,PzAZ,TempS,TempW,iso_dicke,iso_art]


    for el in rohrsystyp:
        name = el.LookupParameter('Typname').AsString()
        f = el.LineColor 
        try:rgb = resize(str(f.Red),3)+','+resize(str(f.Green),3)+','+resize(str(f.Blue),3)
        except:rgb = 'xxx,xxx,xxx'
        klasse = el.LookupParameter('Systemklassifizierung').AsString()
        try:line = doc.GetElement(el.LinePatternId).Name
        except:line = 'keine Überschreibung'
        try:
            GK = el.LookupParameter('IGF_X_Gewerkkürzel').AsString()
            if GK is None:GK = 'X'
        except:GK = 'X'
        try:
            KG = el.LookupParameter('IGF_X_Kostengruppe').AsValueString()
            if KG is None:KG = '000'
            elif len(KG) != 3:KG = '000'
        except:KG = '000'
        try:
            KN01 = el.LookupParameter('IGF_X_Kennnummer_1').AsValueString()
            if KN01 is None:KN01 = '00'
            elif len(KN01) != 2:KN01 = resize(KN01,2)
        except:KN01 = '00'
        try:
            KN02 = el.LookupParameter('IGF_X_Kennnummer_2').AsValueString()
            if KN02 is None:KN02 = '00'
            elif len(KN02) != 2:KN02 = resize(KN02,2)
        except:KN02 = '00'
        try:BIMID = GK +"_"+KG+"_"+KN01+"-"+KN02
        except:BIMID = ''
        try:
            Sysname = el.LookupParameter('IGF_X_SystemName').AsString()
            if Sysname is None:Sysname = ''
        except:Sysname = ''
        try:
            SysKZ = el.LookupParameter('IGF_X_SystemKürzel').AsString()
            if SysKZ is None:SysKZ = ''
        except:SysKZ = ''
        try:
            AnGeAn = el.LookupParameter('IGF_X_AnlagenGeräteAnzahl').AsValueString()
            if AnGeAn is None:AnGeAn = '1'
        except:AnGeAn = '1'
        try:
            AnNr = el.LookupParameter('IGF_X_AnlagenNr').AsValueString()
            if AnGeAn is None:AnGeAn = ''
        except:AnGeAn = ''
        try:
            PzAT = el.LookupParameter('IGF_X_AnlagenProzentualAnteil').AsValueString()
            if PzAT is None:PzAT = ''
            else:
                try:PzAT = str(float(PzAT)/100)+'%'
                except:pass
        except:PzAT = ''
        try:
            PzAZ = el.LookupParameter('IGF_X_AnlagenProzentualAnzahl').AsValueString()
            if PzAZ is None:PzAZ = ''            
        except:PzAZ = ''
        try:
            PzAZ = el.LookupParameter('IGF_X_AnlagenProzentualAnzahl').AsValueString()
            if PzAZ is None:PzAZ = ''            
        except:PzAZ = ''
        try:
            PzAZ = el.LookupParameter('IGF_X_AnlagenProzentualAnzahl').AsValueString()
            if PzAZ is None:PzAZ = ''            
        except:PzAZ = ''
        try:
            TempS = get_value(el.LookupParameter('IGF_RLT_ZuluftTemperaturSo'))
            if TempS is None:TempS = ''            
        except:TempS = ''
        try:
            TempW = get_value(el.LookupParameter('IGF_RLT_ZuluftTemperaturWi'))
            if TempW is None:TempW = ''            
        except:TempW = ''
        try:iso_dicke = get_value(el.LookupParameter('IGF_X_Vorgabe_ISO_Dicke'))
        except:iso_dicke = ''
        try:iso_art = get_value(el.LookupParameter('IGF_X_Vorgabe_ISO_Art'))                  
        except:iso_art = ''

        dict_systemtyp[name] = [klasse,rgb,line,el.Id,GK,KG,KN01,KN02,BIMID,Sysname,SysKZ,AnGeAn,AnNr,PzAT,PzAZ,TempS,TempW,iso_dicke,iso_art]

    return dict_systemtyp

def get_dict_workset_systemtyp():
    dict_workset = {}
    luftsysids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType().ToElementIds()
    rohrsysids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType().ToElementIds()
    for sysid in luftsysids:
        elem = doc.GetElement(sysid)
        sysname = elem.Name
        systyp = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        try:worksetlist = dict_workset[systyp]
        except:worksetlist = []
        elems = elem.DuctNetwork
        for el in elems:
            try:
                system = el.LookupParameter('Systemname').AsString()
                if sysname == system:
                    workset = el.LookupParameter('Bearbeitungsbereich').AsValueString()
                    if workset is None or workset == '':
                        continue
                    if workset not in worksetlist and workset not in ['KG4xx_Musterbereich','KG4xx_Übergreifend']:
                        worksetlist.append(workset)
            except:pass
        dict_workset[systyp] = worksetlist
        
    for sysid in rohrsysids:
        elem = doc.GetElement(sysid)
        sysname = elem.Name
        systyp = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        try:worksetlist = dict_workset[systyp]
        except:worksetlist = []
        elems = elem.PipingNetwork
        for el in elems:
            try:
                system = el.LookupParameter('Systemname').AsString()
                if sysname == system:
                    workset = el.LookupParameter('Bearbeitungsbereich').AsValueString()
                    if workset is None or workset == '':
                        continue
                    if workset not in worksetlist and workset not in ['KG4xx_Musterbereich','KG4xx_Übergreifend']:
                        worksetlist.append(workset)
            except:pass
        dict_workset[systyp] = worksetlist
    return dict_workset

def closeexcel(book,sheet,ex):
    book.Save()
    book.Close()
    Marshal.FinalReleaseComObject(sheet)
    Marshal.FinalReleaseComObject(book)
    ex.Quit()
    Marshal.FinalReleaseComObject(ex)

DICT_Workset = get_bb_dict()

DICT_LINE = get_line_dict()

dict_systemtyp = get_dict_systemtyp()

dict_workset = {}#get_dict_workset_systemtyp()

list_systemtyp_neu = dict_systemtyp.keys()[:]

DICT_MEPSYSTEMCLASS_0 = {
    'Zuluft' : DB.MEPSystemClassification.SupplyAir,
    'Abluft': DB.MEPSystemClassification.ReturnAir,
    'Fortluft': DB.MEPSystemClassification.ExhaustAir,
    'Andere Luft': DB.MEPSystemClassification.OtherAir,
    'Vorlauf': DB.MEPSystemClassification.SupplyHydronic,
    'Rücklauf': DB.MEPSystemClassification.ReturnHydronic,
    'Belüftung': DB.MEPSystemClassification.Vent,
    'Abwasser': DB.MEPSystemClassification.Sanitary,
    'Warmwasser': DB.MEPSystemClassification.DomesticHotWater,
    'Kaltwasser': DB.MEPSystemClassification.DomesticColdWater,
    'Brandschutz - Nass': DB.MEPSystemClassification.FireProtectWet,
    'Brandschutz - Trocken': DB.MEPSystemClassification.FireProtectDry,
    'Brandschutz - Vorgesteuert': DB.MEPSystemClassification.FireProtectPreaction,
    'Brandschutz - Andere': DB.MEPSystemClassification.FireProtectOther,
    'Sonstige': DB.MEPSystemClassification.OtherPipe,
    'Rohrformteil': DB.MEPSystemClassification.Fitting,
    'Global': DB.MEPSystemClassification.Global,
}

DICT_MEPSYSTEMCLASS_1 = {
    DB.MEPSystemClassification.SupplyAir:'Zuluft',
    DB.MEPSystemClassification.ReturnAir:'Abluft',
    DB.MEPSystemClassification.ExhaustAir:'Fortluft',
    DB.MEPSystemClassification.OtherAir:'Andere Luft',
    DB.MEPSystemClassification.SupplyHydronic:'Vorlauf',
    DB.MEPSystemClassification.ReturnHydronic:'Rücklauf',
    DB.MEPSystemClassification.Vent:'Belüftung',
    DB.MEPSystemClassification.Sanitary:'Abwasser',
    DB.MEPSystemClassification.DomesticHotWater:'Warmwasser',
    DB.MEPSystemClassification.DomesticColdWater:'Kaltwasser',
    DB.MEPSystemClassification.FireProtectWet:'Brandschutz - Nass',
    DB.MEPSystemClassification.FireProtectDry:'Brandschutz - Trocken',
    DB.MEPSystemClassification.FireProtectOther:'Brandschutz - Andere',
    DB.MEPSystemClassification.FireProtectPreaction:'Brandschutz - Vorgesteuert',
    DB.MEPSystemClassification.Fitting:'Rohrformteil',
    DB.MEPSystemClassification.OtherPipe:'Sonstige',
    DB.MEPSystemClassification.Global:'Global',
}

class Exceldaten(object):
    def __init__(self):
        self.checked = False
        self.bb = False
        self.erstellen = False
        self.export = False
        self.changeable = True
        self.changeable_import = False
        self.elementid = None
        self.canexport =True
        self.braucht = False

        self.status = 'nicht vorhanden'
        self.abweichung_status = BLACK
        self.abweichung_hinweis = ''

        self.klasse = None
        self.abweichung_klasse = BLACK
        self.revit_klasse = None

        self.farbe = None
        self.abweichung_farbe = BLACK
        self.revit_farbe = None

        self.linie = None
        self.abweichung_linie = BLACK
        self.revit_linie = None

        self.Systemname = None
        
        self.GK = None
        self.KG = None
        self.KN01 = None
        self.KN02 = None
        self.AnNr = None
        self.AnGeAn = None
        self.PzAT = None
        self.PzAZ = None
        self.SysKZ = None
        self.Sysname = None
        self.BIMID = None
        self.TempW = None
        self.TempS = None
        self.ISO_Dicke= None
        self.ISO_Art = None
        self.Workset = None

        self.revit_GK = None
        self.revit_KG = None
        self.revit_KN01 = None
        self.revit_KN02 = None
        self.revit_AnNr = None
        self.revit_AnGeAn = None
        self.revit_PzAT = None
        self.revit_PzAZ = None
        self.revit_SysKZ = None
        self.revit_Sysname = None
        self.revit_BIMID = None
        self.revit_TempW = None
        self.revit_TempS = None
        self.revit_Workset = None
        self.revit_ISO_Dicke= None
        self.revit_ISO_Art = None

        self.row = None
        self.col_ISO_Dicke= None
        self.col_ISO_Art = None
        self.col_systyp = None
        self.col_workset_aus = None
        self.col_anlegen = None
        self.col_status = None
        self.col_klasse = None
        self.col_farbe = None
        self.col_linie = None
        self.col_GK = None
        self.col_KG = None
        self.col_KN01 = None
        self.col_KN02 = None
        self.col_AnNr = None
        self.col_AnGeAn = None
        self.col_PzAT = None
        self.col_PzAZ = None
        self.col_SysKZ = None
        self.col_Sysname = None
        self.col_BIMID = None
        self.col_TempW = None
        self.col_TempS = None

Liste_Luft = ObservableCollection[Exceldaten]()
Liste_Rohr = ObservableCollection[Exceldaten]()
Liste_Luft_braucht = ObservableCollection[Exceldaten]()
Liste_Rohr_braucht = ObservableCollection[Exceldaten]()
Liste_All = ObservableCollection[Exceldaten]()

def datenlesen(filepath):
    ex = Excel.ApplicationClass()
    book = ex.Workbooks.Open(filepath)
    sheet = book.Worksheets['Systeme']
    rows = sheet.UsedRange.Rows.Count
    cols = sheet.UsedRange.Columns.Count
    for col in range(1,cols+1):
        paramname = sheet.Cells[1, col].Value2
        if paramname == 'Systemtyp':
            col_sysname = col
        elif paramname == 'IGF_X_Gewerkkürzel':
            col_GK = col
        elif paramname == 'IGF_X_Vorgabe_ISO_Dicke':
            col_ISO_Dicke = col
        elif paramname == 'IGF_X_Vorgabe_ISO_Art':
            col_ISO_Art = col
        elif paramname == 'Projekt benötig':
            col_braucht = col
        elif paramname == 'IGF_X_Kostengruppe':
            col_KG = col
        elif paramname == 'IGF_X_Kennnummer_1':
            col_KN01 = col
        elif paramname == 'IGF_X_Kennnummer_2':
            col_KN02 = col
        elif paramname == 'IGF_X_BIM-ID':
            col_BIMID = col
        elif paramname == 'IGF_X_AnlagenNr':
            col_annr = col
        elif paramname == 'IGF_X_AnlagenGeräteAnzahl':
            col_gean = col
        elif paramname == 'IGF_RLT_ZuluftTemperaturWi':
            col_TempW = col
        elif paramname == 'IGF_RLT_ZuluftTemperaturSo':
            col_TempS = col
        elif paramname == 'IGF_X_AnlagenProzentualAnteil':
            col_PzAT = col
        elif paramname == 'IGF_X_AnlagenProzentualAnzahl':
            col_PzAZ = col
        elif paramname == 'IGF_X_SystemKürzel':
            col_syskz = col
        elif paramname == 'IGF_X_SystemName':
            col_sysna = col
        elif paramname == 'Bearbeitungsbereich_nach':
            col_workset = col
        elif paramname == 'Bearbeitungsbereich_aus':
            col_workset_aus = col
        ## Systeme erstellen
        elif paramname == 'Systemklasse':
            col_klass = col
        elif paramname == 'Farbe':
            col_farbe = col
        elif paramname == 'Liniemüster':
            col_linie = col
        ## Systeme prüfen
        elif paramname == 'Angelegt':
            col_anlegen = col
        elif paramname == 'Status':
            col_status = col

    if rows > 302:
        rows = 302

    for row in range(32,rows+1):
        sysname = sheet.Cells[row, col_sysname].Value2
        if not sysname:
            continue
        tempclass = Exceldaten()
        tempclass.row = row
        tempclass.col_systyp = col_sysname
        tempclass.col_anlegen = col_anlegen
        tempclass.col_status = col_status
        tempclass.col_workset_aus = col_workset_aus
        tempclass.col_farbe = col_farbe
        tempclass.col_linie = col_linie
        tempclass.col_klasse = col_klass
        tempclass.col_GK = col_GK
        tempclass.col_KG = col_KG
        tempclass.col_KN01 = col_KN01
        tempclass.col_KN02 = col_KN02
        tempclass.col_BIMID = col_BIMID
        tempclass.col_AnGeAn = col_gean
        tempclass.col_AnNr = col_annr
        tempclass.col_PzAT = col_PzAT
        tempclass.col_PzAZ = col_PzAZ
        tempclass.col_SysKZ = col_syskz
        tempclass.col_Sysname = col_sysna
        tempclass.col_TempW = col_TempW
        tempclass.col_TempS = col_TempS
        tempclass.col_ISO_Art = col_ISO_Art
        tempclass.col_ISO_Dicke = col_ISO_Dicke
        
        GK = sheet.Cells[row, col_GK].Value2
        braucht = sheet.Cells[row, col_braucht].Value2
        if GK == None or GK == '':GK = '-'
        try:KG = str(int(sheet.Cells[row, col_KG].Value2))
        except:KG = '000'
        try:
            KN01 = str(int(sheet.Cells[row, col_KN01].Value2))
            if len(KN01) == 1:KN01 = '0' + KN01
        except:KN01 = '00'
        try:
            KN02 = str(int(sheet.Cells[row, col_KN02].Value2))
            if len(KN02) == 1:KN02 = '0' + KN02
        except:KN02 = '00'
        try:bimid = GK +"_"+KG+"_"+KN01+"-"+KN02
        except:bimid = ''

        # bimid = sheet.Cells[row, col_bimid].Value2
        annr = sheet.Cells[row, col_annr].Value2
        gean = sheet.Cells[row, col_gean].Value2
        TempW = sheet.Cells[row, col_TempW].Value2
        TempS = sheet.Cells[row, col_TempS].Value2
        PzAT = sheet.Cells[row, col_PzAT].Value2
        PzAZ = sheet.Cells[row, col_PzAZ].Value2
        syskz = sheet.Cells[row, col_syskz].Value2
        sysna = sheet.Cells[row, col_sysna].Value2
        workset = sheet.Cells[row, col_workset].Value2
        anlegen = sheet.Cells[row, col_anlegen].Value2
        klasse = sheet.Cells[row, col_klass].Value2
        farbe = sheet.Cells[row, col_farbe].Value2
        linie = sheet.Cells[row, col_linie].Value2
        iso_dicke = sheet.Cells[row, col_ISO_Dicke].Value2
        iso_art = sheet.Cells[row, col_ISO_Art].Value2
        try:
            if braucht.upper().find('JA') != -1:
                tempclass.braucht = True
        except:tempclass.braucht = False

        tempclass.AnNr = annr
        tempclass.AnGeAn = gean
        tempclass.SysKZ = syskz
        tempclass.Sysname = sysna
        tempclass.Systemname = sysname
        tempclass.GK = GK
        tempclass.KG = KG
        tempclass.KN01 = KN01
        tempclass.KN02 = KN02
        tempclass.TempW = TempW
        tempclass.TempS = TempS
        tempclass.PzAT = PzAT
        tempclass.PzAZ = PzAZ
        tempclass.BIMID = bimid
        tempclass.Workset = workset
        if anlegen == '+':tempclass.erstellen = True
        tempclass.klasse = klasse
        tempclass.farbe = farbe
        tempclass.linie = linie
        tempclass.ISO_Dicke = iso_dicke
        tempclass.ISO_Art = iso_art
        
        if sysname in dict_systemtyp.keys():
            list_systemtyp_neu.remove(sysname)
            tempclass.changeable_import = True
            tempclass.revit_klasse = dict_systemtyp[sysname][0]
            tempclass.revit_farbe = dict_systemtyp[sysname][1]
            tempclass.revit_linie = dict_systemtyp[sysname][2]
            tempclass.elementid = dict_systemtyp[sysname][3]
            tempclass.revit_GK = dict_systemtyp[sysname][4]
            tempclass.revit_KG = dict_systemtyp[sysname][5]
            tempclass.revit_KN01 = dict_systemtyp[sysname][6]
            tempclass.revit_KN02 = dict_systemtyp[sysname][7]
            tempclass.revit_BIMID= dict_systemtyp[sysname][8]
            tempclass.revit_Sysname = dict_systemtyp[sysname][9]
            tempclass.revit_SysKZ = dict_systemtyp[sysname][10]
            tempclass.revit_AnGeAn = dict_systemtyp[sysname][11]
            tempclass.revit_AnNr= dict_systemtyp[sysname][12]
            tempclass.revit_PzAT = dict_systemtyp[sysname][13]
            tempclass.revit_PzAZ = dict_systemtyp[sysname][14]
            tempclass.revit_TempS = dict_systemtyp[sysname][15]
            tempclass.revit_TempW = dict_systemtyp[sysname][16]
            tempclass.revit_ISO_Dicke = dict_systemtyp[sysname][17]
            tempclass.revit_ISO_Art= dict_systemtyp[sysname][18]
            if tempclass.klasse == tempclass.revit_klasse and tempclass.revit_farbe == tempclass.farbe and tempclass.revit_linie == tempclass.linie:
                tempclass.status = 'vorhanden'
                tempclass.abweichung_hinweis = 'vorhanden'
                tempclass.changeable = False
                tempclass.erstellen = False
                
            else:
                tempclass.status = 'geändert'
                if tempclass.klasse != tempclass.revit_klasse:
                    tempclass.abweichung_hinweis += ' Systemklasse passt nicht.'
                    tempclass.abweichung_status = RED
                    tempclass.abweichung_klasse = RED
                    tempclass.changeable = False
                if tempclass.revit_farbe != tempclass.farbe:
                    tempclass.abweichung_hinweis += ' Farbe passt nicht.'
                    tempclass.abweichung_status = RED
                    tempclass.abweichung_farbe = RED
                if tempclass.revit_linie != tempclass.linie:
                    tempclass.abweichung_hinweis += ' Linie passt nicht.'
                    tempclass.abweichung_status = RED
                    tempclass.abweichung_linie = RED
        else:
            tempclass.canexport = False
        if not (tempclass.klasse and tempclass.farbe and tempclass.linie):
            tempclass.changeable = False

        if sysname in dict_workset.keys():
            revitworkset = ''
            if len(dict_workset[sysname]) > 0:
                for workset in dict_workset[sysname]:
                    revitworkset = revitworkset + workset + ', '
            
            tempclass.revit_Workset = revitworkset[:-2]

        Liste_All.Add(tempclass)
        if tempclass.GK == 'R':
            Liste_Luft.Add(tempclass)
            if tempclass.braucht:Liste_Luft_braucht.Add(tempclass)

        else:
            Liste_Rohr.Add(tempclass)
            if tempclass.braucht:Liste_Rohr_braucht.Add(tempclass)


    closeexcel(book,sheet,ex)

    n = 303
    for sysname in list_systemtyp_neu:
        n += 1
        tempclass = Exceldaten()
        tempclass.Systemname = sysname
        tempclass.row = n
        tempclass.col_anlegen = col_anlegen
        tempclass.col_systyp = col_sysname
        tempclass.col_status = col_status
        tempclass.col_workset_aus = col_workset_aus
        tempclass.col_farbe = col_farbe
        tempclass.col_linie = col_linie
        tempclass.col_klasse = col_klass
        tempclass.col_GK = col_GK
        tempclass.col_KG = col_KG
        tempclass.col_KN01 = col_KN01
        tempclass.col_KN02 = col_KN02
        tempclass.col_BIMID = col_BIMID
        tempclass.col_AnGeAn = col_gean
        tempclass.col_AnNr = col_annr
        tempclass.col_PzAT = col_PzAT
        tempclass.col_PzAZ = col_PzAZ
        tempclass.col_SysKZ = col_syskz
        tempclass.col_Sysname = col_sysna
        tempclass.col_TempW = col_TempW
        tempclass.col_TempS = col_TempS
        tempclass.revit_klasse = dict_systemtyp[sysname][0]
        tempclass.revit_farbe = dict_systemtyp[sysname][1]
        tempclass.revit_linie = dict_systemtyp[sysname][2]
        tempclass.elementid = dict_systemtyp[sysname][3]
        tempclass.revit_GK = dict_systemtyp[sysname][4]
        tempclass.revit_KG = dict_systemtyp[sysname][5]
        tempclass.revit_KN01 = dict_systemtyp[sysname][6]
        tempclass.revit_KN02 = dict_systemtyp[sysname][7]
        tempclass.revit_BIMID= dict_systemtyp[sysname][8]
        tempclass.revit_Sysname = dict_systemtyp[sysname][9]
        tempclass.revit_SysKZ = dict_systemtyp[sysname][10]
        tempclass.revit_AnGeAn = dict_systemtyp[sysname][11]
        tempclass.revit_AnNr= dict_systemtyp[sysname][12]
        tempclass.revit_PzAT = dict_systemtyp[sysname][13]
        tempclass.revit_PzAZ = dict_systemtyp[sysname][14]
        tempclass.revit_TempS = dict_systemtyp[sysname][15]
        tempclass.revit_TempW = dict_systemtyp[sysname][16]
        tempclass.revit_ISO_Dicke = dict_systemtyp[sysname][17]
        tempclass.revit_ISO_Art= dict_systemtyp[sysname][18]
        tempclass.klasse = dict_systemtyp[sysname][0]
        tempclass.farbe = dict_systemtyp[sysname][1]
        tempclass.linie = dict_systemtyp[sysname][2]
        tempclass.GK = dict_systemtyp[sysname][4]
        tempclass.KG = dict_systemtyp[sysname][5]
        tempclass.KN01 = dict_systemtyp[sysname][6]
        tempclass.KN02 = dict_systemtyp[sysname][7]
        tempclass.BIMID= dict_systemtyp[sysname][8]
        tempclass.Sysname = dict_systemtyp[sysname][9]
        tempclass.SysKZ = dict_systemtyp[sysname][10]
        tempclass.AnGeAn = dict_systemtyp[sysname][11]
        tempclass.AnNr= dict_systemtyp[sysname][12]
        tempclass.PzAT = dict_systemtyp[sysname][13]
        tempclass.PzAZ = dict_systemtyp[sysname][14]
        tempclass.TempS = dict_systemtyp[sysname][15]
        tempclass.TempW = dict_systemtyp[sysname][16]
        tempclass.ISO_Dicke = dict_systemtyp[sysname][17]
        tempclass.ISO_Art= dict_systemtyp[sysname][18]
        tempclass.changeable = False
        tempclass.changeable_import = False
        if sysname in dict_workset.keys():
            revitworkset = ''
            if len(dict_workset[sysname]) > 0:
                for e in dict_workset[sysname]:
                    revitworkset += e +', '
            tempclass.revit_Workset = revitworkset[:-2]
            tempclass.Workset = revitworkset[:-2]
        tempclass.status = 'Revit'
        tempclass.abweichung_hinweis = 'Nur in Revit'
        tempclass.abweichung_status = RED
        Liste_All.Add(tempclass)
        if tempclass.GK == 'R':
            Liste_Luft.Add(tempclass)
        else:
            Liste_Rohr.Add(tempclass)

def daten_lesen_main():
    try:
        Adresse = bimid_config.bimid
        if os.path.exists(Adresse):
            try:
                datenlesen(Adresse)
            except Exception as e:
                logger.error(e)
    except:
        pass

daten_lesen_main()

# ExcelBimId GUI
class ExcelBimId(WPFWindow):
    def __init__(self, xaml_file_name,liste_Luft,liste_Rohr,liste_Luft_braucht,liste_Rohr_braucht):
        self.liste_Luft = liste_Luft
        self.liste_Rohr = liste_Rohr
        self.liste_Luft_braucht = liste_Luft_braucht
        self.liste_Rohr_braucht = liste_Rohr_braucht
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[Exceldaten]()
        self.altdatagrid = None
        self.read_config()
        self.datenimport_status = False 
        self.dataexport = False

        try:
            self.dataGrid.ItemsSource = self.liste_Luft_braucht
            self.altdatagrid = self.liste_Luft_braucht
            self.backAll()
            self.click(self.luft)
        except Exception as e:
            logger.error(e)

        self.systemsuche.TextChanged += self.search_txt_changed
        self.Adresse.TextChanged += self.excel_changed


    
    def changestyle(self,sender,args):
        if sender.IsChecked:
            if self.rohr.FontWeight == FontWeights.Bold:
                self.dataGrid.ItemsSource = self.liste_Rohr_braucht
                self.altdatagrid = self.liste_Rohr_braucht
            else:
                self.dataGrid.ItemsSource = self.liste_Luft_braucht
                self.altdatagrid = self.liste_Luft_braucht
        else:
            if self.rohr.FontWeight == FontWeights.Bold:
                self.dataGrid.ItemsSource = self.liste_Rohr
                self.altdatagrid = self.liste_Rohr
            else:
                self.dataGrid.ItemsSource = self.liste_Luft
                self.altdatagrid = self.liste_Luft
        self.systemsuche.Text = ''

    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic
    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal
    def backAll(self):
        self.back(self.luft)
        self.back(self.rohr)

    def rohre(self,sender,args):
        self.backAll()
        self.click(self.rohr)
        if self.zeigenall.IsChecked:

            self.dataGrid.ItemsSource = self.liste_Rohr_braucht
            self.altdatagrid = self.liste_Rohr_braucht
        else:
            self.dataGrid.ItemsSource = self.liste_Rohr
            self.altdatagrid = self.liste_Rohr
        self.Sys_2.Visibility = Visibility.Hidden
        self.Sys_3.Visibility = Visibility.Hidden
        self.Sys_4.Visibility = Visibility.Hidden
        self.Sys_5.Visibility = Visibility.Hidden
        self.systemsuche.Text = ''

    def luftung(self,sender,args):
        self.backAll()
        self.click(self.luft)
        if self.zeigenall.IsChecked:
            self.dataGrid.ItemsSource = self.liste_Luft_braucht
            self.altdatagrid = self.liste_Luft_braucht
        else:
            self.dataGrid.ItemsSource = self.liste_Luft
            self.altdatagrid = self.liste_Luft
        if not self.detail_anlagen.IsChecked:
            self.Sys_2.Visibility = Visibility.Hidden
            self.Sys_3.Visibility = Visibility.Hidden
            self.Sys_4.Visibility = Visibility.Hidden
            self.Sys_5.Visibility = Visibility.Hidden
        else:
            self.Sys_2.Visibility = Visibility.Visible
            self.Sys_3.Visibility = Visibility.Visible
            self.Sys_4.Visibility = Visibility.Visible
            self.Sys_5.Visibility = Visibility.Visible

        self.systemsuche.Text = ''

    def read_config(self):
        try:
            if os.path.exists(bimid_config.bimid):
                self.Adresse.Text = str(bimid_config.bimid)
            else:self.Adresse.Text = bimid_config.bimid = ""

        except:
            self.Adresse.Text = bimid_config.bimid = ""

    def write_config(self):
        bimid_config.bimid = self.Adresse.Text
        script.save_config()

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.systemsuche.Text.upper()
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.Systemname.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def excel_changed(self, sender, args):
        Liste_Luft.Clear()
        Liste_Rohr.Clear()
        try:
            datenlesen(self.Adresse.Text)
        except:
            pass
        self.liste_Luft = Liste_Luft
        self.dataGrid.ItemsSource = Liste_Luft
        self.altdatagrid = self.liste_Luft
        self.backAll()
        self.click(self.luft)

    def durchsuchen(self,sender,args):
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "Excel suchen"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            self.Adresse.Text = dialog.FileName
        self.write_config()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = True
        self.dataGrid.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = False
        self.dataGrid.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.checked
            item.checked = not value
        self.dataGrid.Items.Refresh()
    def checkallbb(self,sender,args):
        for item in self.dataGrid.Items:
            item.bb = True
        self.dataGrid.Items.Refresh()       

    def uncheckallbb(self,sender,args):
        for item in self.dataGrid.Items:
            item.bb = False
        self.dataGrid.Items.Refresh()

    def toggleallbb(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.bb
            item.bb = not value
        self.dataGrid.Items.Refresh()

    def datenimport(self,sender,args):
        self.datenimport_status = True
        self.Close()

    def system_erstellen(self,sender,args):
        start = False
        for el in Liste_All:
            if el.erstellen:
                start = True
                break
        if not start:return
        t = DB.Transaction(doc,'System erstellen')
        t.Start()
        ex = Excel.ApplicationClass()
        book = ex.Workbooks.Open(bimid_config.bimid)
        sheet = book.Worksheets['Systeme']
        for el in self.liste_Rohr:
            if el.erstellen:
                el.abweichung_status = BLACK
                el.abweichung_hinweis = 'vorhanden'
                el.revit_klasse = el.klasse
                el.abweichung_klasse = BLACK
                el.revit_linie = el.linie
                el.abweichung_klasse = BLACK
                el.revit_farbe = el.farbe
                el.abweichung_farbe = BLACK
                el.changeable_import = True

                try:
                    r = int(el.farbe[:3])
                    g = int(el.farbe[4:7])
                    b = int(el.farbe[8:11])
                except:logger.error('Farbeformat in Excel passt nicht, Zeile {}, Spalte: {}'.format(el.row,el.col_farbe))

                try:
                    if el.linie != 'keine Überschreibung':linie= DICT_LINE[el.linie]
                except:logger.error('Linievorlage {} nicht vorhanden'.format(el.linie))

                if el.status == 'nicht vorhanden':
                    systyp = DB.Plumbing.PipingSystemType.Create(doc,DICT_MEPSYSTEMCLASS_0[el.klasse],el.Systemname)
                    el.elementid = systyp.Id
                    
                elif el.status == 'geändert':
                    systyp = doc.GetElement(el.elementid)
                
                try:
                    farbe = DB.Color(System.Byte(r),System.Byte(g),System.Byte(b))
                    systyp.LineColor = farbe
                except:logger.error('Fehler bei der Einstellung der Farbe von Systemtyp {}'.format(el.Systemname))
                try:
                    if el.linie != 'keine Überschreibung':systyp.LinePatternId = linie
                except:logger.error('Fehler bei der Einstellung der Linie von Systemtyp {}'.format(el.Systemname))
                el.status = 'vorhanden'
                sheet.Cells[el.row,el.col_status] = 'vorhanden'
                el.erstellen = False
                el.changeable = False


        for el in self.liste_Luft:
            if el.erstellen:
                el.abweichung_status = BLACK
                el.abweichung_hinweis = 'vorhanden'
                el.revit_klasse = el.klasse
                el.abweichung_klasse = BLACK
                el.revit_linie = el.linie
                el.abweichung_klasse = BLACK
                el.revit_farbe = el.farbe
                el.abweichung_farbe = BLACK
                el.erstellen = False
                el.changeable_import = True
                try:
                    r = int(el.farbe[:3])
                    g = int(el.farbe[4:7])
                    b = int(el.farbe[8:11])
                except:logger.error('Farbeformat in Excel passt nicht, Zeile {}, Spalte: {}'.format(el.row,el.col_farbe))
                try:
                    if el.linie != 'keine Überschreibung':linie= DICT_LINE[el.linie]
                except:logger.error('Linievorlage {} nicht vorhanden'.format(el.linie))
                    
                if el.status == 'nicht vorhanden':
                    systyp = DB.Mechanical.MechanicalSystemType.Create(doc,DICT_MEPSYSTEMCLASS_0[el.klasse],el.Systemname)
                    el.elementid = systyp.Id
                    
                elif el.status == 'geändert':
                    systyp = doc.GetElement(el.elementid)

                try:
                    farbe = DB.Color(System.Byte(r),System.Byte(g),System.Byte(b))
                    systyp.LineColor = farbe
                except:logger.error('Fehler bei der Einstellung der Farbe von Systemtyp {}'.format(el.Systemname))
                try:
                    if el.linie != 'keine Überschreibung':systyp.LinePatternId = linie
                except:logger.error('Fehler bei der Einstellung der Linie von Systemtyp {}'.format(el.Systemname))

                el.status = 'vorhanden'
                sheet.Cells[el.row,el.col_status] = 'vorhanden'
                el.erstellen = False
                el.changeable = False
            
        
        doc.Regenerate()
        t.Commit()

        closeexcel(book,sheet,ex)
        self.dataGrid.Items.Refresh()


    def datenexport(self,sender,args):
        start = False
        for el in Liste_All:
            if el.export:
                start = True
                break
        if not start:return
        ex = Excel.ApplicationClass()
        book = ex.Workbooks.Open(bimid_config.bimid)
        sheet = book.Worksheets['Systeme']
        try:
            for el in Liste_All:
                if not el.export:continue
                el.export = False
                el.status = 'vorhanden'
                el.abweichung_status = BLACK
                el.abweichung_hinweis = 'vorhanden'
                
                sheet.Cells[el.row,el.col_status] = 'vorhanden'
                sheet.Cells[el.row,el.col_systyp] = el.Systemname
                sheet.Cells[el.row,el.col_klasse] = el.revit_klasse
                sheet.Cells[el.row,el.col_farbe] = str(el.revit_farbe)
                sheet.Cells[el.row,el.col_linie] = str(el.revit_linie)
                sheet.Cells[el.row,el.col_anlegen] = ''
                sheet.Cells[el.row,el.col_BIMID] = el.revit_BIMID
                sheet.Cells[el.row,el.col_GK] = el.revit_GK
                sheet.Cells[el.row,el.col_KG] = el.revit_KG
                sheet.Cells[el.row,el.col_KN01] = el.revit_KN01
                sheet.Cells[el.row,el.col_KN02] = el.revit_KN02

                sheet.Cells[el.row,el.col_AnNr] = el.revit_AnNr
                sheet.Cells[el.row,el.col_AnGeAn] = el.revit_AnGeAn
                sheet.Cells[el.row,el.col_TempW] = el.revit_TempW
                sheet.Cells[el.row,el.col_TempS] = el.revit_TempS
                sheet.Cells[el.row,el.col_PzAT] = el.revit_PzAT
                sheet.Cells[el.row,el.col_PzAZ] = el.revit_PzAZ
                sheet.Cells[el.row,el.col_SysKZ] = el.revit_SysKZ
                sheet.Cells[el.row,el.col_Sysname] = el.revit_Sysname
                sheet.Cells[el.row,el.col_ISO_Dicke] = el.revit_ISO_Dicke
                sheet.Cells[el.row,el.col_ISO_Art] = el.revit_ISO_Art
                sheet.Cells[el.row,el.col_workset_aus] = el.revit_Workset
            closeexcel(book,sheet,ex)
            self.dataGrid.Items.Refresh()
        except Exception as e:
            logger.error(e)
            closeexcel(book,sheet,ex)


    def detail_GD_einaus(self,sender,args):
        ischecked = sender.IsChecked
        if ischecked:
            self.GD_1.Visibility = Visibility.Visible
            self.GD_2.Visibility = Visibility.Visible
            self.GD_3.Visibility = Visibility.Visible
        else:
            self.GD_1.Visibility = Visibility.Hidden
            self.GD_2.Visibility = Visibility.Hidden
            self.GD_3.Visibility = Visibility.Hidden
    def detail_iso_einaus(self,sender,args):
        ischecked = sender.IsChecked
        if ischecked:
            self.iso_1.Visibility = Visibility.Visible
            self.iso_2.Visibility = Visibility.Visible
        else:
            self.iso_1.Visibility = Visibility.Hidden
            self.iso_2.Visibility = Visibility.Hidden

    def detail_BIMID_einaus(self,sender,args):
        ischecked = sender.IsChecked
        if ischecked:
            self.BIMID_1.Visibility = Visibility.Visible
            self.BIMID_2.Visibility = Visibility.Visible
            self.BIMID_3.Visibility = Visibility.Visible
            self.BIMID_4.Visibility = Visibility.Visible
        else:
            self.BIMID_1.Visibility = Visibility.Hidden
            self.BIMID_2.Visibility = Visibility.Hidden
            self.BIMID_3.Visibility = Visibility.Hidden
            self.BIMID_4.Visibility = Visibility.Hidden

    def detail_SYS_einaus(self,sender,args):
        ischecked = sender.IsChecked
        if ischecked:
            self.Sys_1.Visibility = Visibility.Visible
            if self.luft.FontWeight == FontWeights.Bold:
                self.Sys_2.Visibility = Visibility.Visible
                self.Sys_3.Visibility = Visibility.Visible
                self.Sys_4.Visibility = Visibility.Visible
                self.Sys_5.Visibility = Visibility.Visible
            else:
                self.Sys_2.Visibility = Visibility.Hidden
                self.Sys_3.Visibility = Visibility.Hidden
                self.Sys_4.Visibility = Visibility.Hidden
                self.Sys_5.Visibility = Visibility.Hidden

        else:
            self.Sys_1.Visibility = Visibility.Hidden
            self.Sys_2.Visibility = Visibility.Hidden
            self.Sys_3.Visibility = Visibility.Hidden
            self.Sys_4.Visibility = Visibility.Hidden
            self.Sys_5.Visibility = Visibility.Hidden

    def exportchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.dataGrid.SelectedItem is not None:
            try:
                if sender.DataContext in self.dataGrid.SelectedItems:
                    for item in self.dataGrid.SelectedItems:
                        try:
                            item.export = Checked
                        except:
                            pass                 
                    
                    self.dataGrid.Items.Refresh()                       
                else:
                    pass
            except:
                pass    
    def erstellenchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.dataGrid.SelectedItem is not None:
            try:
                if sender.DataContext in self.dataGrid.SelectedItems:
                    for item in self.dataGrid.SelectedItems:
                        try:
                            if item.changeable:
                                item.erstellen = Checked
                            else:
                                item.erstellen = False
                        except:
                            pass
                    self.dataGrid.Items.Refresh()                       
                else:
                    pass
            except:
                pass 
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.dataGrid.SelectedItem is not None:
            try:
                if sender.DataContext in self.dataGrid.SelectedItems:
                    for item in self.dataGrid.SelectedItems:
                        try:
                            if item.changeable_import:
                                item.checked = Checked
                            else:
                                item.checked = False
                        except:
                            pass
                    self.dataGrid.Items.Refresh()                       
                else:
                    pass
            except:
                pass    

    def bbchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.dataGrid.SelectedItem is not None:
            try:
                if sender.DataContext in self.dataGrid.SelectedItems:
                    for item in self.dataGrid.SelectedItems:
                        try:
                            if item.changeable_import:
                                item.bb = Checked
                            else:
                                item.bb = False
                        except:
                            pass
                    self.dataGrid.Items.Refresh()                       
                else:
                    pass
            except:
                pass      


windowExcelBimId = ExcelBimId("Window.xaml",Liste_Luft,Liste_Rohr,Liste_Luft_braucht,Liste_Rohr_braucht)

try:
    windowExcelBimId.ShowDialog()
except Exception as e:
    logger.error(e)
    windowExcelBimId.Close()
    script.exit()

muster_bb_bearbeiten = windowExcelBimId.muster.IsChecked
sichtbar = windowExcelBimId.sichtbar.IsChecked
uebergreifend_beruecksichtigt = windowExcelBimId.mehrgewerke.IsChecked

def create_bb():
    worksset_Excel = [e.Workset for e in Liste_All if e.checked or e.bb]
    worksset_Excel = list(worksset_Excel)
    worksset_Excel.append('KG4xx_Übergreifend')
    fehlendeworkset = []
    if len(worksset_Excel) > 0:
        for item in worksset_Excel:
            if not item in DICT_Workset.keys():
                fehlendeworkset.append(item)
    fehlendeworkset = set(fehlendeworkset)
    fehlendeworkset = list(fehlendeworkset)

    # Bearbeitungsbereich erstellen
    if len(fehlendeworkset) > 0:
        logger.error('folgende Bearbeitungsbereiche fehlt:')
        for e in fehlendeworkset:
            logger.error(e)

        if forms.alert("fehlende Bearbeitungsbereiche erstellen?", ok=False, yes=True, no=True):
            t = DB.Transaction(doc)
            t.Start('Bearbeitungsbereich erstellen')
            for el in fehlendeworkset:
                logger.info(30*'-')
                logger.info(el)
                try:
                    item = DB.Workset.Create(doc,el)
                    DB.WorksetDefaultVisibilitySettings.SetWorksetVisibility(DB.WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc),item.Id,sichtbar)
                    DICT_Workset[el] = item.Id.ToString()
                    logger.info('Bearbeitungsbereich {} erstellt'.format(el))
                except Exception as e:
                    logger.error(e)
            doc.Regenerate()
            t.Commit()
            t.Dispose()

create_bb()

class MEP_System:
    def __init__(self,System_Excel):
        self.System_Excel = System_Excel
        self.bimid = System_Excel.checked
        self.systemname = System_Excel.Systemname
        self.GK = System_Excel.GK
        self.KG = System_Excel.KG
        self.KN01 = System_Excel.KN01
        self.KN02 = System_Excel.KN02
        self.BIMID = System_Excel.BIMID
        self.AnNr = System_Excel.AnNr
        self.AnGeAn = System_Excel.AnGeAn
        self.SysKZ = System_Excel.SysKZ
        self.Sysname = System_Excel.Sysname
        self.PzAT = System_Excel.PzAT
        self.PzAZ = System_Excel.PzAZ
        self.TempW = System_Excel.TempW
        self.TempS = System_Excel.TempS
        self.Workset = System_Excel.Workset
        self.bb = System_Excel.bb
        self.ISO_Dicke = System_Excel.ISO_Dicke
        self.ISO_Art = System_Excel.ISO_Art
        

        try:
            self.PzAT = float(self.PzAT[:self.PzAT.find('%')])/100
        except:
            pass
        
        self.liste_system = []
        self.liste_bauteile = []
        self.typ = None
        self.dict_bauteile_ueber = {}
        self.dict_external = {}
        
    def get_elemente(self):
        for el in self.liste_system:
            systemid = el.Id.ToString()
            try:
                elemente = el.DuctNetwork
            except:
                elemente = el.PipingNetwork

            for elem in elemente:
                cate = elem.Category.Id.ToString()
                # Leitung
                if cate in ['-2008000','-2008020','-2008050','-2008044']:
                    try:
                        if elem.MEPSystem.Id.ToString() == systemid:
                            if elem.Id.ToString() not in self.liste_bauteile:
                                self.liste_bauteile.append(elem.Id.ToString())
                    except:
                        pass
                # Auslass, Sprinkler
                elif cate in ['-2008013','-2008099']:
                    try:
                        if list(elem.MEPModel.ConnectorManager.Connectors)[0].MEPSystem.Id.ToString() == systemid:
                            if elem.Id.ToString() not in self.liste_bauteile:
                                self.liste_bauteile.append(elem.Id.ToString())
                    except:
                        pass
                # Bauteile
                elif cate in ['-2008010','-2008016','-2001140','-2008049','-2008055','-2001160']:
                    
                    conns = elem.MEPModel.ConnectorManager.Connectors
                    In = {}
                    Out = {}
                    Unverbunden = {}
                    for conn in conns:
                        if conn.IsConnected:
                            if conn.Direction.ToString() == 'In':
                                In[conn.Id] = conn
                            else:
                                Out[conn.Id] = conn
                        else:
                            Unverbunden[conn.Id] = conn
                    sorted(In)
                    sorted(Out)
                    sorted(Unverbunden)
                    conns = In.values()[:]
                    connouts = Out.values()[:]
                    connunvers = Unverbunden.values()[:]
                    conns.extend(connouts)
                    conns.extend(connunvers)
                    if not uebergreifend_beruecksichtigt:
                        try:
                            for conn in conns:
                                if not conn.MEPSystem:
                                    continue
                                else:
                                    if conn.MEPSystem.Id.ToString() == systemid:
                                        if elem.Id.ToString() not in self.liste_bauteile:
                                            self.liste_bauteile.append(elem.Id.ToString())
                                    break
                                
                        except:
                            pass
                    else:
                        gewerke = []
                        try:
                            for conn in conns:
                                if not conn.MEPSystem:
                                    continue
                                else:
                                    if conn.MEPSystem.Id.ToString() == systemid:
                                        if elem.Id.ToString() not in self.liste_bauteile:
                                            gewerke.append(self.GK)
                                    break
                        except:
                            pass
                        if len(gewerke) > 0:
                            try:
                                for conn in conns:
                                    if not conn.MEPSystem:
                                        if '' not in gewerke:
                                            gewerke.append('')
                                        
                                    else:
                                        systemtyp = conn.MEPSystem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
                                        try:
                                            gewerk = self.dict_external[systemtyp].GK
                                            if gewerk and gewerk != None:
                                                if gewerk not in gewerke:
                                                    gewerke.append(gewerk)
                                        except Exception as e:
                                            logger.error(e)
                            except:
                                pass
                            if len(gewerke) == 1:
                                if elem.Id.ToString() not in self.liste_bauteile:
                                    self.liste_bauteile.append(elem.Id.ToString())
                            else:
                                if elem.Id.ToString() not in self.dict_bauteile_ueber.keys():
                                    self.dict_bauteile_ueber[elem.Id.ToString()] = sorted(gewerke)


    def get_systemtyp(self):
        self.typ = doc.GetElement(self.System_Excel.elementid)
          
    
    def wert_schreiben(self, elem, param_name, wert):
        if wert not in [None,""]:
            para = elem.LookupParameter(param_name)
            
            if para:
                if para.IsReadOnly:
                    return
                if para.StorageType.ToString() == 'Double':
                    try:para.SetValueString(str(wert))
                    except:pass
                       
                elif para.StorageType.ToString() == 'Integer':
                    try:para.Set(int(wert))
                    except:pass
                   
                elif para.StorageType.ToString() == 'String':
                    try:para.Set(str(wert))
                    except:pass
                       
                  

    def wert_schreiben_system(self):
        self.get_systemtyp()
        self.wert_schreiben(self.typ,'IGF_X_Gewerkkürzel',self.GK)
        self.wert_schreiben(self.typ,'IGF_X_Kostengruppe',self.KG)
        self.wert_schreiben(self.typ,'IGF_X_Kennnummer_1',self.KN01)
        self.wert_schreiben(self.typ,'IGF_X_Kennnummer_2',self.KN02)
        self.wert_schreiben(self.typ,'IGF_X_BIM-ID',self.BIMID)
        self.wert_schreiben(self.typ,'IGF_X_SystemName',self.Sysname)
        self.wert_schreiben(self.typ,'IGF_X_SystemKürzel',self.SysKZ)
        self.wert_schreiben(self.typ,'IGF_X_AnlagenGeräteAnzahl',self.AnGeAn)
        self.wert_schreiben(self.typ,'IGF_X_AnlagenNr',self.AnNr)
        self.wert_schreiben(self.typ,'IGF_X_AnlagenProzentualAnteil',self.PzAT)
        self.wert_schreiben(self.typ,'IGF_X_AnlagenProzentualAnzahl',self.PzAZ)
        self.wert_schreiben(self.typ,'IGF_RLT_ZuluftTemperaturSo',self.TempS)
        self.wert_schreiben(self.typ,'IGF_RLT_ZuluftTemperaturWi',self.TempW)
        self.wert_schreiben(self.typ,'IGF_X_Vorgabe_ISO_Dicke',self.ISO_Dicke)
        self.wert_schreiben(self.typ,'IGF_X_Vorgabe_ISO_Art',self.ISO_Art)
        

class Bauteil(object):
    def __init__(self,elemid,system):
        self.elemid = DB.ElementId(int(elemid))
        self.elem = doc.GetElement(self.elemid)
        self.system = system
        self.bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()

    def wert_schreiben(self, elem, param_name, wert):
        if wert not in [None,""]:
            para = elem.LookupParameter(param_name)
            
            if para:
                if para.IsReadOnly:
                    return
                if para.StorageType.ToString() == 'Double':
                    try:
                        para.SetValueString(str(wert))
                    except:
                        logger.error(self.elemid,wert,param_name)
                elif para.StorageType.ToString() == 'Integer':
                    try:
                        para.Set(int(wert))
                    except:
                        logger.error(self.elemid,wert,param_name)
                elif para.StorageType.ToString() == 'String':
                    try:
                        para.Set(str(wert))
                    except:
                        logger.error(self.elemid,wert,param_name)

    
    
    def werte_schreiben_bimid(self):
        self.wert_schreiben(self.elem,'IGF_X_Gewerkkürzel_Exemplar',self.system.GK)
        self.wert_schreiben(self.elem,'IGF_X_KG_Exemplar',self.system.KG)
        self.wert_schreiben(self.elem,'IGF_X_KN01_Exemplar',self.system.KN01)
        self.wert_schreiben(self.elem,'IGF_X_KN02_Exemplar',self.system.KN02)
        self.wert_schreiben(self.elem,'IGF_X_BIM-ID_Exemplar',self.system.BIMID)

        self.wert_schreiben(self.elem,'IGF_X_AnlagenGeräteAnzahl_Exemplar',self.system.AnGeAn)
        self.wert_schreiben(self.elem,'IGF_X_AnlagenNr_Exemplar',self.system.AnNr)
        self.wert_schreiben(self.elem,'IGF_X_SystemKürzel_Exemplar',self.system.SysKZ)
        self.wert_schreiben(self.elem,'IGF_X_SystemName_Exemplar',self.system.Sysname)

        self.wert_schreiben(self.elem,'IGF_RLT_ZuluftTemperaturSo_Exemplar',self.system.TempS)
        self.wert_schreiben(self.elem,'IGF_RLT_ZuluftTemperaturWi_Exemplar',self.system.TempW)

    def werte_schreiben_BB(self):
        try:
            self.wert_schreiben(self.elem,'Bearbeitungsbereich',int(DICT_Workset[self.system.Workset]))
        except Exception as e:
            logger.error(e)

class Bauteil_ueber(object):
    def __init__(self,elemid,system,Gewerk):
        self.elemid = DB.ElementId(int(elemid))
        self.elem = doc.GetElement(self.elemid)
        self.Fam = self.elem.Symbol.FamilyName
        self.typ = self.elem.Name
        self.checked = False
        self.gewerk = Gewerk
        self.gewerkschreiben = ''
        self.systemname = system.systemname
        self.system = system
        self.bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
        self.R = False
        self.B = False
        self.G = False
        self.H = False
        self.K = False
        self.S = False
        self.M = False
        self.set_up()
    
    def set_up(self):
        if 'R' in self.gewerk:
            self.R = True
            self.RReadonly = True
        if 'B' in self.gewerk:
            self.B = True
            self.BReadonly = True
        if 'G' in self.gewerk:
            self.G = True
            self.GReadonly = True
        if 'H' in self.gewerk:
            self.H = True
            self.HReadonly = True
        if 'K' in self.gewerk:
            self.K = True
            self.KReadonly = True
        if 'S' in self.gewerk:
            self.S = True
            self.SReadonly = True
        if 'M' in self.gewerk:
            self.M = True
            self.MReadonly = True

    def get_gewerk(self):
        out_ = ''
        if self.B: out_ += 'B'
        if self.G: out_ += 'G'
        if self.H: out_ += 'H'
        if self.K: out_ += 'K'
        if self.M: out_ += 'M'
        if self.R: out_ += 'R'
        if self.S: out_ += 'S'
        return out_

    def wert_schreiben(self, elem, param_name, wert):
        if wert not in [None,""]:
            para = elem.LookupParameter(param_name)
            
            if para:
                if para.IsReadOnly:
                    return
                if para.StorageType.ToString() == 'Double':
                    try:
                        para.SetValueString(str(wert))
                    except:
                        logger.error(self.elemid,wert,param_name)
                elif para.StorageType.ToString() == 'Integer':
                    try:
                        para.Set(int(wert))
                    except:
                        logger.error(self.elemid,wert,param_name)
                elif para.StorageType.ToString() == 'String':
                    try:
                        para.Set(str(wert))
                    except:
                        logger.error(self.elemid,wert,param_name)

    
    
    def werte_schreiben_bimid(self):
        self.wert_schreiben(self.elem,'IGF_X_Gewerkkürzel_Exemplar',self.gewerkschreiben)
        self.wert_schreiben(self.elem,'IGF_X_KG_Exemplar',self.system.KG)
        self.wert_schreiben(self.elem,'IGF_X_KN01_Exemplar',self.system.KN01)
        self.wert_schreiben(self.elem,'IGF_X_KN02_Exemplar',self.system.KN02)
        self.wert_schreiben(self.elem,'IGF_X_BIM-ID_Exemplar',self.system.BIMID)

        self.wert_schreiben(self.elem,'IGF_X_AnlagenGeräteAnzahl_Exemplar',self.system.AnGeAn)
        self.wert_schreiben(self.elem,'IGF_X_AnlagenNr_Exemplar',self.system.AnNr)
        self.wert_schreiben(self.elem,'IGF_X_SystemKürzel_Exemplar',self.system.SysKZ)
        self.wert_schreiben(self.elem,'IGF_X_SystemName_Exemplar',self.system.Sysname)

        self.wert_schreiben(self.elem,'IGF_RLT_ZuluftTemperaturSo_Exemplar',self.system.TempS)
        self.wert_schreiben(self.elem,'IGF_RLT_ZuluftTemperaturWi_Exemplar',self.system.TempW)

        

    def werte_schreiben_BB(self):
        try:
            self.wert_schreiben(self.elem,'Bearbeitungsbereich',int(DICT_Workset['KG4xx_Übergreifend']))
        except Exception as e:
            logger.error(e)

dict_systeme_alle = {}
dict_systeme = {}

def systemzuordnen():
    luftsysids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType().ToElementIds()
    rohrsysids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType().ToElementIds()
    for system in Liste_Luft:
        dict_systeme_alle[system.Systemname] = MEP_System(system)
        if system.checked or system.bb:
            dict_systeme[system.Systemname] = MEP_System(system)
    for system in Liste_Rohr:
        dict_systeme_alle[system.Systemname] = MEP_System(system)
        if system.checked or system.bb:
            dict_systeme[system.Systemname] = MEP_System(system)

    if len(dict_systeme.keys()) == 0:
        return

    for sysid in luftsysids:
        elem = doc.GetElement(sysid)
        systyp = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        if systyp in dict_systeme.keys():
            system = dict_systeme[systyp]
            system.liste_system.append(elem)

    for sysid in rohrsysids:
        elem = doc.GetElement(sysid)
        systyp = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        if systyp in dict_systeme.keys():
            system = dict_systeme[systyp]
            system.liste_system.append(elem)

systemzuordnen()

liste = dict_systeme.keys()[:]
liste_neu = []

def bauteilzuordnen():
    if len(liste) == 0:
        return False
    with forms.ProgressBar(title="{value}/{max_value} Systeme --- Datenermittlung",cancellable=True, step=1) as pb:
        for n,typ in enumerate(liste):
            if pb.cancelled:
                frage_schicht = abfrage(title= __title__,
                        info = 'Vorgang abrechen oder ermittlete Daten behalten?' , 
                        ja = True,ja_text= 'abbrechen',nein_text='weiter').ShowDialog()
                if frage_schicht.antwort == 'abbrechen':
                    return False
                else:
                    break
            mepsystem = dict_systeme[typ]
            mepsystem.dict_external = dict_systeme_alle
            mepsystem.get_elemente()
            logger.info('{} Elemente in System {}'.format(len(mepsystem.liste_bauteile),mepsystem.systemname))
            liste_neu.append(mepsystem)
            pb.update_progress(n+1,len(liste))
    return True

result_bauteilzuordnen = bauteilzuordnen()

########Übergreifend###########

class Bauteile_ueber_wpf(WPFWindow):
    def __init__(self, xaml_file_name,liste_Bauteile):
        self.liste_Bauteile = liste_Bauteile
        self.tempcoll = ObservableCollection[Bauteil_ueber]()
        WPFWindow.__init__(self, xaml_file_name)

        try:
            self.dataGrid.ItemsSource = self.liste_Bauteile
        except Exception as e:
            logger.error(e)
        self.systemauswahl = self.system_click.IsChecked
        self.famauswahl = self.fam_click.IsChecked
        self.typauswahl = self.typ_click.IsChecked
        self.suche.TextChanged += self.auswahl_txt_changed
        self.system_click.Click += self.auswahl_txt_changed
        self.fam_click.Click += self.auswahl_txt_changed
        self.typ_click.Click += self.auswahl_txt_changed
        self.text_upper = ''

    def auswahl_txt_changed(self,sender,args):
        if self.suche.Text == None:
            text_typ = ''
            self.text_upper = text_typ
            self.systemauswahl = self.system_click.IsChecked
            self.famauswahl = self.fam_click.IsChecked
            self.typauswahl = self.typ_click.IsChecked
            self.dataGrid.ItemsSource = self.liste_Bauteile
            return

        else:
            text_typ = self.suche.Text.upper()
        if text_typ == self.text_upper \
            and self.systemauswahl == self.system_click.IsChecked \
            and self.famauswahl == self.fam_click.IsChecked \
            and self.typauswahl == self.typ_click.IsChecked:return

        self.tempcoll.Clear()
        self.text_upper = text_typ
        self.systemauswahl = self.system_click.IsChecked
        self.famauswahl = self.fam_click.IsChecked
        self.typauswahl = self.typ_click.IsChecked

        if self.systemauswahl:
            for item in self.liste_Bauteile:
                if item.system.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
        elif self.famauswahl:
            for item in self.liste_Bauteile:
                if item.Fam.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
        elif self.typauswahl:
            for item in self.liste_Bauteile:
                if item.typ.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
        
        self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def bchanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:
                    el.B = Checked
            self.dataGrid.Items.Refresh()
    def gchanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:el.G = Checked
            self.dataGrid.Items.Refresh()
    def hchanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:el.H = Checked
            self.dataGrid.Items.Refresh()
    def kchanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:el.K = Checked
            self.dataGrid.Items.Refresh()
    def mchanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:el.M = Checked
            self.dataGrid.Items.Refresh()
    def rchanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:el.R = Checked
            self.dataGrid.Items.Refresh()
    def schanged(self,sender,args):
        Checked = sender.IsChecked
        if sender.DataContext.checked:
            for el in self.liste_Bauteile:
                if el.checked:el.S = Checked
            self.dataGrid.Items.Refresh()

        
        
    def toggle(self,sender,args):
        for el in self.dataGrid.Items:
            el.checked = not el.checked
        self.dataGrid.Items.Refresh()

    def check(self,sender,args):
        for el in self.dataGrid.Items:
            el.checked = True
        self.dataGrid.Items.Refresh()
    def uncheck(self,sender,args):
        for el in self.dataGrid.Items:
            el.checked = False
        self.dataGrid.Items.Refresh()
    def ok(self,sender,args):
        self.dataGrid.Items.Refresh()
        self.Close()

def main(): 
    if not result_bauteilzuordnen:
        return
    if uebergreifend_beruecksichtigt:
        liste_bauteile_ueber = ObservableCollection[Bauteil_ueber]()
        for system in liste_neu:
            dict_bauteile_ueber = system.dict_bauteile_ueber
            for elemid in dict_bauteile_ueber.keys():
                bauteil = Bauteil_ueber(elemid,system,dict_bauteile_ueber[elemid])
                liste_bauteile_ueber.Add(bauteil)
        if liste_bauteile_ueber.Count > 0:
            windowBauteil = Bauteile_ueber_wpf("bauteile.xaml",liste_bauteile_ueber)
            try:
                windowBauteil.ShowDialog()
            except Exception as e:
                logger.error(e)
                windowBauteil.Close()
                return
            
            t = DB.Transaction(doc,'Bauteile übergreifend')
            t.Start()
            if forms.alert('Daten in übergreifenden Bauteilen schreiben?',yes=True,no=True,ok=False):
                with forms.ProgressBar(title='{value}/{max_value} Bauteile',cancellable=True, step=10) as pb2:
                    n = 0
                    for bauteil in liste_bauteile_ueber:
                        if pb2.cancelled:
                            frage_schicht = abfrage(title= __title__,
                                    info = 'Vorgang abrechen oder Daten behalten?' , 
                                    ja = True,ja_text= 'abbrechen',nein_text='weiter').ShowDialog()
                            if frage_schicht.antwort == 'abbrechen':
                                t.RollBack()
                                t.Dispose()
                                return
                            else:
                                break
                        bauteil.gewerkschreiben = bauteil.get_gewerk()
                        if bauteil.system.bimid:
                            bauteil.werte_schreiben_bimid()
                        if bauteil.system.bb:
                            if bauteil.bb in muster_bb and not muster_bb_bearbeiten:
                                pass
                            else:
                                bauteil.werte_schreiben_BB()

                        pb2.update_progress(n+1, liste_bauteile_ueber.Count)
                        n+=1
                        
            t.Commit()
            t.Dispose()

    bearbeitet = []
    nichtbearbeitet = liste_neu[:]
    t = DB.Transaction(doc,'BIM-ID')
    t.Start()
    if forms.alert('Daten schreiben?',yes=True,no=True,ok=False):
        with forms.ProgressBar(title='Systeme --- Datene schreiben',cancellable=True, step=10) as pb2:
            for n,mepsystem in enumerate(liste_neu):
                systeme_title = str(mepsystem.systemname) + ' --- ' + str(n+1) + '/' + str(len(liste_neu)) + 'Systeme'
                pb2.title = '{value}/{max_value} Elemente in System ' +  systeme_title
                pb2.step = int((len(mepsystem.liste_bauteile)) /1000) + 10
                mepsystem.get_systemtyp()
                try:
                    mepsystem.wert_schreiben_system()
                except:
                    pass
                for n1,bauteilid in enumerate(mepsystem.liste_bauteile):
                    if pb2.cancelled:
                        if forms.alert('bisherige Änderung behalten?',yes=True,no=True,ok=False):
                            t.Commit()
                            logger.info('Folgenede Systeme sind bereits bearbeitet.')
                            for el in bearbeitet:
                                logger.info(el.systemname)
                            logger.info('---------------------------------------')
                            logger.info('Folgenede Systeme sind nicht bearbeitet.')
                            for el in nichtbearbeitet:
                                logger.info(el.systemname)
                        else:
                            t.RollBack()
                        
                        return
                    bauteil = Bauteil(bauteilid,mepsystem)
                    if bauteil.system.bimid:
                        bauteil.werte_schreiben_bimid()
                    if bauteil.system.bb:
                        if bauteil.bb in muster_bb and not muster_bb_bearbeiten:
                            pass
                        else:
                            bauteil.werte_schreiben_BB()

                    pb2.update_progress(n1+1, len(mepsystem.liste_bauteile))
                bearbeitet.append(mepsystem)
                nichtbearbeitet.remove(mepsystem)
                
    t.Commit()
    t.Dispose()

main()