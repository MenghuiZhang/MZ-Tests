# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
import time
from IGF_lib import get_value
import xlsxwriter

__title__ = "Fläche export"
__doc__ = """
[2022.01.24]
Version: 1.3
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
start = time.time()
uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass

luftleitungen = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElementIds()
luftformteile = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElementIds()
Export_Liste = {}
class LuftLeitung:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.Flaeche = get_value(self.elem.LookupParameter('Fläche'))
        ##self.Matetial = doc.GetElement(self.elem.GetTypeId()).LookupParameter('CAx Materialkz').AsString()
        self.Matetial = self.elem.LookupParameter('SBI_Materialien').AsString()
        self.Art = self.elem.LookupParameter('Systemklassifizierung').AsString()

class Luftkanalformteile:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        try:
            self.Flaeche = get_value(self.elem.LookupParameter('IGF_RLT_Fläche'))
        except:
            self.Flaeche = 0
        try:
            self.Matetial = self.elem.LookupParameter('SBI_Materialien').AsString()
        except:
            self.Matetial = None
        try:
            self.Art = self.elem.LookupParameter('Systemklassifizierung').AsString()
        except:
            self.Art = None

        
        

with forms.ProgressBar(title='{value}/{max_value} Luftkanäle',cancellable=True, step=int(len(luftleitungen)/200)+1) as pb:
    for n,elemid in enumerate(luftleitungen):
        if pb.cancelled:
            script.exit()
        
        pb.update_progress(n+1,len(luftleitungen))

        luftkanal = LuftLeitung(elemid)
        if luftkanal.Art == 'Nicht definiert':
            luftkanal.Art = 'Zuluft'
        if luftkanal.Art not in Export_Liste.keys():
            Export_Liste[luftkanal.Art] = {}
        if luftkanal.Matetial not in Export_Liste[luftkanal.Art].keys():
            Export_Liste[luftkanal.Art][luftkanal.Matetial] = luftkanal.Flaeche
        else:
            Export_Liste[luftkanal.Art][luftkanal.Matetial] += luftkanal.Flaeche
        
with forms.ProgressBar(title='{value}/{max_value} Luftkanalformteile',cancellable=True, step=int(len(luftformteile)/200)+1) as pb:
    for n,elemid in enumerate(luftformteile):
        if pb.cancelled:
            script.exit()
        
        pb.update_progress(n+1,len(luftformteile))

        luftkanalformteil = Luftkanalformteile(elemid)
        if not luftkanalformteil.Flaeche:
            continue
        if luftkanalformteil.Art not in Export_Liste.keys():
            Export_Liste[luftkanalformteil.Art] = {}
        if luftkanalformteil.Matetial not in Export_Liste[luftkanalformteil.Art].keys():
            Export_Liste[luftkanalformteil.Art][luftkanalformteil.Matetial] = luftkanalformteil.Flaeche
        else:
            Export_Liste[luftkanalformteil.Art][luftkanalformteil.Matetial] += luftkanalformteil.Flaeche






	
	
	
	
	
		
		
			
			
            



excelpath = r'C:\Users\Zhang\Desktop\BBIM_Luftleitung_Fläche.xlsx'
workbook = xlsxwriter.Workbook(excelpath)
worksheet = workbook.add_worksheet()
worksheet.write(0,0,'Art')
worksheet.write(0,1,'Material')
worksheet.write(0,2,'Fläche')
n = 0
for art in Export_Liste.keys():
    for Mate in Export_Liste[art].keys():
        n+=1
        worksheet.write(n,0,art)
        worksheet.write(n,1,Mate)
        worksheet.write_number(n,2,round(Export_Liste[art][Mate],1))   
worksheet.autofilter(0, 0, n-1, 2)
workbook.close()



