# coding: utf8
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import time
from pyrevit import script, forms
from rpw import *
import System


start = time.time()

__title__ = "3.02 Raumnutzung und Lüftungswechsel"
__doc__ = """Raumnutzung und Lüftungswechsel"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
my_config = script.get_config()
from pyIGF_logInfo import getlog
getlog(__title__)

uidoc = revit.uidoc
doc = revit.doc
app = revit.app

exapp = Excel.ApplicationClass()

spaces_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()

logger.info("{} MEP Räume ausgewählt".format(len(spaces)))
excelPath_default = "R:\\Vorlagen\\_IGF\\Revit_Parameter\\Raumnutzung und Lüftwechsel.xlsx"


def ExcelLesen(inExcelPath):
    outDaten = {}
    book = exapp.Workbooks.Open(inExcelPath)
    for sheet in book.Worksheets:
        rows = sheet.UsedRange.Rows.Count
        for row in range(2, rows + 1):
            zelleDaten = []
            name = sheet.UsedRange.Cells[row, 1].Value2
            for col in [2,3,4]:
                werte = sheet.UsedRange.Cells[row, col].Value2
                zelleDaten.append(werte)
            outDaten[name] = zelleDaten
    book.Save()
    book.Close()
    return outDaten

try:
    excelPath_default = my_config.excelPath_.decode('utf-8')
except:
    pass

excelPath = ui.forms.TextInput('Excel: ', default = excelPath_default)
excelDaten = ExcelLesen(excelPath)
for i in excelDaten:
    print(i,excelDaten[i])
my_config.excelPath_ = excelPath.encode('utf-8')
script.save_config()


class MEPRaum:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        attr = [
            ["name", "Name"],
            ["nummer", "Nummer"],
            ["kez", "TGA_RLT_VolumenstromProNummer"],
            ["vol_faktor", "TGA_RLT_VolumenstromProFaktor"],
        ]
        for a in attr:
            python_name, revit_name = a
            setattr(self, python_name, self.__get_element_attr(revit_name))

        logger.info(30 * "=")
        logger.info("{}\t{}".format(self.nummer, self.name))
        self.namelist1 = self.name.split(" ")
        self.WC_TH = self.name[:2]
        self.namelist2 = self.name.split("/")
        self.namelist = self.namelist1 + self.namelist2
        self.namelist.append(self.WC_TH)
        self.newKez,self.newFaktor = self.werteAusexcel(self.namelist)


    def __get_element_attr(self, param_name):
        param = self.element.LookupParameter(param_name)

        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return

        return self.__get_value_in_project_units(param)

    def __get_value_in_project_units(self, param):
        """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""

        value = revit.query.get_param_value(param)

        try:
            unit = param.DisplayUnitType

            value = DB.UnitUtils.ConvertFromInternalUnits(
                value,
                unit)

        except Exception as e:
            pass

        return value

    def werteAusexcel(self,Liste):
        newKez = None
        newFaktor = None
        for i in Liste:
            if i in excelDaten.keys():
                daten = excelDaten[i]
                newKez = daten[1]
                newFaktor = daten[2]
                break
        return newKez,newFaktor


    def table_row(self):
        """ Gibt eine Datenreihe für den MEP Raum aus. Für die tabellarische Übersicht."""
        return [
            self.nummer,
            self.name,
            self.kez,
            self.vol_faktor,
            self.newKez,
            self.newFaktor,
        ]

    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                logger.info(
                    "{} - {} Werte schreiben ({})".format(self.nummer, param_name, wert))
                if self.element.LookupParameter(param_name):
                    self.element.LookupParameter(param_name).SetValueString(str(wert))
        if not self.kez:
            wert_schreiben("TGA_RLT_VolumenstromProNummer", self.newKez)
            wert_schreiben("TGA_RLT_VolumenstromProFaktor", self.newFaktor)

table_data = []
mepraum_liste = []
with forms.ProgressBar(title="{value}/{max_value} Luftmengenberechnung",
                       cancellable=True, step=10) as pb:

    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)
        table_data.append(mepraum.table_row())

table_data.sort()

output.print_table(
    table_data=table_data,
    title="Luftmengen",
    columns=["Nummer", "Name", "TGA_RLT_VolumenstromProNummer",
             "TGA_RLT_VolumenstromProFaktor", "Kez aus Excel",
             "Vol_Faktor aus Excel"
             ]
)

if forms.alert("Werte in Modell schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Werte schreiben",
                           cancellable=True, step=10) as pb2:

        n_1 = 0

        with rpw.db.Transaction("Raumnutzung und Lüftungswechsel"):
            for mepraum in mepraum_liste:
                if pb2.cancelled:
                    script.exit()
                n_1 += 1
                pb2.update_progress(n_1, len(spaces))

                mepraum.werte_schreiben()


total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))
