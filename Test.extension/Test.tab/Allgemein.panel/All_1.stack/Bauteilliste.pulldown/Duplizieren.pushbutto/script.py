# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from System.Windows.Forms import FolderBrowserDialog,DialogResult
import os
__title__ = "Duplizieren"
__doc__ = """Duplizieren"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
exportfolder = script.get_config()

uidoc = revit.uidoc
doc = revit.doc

from pyIGF_logInfo import getlog
getlog(__title__)

bauteilliste_coll = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Schedules) \
    .WhereElementIsNotElementType()
bauteillisteids = bauteilliste_coll.ToElementIds()
bauteilliste_coll.Dispose()

dict_bauteilliste = {}
liste_bauteilliste = []

for elid in bauteillisteids:
    elem = doc.GetElement(elid)
    name = elem.Name
    bb = ''
    try:
        bb = elem.LookupParameter('Bearbeitungsbereich').AsValueString()
    except:
        continue
    if bb.find(name) != -1:
        liste_bauteilliste.append(name)
        dict_bauteilliste[name] = elem



if not bauteillisteids:
    logger.error('Keine Bauteilliste in Projekt')
    script.exit()

class bauteilelistewpf(object):
    def __init__(self):
        self.checked = False
        self.schedulename = ''
        self.elementid = ''

    @property
    def schedulename(self):
        return self._schedulename
    @schedulename.setter
    def schedulename(self, value):
        self._schedulename = value
    @property
    def elementid(self):
        return self._elementid
    @elementid.setter
    def elementid(self, value):
        self._elementid = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value

Liste_bauteilelistewpf = ObservableCollection[bauteilelistewpf]()

for elid in bauteillisteids:
    elem = doc.GetElement(elid)
    name = elem.Name
    elemid = elid.ToString()
    bb = ''
    try:
        bb = elem.LookupParameter('Bearbeitungsbereich').AsValueString()
    except:
        continue
    if bb.find(name) != -1:
        temp = bauteilelistewpf()
        temp.schedulename = name
        temp.elementid = elemid
        if name == viewname:
            temp.checked = True
        Liste_bauteilelistewpf.Add(temp)

global exportOrNot
exportOrNot = False
# GUI Pläne
class ScheduleUI(WPFWindow):
    def __init__(self, xaml_file_name,liste_schedule):
        self.liste_schedule = liste_schedule
        WPFWindow.__init__(self, xaml_file_name)
        self.dataGrid.ItemsSource = liste_schedule
        self.tempcoll = ObservableCollection[bauteilelistewpf]()
        self.altdatagrid = liste_schedule
        self.read_config()

        self.suche.TextChanged += self.search_txt_changed

    def read_config(self):
        try:
            self.exportto.Text = str(exportfolder.folder)
        except:
            self.exportto.Text = exportfolder.folder = ""

    def write_config(self):
        exportfolder.folder = self.exportto.Text.encode('utf-8')
        script.save_config()
    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.schedulename.find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def export(self,sender,args):
        global exportOrNot
        exportOrNot = True
        try:
            if os.path.exists(exportfolder.folder):
                self.Close()
            else:
                UI.TaskDialog.Show('Fehler','Ordner nicht vorhanden')
        except Exception as e:
            logger.error(e)
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

    def durchsuchen(self,sender,args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Ordner auswählen"
        dialog.ShowNewFolderButton = True
        if dialog.ShowDialog() == DialogResult.OK:
            folder = dialog.SelectedPath
            self.exportto.Text = folder
        self.write_config()

ScheduleWPF = ScheduleUI("window.xaml",Liste_bauteilelistewpf)
ScheduleWPF.ShowDialog()

if not exportOrNot:
    script.exit()


def Datenermitteln(elem):
    datelliste = []
    tableData = elem.GetTableData()
    sectionBody = tableData.GetSectionData(SectionType.Body)
    rs = sectionBody.NumberOfRows
    cs = sectionBody.NumberOfColumns
    for r in range(rs):
        rowliste = []
        for c in range(cs):
            cellText = elem.GetCellText(SectionType.Body, r, c)
            rowliste.append(cellText)
        datelliste.append(rowliste)
    return datelliste
def excelexport(daten,path):
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for col in range(len(daten[0])):
        for row in range(len(daten)):
            try:
                liste = daten[row][col].Split(' ')
                if len(liste) <3:
                    float(liste[0])
                    worksheet.write_number(row, col, float(liste[0]))
                else:
                    worksheet.write(row, col, daten[row][col])
            except:
                worksheet.write(row, col, daten[row][col])
    worksheet.autofilter(0, 0, int(len(daten))-1, int(len(daten[0])-1))
    workbook.close()

for el in Liste_bauteilelistewpf:
    if el.checked == True:
        elem = doc.GetElement(ElementId(int(el.elementid)))
        name = elem.Name
        excelpath = str(exportfolder.folder) + '\\' + name + '.xlsx'
        daten_liste = Datenermitteln(elem)
        excelexport(daten_liste,excelpath)
