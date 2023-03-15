# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, UI, DB
from pyrevit import script, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Forms import FolderBrowserDialog,DialogResult
import xlsxwriter
import os



__title__ = "Exportieren"
__doc__ = """exportiert ausgewählte Bauteileliste.


[2021.10.04]
Version: 1.1
"""
__author__ = "Menghui Zhang"


try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()
exportfolder = script.get_config()

uidoc = revit.uidoc
doc = revit.doc
view = doc.ActiveView
viewname = None

if view.ToString() == 'Autodesk.Revit.DB.ViewSchedule':
    viewname = view.Name

bauteilliste_coll = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Schedules).WhereElementIsNotElementType()
bauteillisteids = bauteilliste_coll.ToElementIds()
bauteilliste_coll.Dispose()

if not bauteillisteids:
    logger.error('Keine Bauteilliste in Projekt')
    script.exit()

class Bauteilelist(object):
    def __init__(self,elementid):
        self.checked = False
        self.elementid = elementid.ToString()
        self.element = doc.GetElement(elementid)

    @property
    def schedulename(self):
        return self.element.Name

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value

Liste_bauteileliste = ObservableCollection[Bauteilelist]()

for elid in bauteillisteids:
    elem = doc.GetElement(elid)
    name = elem.Name
    elemid = elid.ToString()

    if elem.OwnerViewId.ToString() != '-1':
        continue
    if elem.IsTemplate:
        continue
    
    temp = Bauteilelist(elid)

    if name == viewname:
        temp.checked = True
    Liste_bauteileliste.Add(temp)
global exportOrNot
exportOrNot = False


# GUI Pläne
class ScheduleUI(forms.WPFWindow):
    def __init__(self, xaml_file_name,liste_schedule):
        self.liste_schedule = liste_schedule
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.listview_bauteileliste.ItemsSource = liste_schedule
        self.tempcoll = ObservableCollection[Bauteilelist]()
        self.altdatagrid = liste_schedule
        self.read_config()

        self.suche.TextChanged += self.search_txt_changed

    def read_config(self):
        try:
            self.exportto.Text = exportfolder.folder
        except:
            self.exportto.Text = exportfolder.folder = ""

    def write_config(self):
        exportfolder.folder = self.exportto.Text
        script.save_config()
    
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.listview_bauteileliste.SelectedItem is not None:
            try:
                if sender.DataContext in self.listview_bauteileliste.SelectedItems:
                    for item in self.listview_bauteileliste.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.listview_bauteileliste.Items.Refresh()
                else:
                    pass
            except:
                pass

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.listview_bauteileliste.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.schedulename.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.listview_bauteileliste.ItemsSource = self.tempcoll
        self.listview_bauteileliste.Items.Refresh()

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
        for item in self.listview_bauteileliste.Items:
            item.checked = True
        self.listview_bauteileliste.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.listview_bauteileliste.Items:
            item.checked = False
        self.listview_bauteileliste.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.listview_bauteileliste.Items:
            value = item.checked
            item.checked = not value
        self.listview_bauteileliste.Items.Refresh()

    def durchsuchen(self,sender,args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Ordner auswählen"
        dialog.ShowNewFolderButton = True
        if dialog.ShowDialog() == DialogResult.OK:
            folder = dialog.SelectedPath
            self.exportto.Text = folder
        self.write_config()

ScheduleWPF = ScheduleUI("window.xaml",Liste_bauteileliste)
try:
    ScheduleWPF.ShowDialog()
except Exception as e:
    logger.error(e)
    ScheduleWPF.Close()

if not exportOrNot:
    script.exit()

def Datenermitteln(elem):
    datelliste = []
    paramname = {}
    paramnameid = {}
    params = elem.Definition.GetFieldOrder()
    for param in params:
        typ = elem.Definition.GetField(param).UnitType.ToString()
        name = elem.Definition.GetField(param).ColumnHeading
        if typ == 'UT_Undefined':
            paramname[name] = 'Text'
        elif typ == 'UT_Number':
            paramname[name] = 'Zahl'
        else:
            paramname[name] = 'Double'
    
    tableData = elem.GetTableData()
    sectionBody = tableData.GetSectionData(DB.SectionType.Body)
    rs = sectionBody.NumberOfRows
    cs = sectionBody.NumberOfColumns
    rowliste_temp = []
    for c in range(cs):
        cellText = elem.GetCellText(DB.SectionType.Body, 0, c)
        paramnameid[c] = cellText
        rowliste_temp.append(cellText)
    datelliste.append(rowliste_temp)

    for r in range(1,rs):
        rowliste = []
        
        for c in range(cs):
            
            typ = paramname[paramnameid[c]]
            if typ == 'Text':
                cellText = elem.GetCellText(DB.SectionType.Body, r, c)
            elif typ == 'Zahl':
                cellText = elem.GetCellText(DB.SectionType.Body, r, c)

                if cellText.find('%') == -1:
                    try:
                        cellText = float(cellText)
                    except:
                        pass
                

            else:
                cellText = elem.GetCellText(DB.SectionType.Body, r, c)
                try:
                    cellText = float(cellText)
                except:
                    try:
                        cellText = float(cellText[:cellText.find(' ')])
                    except:
                        pass

            rowliste.append(cellText)
        datelliste.append(rowliste)
    return datelliste,paramname,paramnameid

def excelexport(daten,path,paramname,paramnameid):
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    for col in range(len(daten[0])):
        type1 = paramname[paramnameid[col]]
        for row in range(len(daten)):
            wert = daten[row][col]
            typ = wert.GetType().ToString()
            if typ == 'System.String':
                if type1 == 'Zahl' and wert.find('%') != -1:
                    try:
                        zahl = wert[:wert.find('%')]
                        kommestelle = zahl[zahl.find('.')+1:]
                        if len(kommestelle) > 0:
                            form = '0.'+len(kommestelle)*'0'+'%'
                            cellform = workbook.add_format({'num_format': '{}'.format(form)})
                            try:
                                werte = float(wert[:wert.find('%')]) / 100
                                worksheet.write_number(row, col, werte,cellform)
                                
                            except Exception as e:
                                print(e)
                                worksheet.write(row, col, wert)
                    except:
                        worksheet.write(row, col, wert)
                else:
                    worksheet.write(row, col, wert)
            else:
                try:
                    worksheet.write_number(row, col, wert)
                except:
                    worksheet.write(row, col, wert)

    worksheet.autofilter(0, 0, int(len(daten))-1, int(len(daten[0])-1))
    workbook.close()

for el in Liste_bauteileliste:
    if el.checked == True:
        elem = doc.GetElement(DB.ElementId(int(el.elementid)))
        name = elem.Name
        excelpath = str(exportfolder.folder) + '\\' + name + '.xlsx'
        daten_liste,namedict,iddict = Datenermitteln(elem)
        excelexport(daten_liste,excelpath,namedict,iddict)