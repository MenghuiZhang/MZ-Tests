# coding: utf8
# from ansicht import Ansicht,ObservableCollection
from excel import System
import xlsxwriter
from IGF_log import getlog,getloglocal
from pyrevit import script, forms, UI
from System.Windows.Forms import FolderBrowserDialog,DialogResult
import math
from config import ProjektConfig
from IGF_Klasse._ansicht import ObservableCollection,Ansicht,Bauteilliste
import time


__title__ = "0.10 Export der Bauteilliste "
__doc__ = """
exportiert ausgewählte Bauteilliste.

[2022.12.06]
Version: 2.1
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

uidoc = Ansicht.uidoc
doc = Ansicht.doc
viewname = uidoc.ActiveView.Name if uidoc.ActiveView.ToString() == 'Autodesk.Revit.DB.ViewSchedule' else ''
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number

date = time.strftime(r'%d.%m',time.localtime())

exportfolder = ProjektConfig.get_config_static(number + ' - ' + name,'Bauteilliste export',True)

Liste_bauteileliste = Ansicht.get_Schedule()
if Liste_bauteileliste.Count == 0:
    UI.TaskDialog.Show('Info','Kein Bauteilliste vorhanden')
    script.exit()

for el in Liste_bauteileliste:
    if el.name == viewname:el.checked = True 

def rgb_to_hex(liste):
    return '''#{:02X}{:02X}{:02X}'''.format(int(liste[0]),int(liste[1]),int(liste[2]))

def get_number_format(data):
    accuracy = data.accuracy
    units = data.units
    if accuracy:
        komma = int(math.log10(accuracy))
        if komma >= 0:
            try:
                return '''{} "{}"'''.format('0'*(komma+1),units)

            except Exception as e:
                print(e)
                return None
        else:
            try:
                return '''{} "{}"'''.format('0.'+'0'*(0-komma),units)

            except Exception as e:
                print(e)
                return None


def write_one_sheet(d,p,sheetname = None):
    e = xlsxwriter.Workbook(p)
    if sheetname:worksheet = e.add_worksheet(sheetname)
    else:worksheet = e.add_worksheet()
    if len(d) == 0:
        return
    else:
        if len(d[0]) == 0:
            return
    for c in range(len(d[0])):
        try:worksheet.set_column(c, c, width=d[0][c].width*0.8)
        except:pass
        for r in range(len(d)):
     
            celldata = d[r][c]
            if celldata.data.GetType().ToString() == 'System.String':
                cellformat = e.add_format()
                cellformat.set_font_color(rgb_to_hex(celldata.textcolor))
                if rgb_to_hex(celldata.background) != '#FFFFFF':
                    cellformat.set_bg_color(rgb_to_hex(celldata.background))
                cellformat.set_align(celldata.textalign.lower())
                if r == 0:
                    cellformat.set_bold()
                worksheet.write(r, c, celldata.data,cellformat)
                
            else:
                cellformat = e.add_format()
                cellformat.set_font_color(rgb_to_hex(celldata.textcolor))
                if rgb_to_hex(celldata.background) != '#FFFFFF':
                    cellformat.set_bg_color(rgb_to_hex(celldata.background))
                cellformat.set_align(celldata.textalign.lower())
                number_format = get_number_format(celldata)
                if number_format:
                    cellformat.set_num_format(number_format)

                try:worksheet.write_number(r, c, celldata.data,cellformat)
                except:worksheet.write(r, c, celldata.data,cellformat)
    worksheet.autofilter(0, 0, int(len(d))-1, int(len(d[0])-1))
    worksheet.freeze_panes(1, 0)
    e.close()
# GUI Pläne
class ScheduleUI(forms.WPFWindow):
    def __init__(self):
        self.liste_schedule = Liste_bauteileliste
        self.tempcoll = ObservableCollection[Bauteilliste]()
        self.altdatagrid = Liste_bauteileliste
        forms.WPFWindow.__init__(self, "window.xaml")

        self.LB_Schedule.ItemsSource = Liste_bauteileliste
        self.read_config()
        self.suche.TextChanged += self.search_txt_changed
    
    def movewindow(self, sender, args):
        self.DragMove()

    def close(self, sender, args):
        self.Close()

    def read_config(self):
        try:   self.exportto.Text = exportfolder.get_value('folder')  if System.IO.Directory.Exists(exportfolder.get_value('folder')) else ''
        except:self.exportto.Text = ""

    def write_config(self):
        try:
            exportfolder.set_value('folder', self.exportto.Text)
            exportfolder.save_config()
        except:pass
    
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LB_Schedule.SelectedItem is not None:
            try:
                if sender.DataContext in self.LB_Schedule.SelectedItems:
                    for item in self.LB_Schedule.SelectedItems:
                        try:     item.checked = Checked
                        except:  pass
                    self.LB_Schedule.Items.Refresh()
                else:  pass
            except:  pass

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            text_typ = ''
            self.LB_Schedule.ItemsSource = self.altdatagrid

        else:
            for item in self.altdatagrid:
                if item.name.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.LB_Schedule.ItemsSource = self.tempcoll
        self.LB_Schedule.Items.Refresh()

    def export(self,sender,args):
        if not self.exportto.Text:
            UI.TaskDialog.Show('Info.','Kein Ordner ausgewählt')
            return
        elif System.IO.Directory.Exists(self.exportto.Text) is False:
            UI.TaskDialog.Show('Info.','Ordner nicht vorhanden')
            self.exportto.Text = ''
            return
        else:
            self.write_config()
            self.Close()
            Liste_neu = [el for el in Liste_bauteileliste if el.checked == True]
            if len(Liste_neu) == 0:return
            with forms.ProgressBar(title="{value}/{max_value} Bauteilliste", step=1) as pb:
                for n, el in enumerate(Liste_neu):
                    pb.update_progress(n + 1, len(Liste_neu))
                    data = el.get_Data_MitFormat()
                    write_one_sheet(data,str(exportfolder.get_value('folder')) \
                        + '\\' + el.name + '_'+ date + '.xlsx')

    def checkall(self,sender,args):
        for item in self.LB_Schedule.Items:
            item.checked = True
        self.LB_Schedule.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.LB_Schedule.Items:
            item.checked = False
        self.LB_Schedule.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.LB_Schedule.Items:
            value = item.checked
            item.checked = not value
        self.LB_Schedule.Items.Refresh()

    def durchsuchen(self,sender,args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Ordner auswählen"
        dialog.ShowNewFolderButton = True
        if dialog.ShowDialog() == DialogResult.OK:
            folder = dialog.SelectedPath
            self.exportto.Text = folder
        self.write_config()

ScheduleWPF = ScheduleUI()
try: ScheduleWPF.ShowDialog()
except Exception as e:
    script.get_logger().error(e)
    ScheduleWPF.Close()
    script.exit()