# coding: utf8
from ansicht import Ansicht,ObservableCollection
from customclass._ansicht import RBLItem,UI
from excel import System
import xlsxwriter
from IGF_log import getlog,getloglocal
from pyrevit import script, forms
from System.Windows.Forms import FolderBrowserDialog,DialogResult

__title__ = "0.10 Export der Bauteilliste "
__doc__ = """
exportiert ausgewählte Bauteilliste.

[2022.06.20]
Version: 2.0
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

exportfolder = script.get_config(doc.ProjectInformation.Name+\
    doc.ProjectInformation.Number+\
        'Bauteilliste export')

Liste_bauteileliste = Ansicht.get_Schedule()
if Liste_bauteileliste.Count == 0:
    UI.TaskDialog.Show('Info','Kein Bauteilliste vorhanden')
    script.exit()

for el in Liste_bauteileliste:
    el.checked = True if el.name == viewname else False

def rgb_to_hex(liste):
    return '''#{:02X}{:02X}{:02X}'''.format(int(liste[0]),int(liste[1]),int(liste[2]))

def get_number_format(data):
    accuracy = data.accuracy
    units = data.units
    if str(accuracy) == '1.0':

        try:
            return '''{} "{}"'''.format(0,units)

        except:
            return None
    try:
        return '''{} "{}"'''.format(str(accuracy).replace('1','0'),units)

    except:
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
        worksheet.set_column(c, c, width=d[0][c].width)
        for r in range(len(d)):
            celldata = d[r][c]
            if celldata.data.GetType().ToString() == 'System.String':
                cellformat = e.add_format()
                # cellformat.set_font_color('#FF0000')
                # cellformat.set_font_color(rgb_to_hex(celldata.textcolor))
                if rgb_to_hex(celldata.background) != '#FFFFFF':
                    cellformat.set_bg_color(rgb_to_hex(celldata.background))
                cellformat.set_align(celldata.textalign.lower())
                worksheet.write(r, c, celldata.data,cellformat)
                
           
            else:
                cellformat = e.add_format()
                # cellformat.set_font_color(rgb_to_hex(celldata.textcolor))
                # print(rgb_to_hex(celldata.textcolor))
                # cellformat.set_font_color(rgb_to_hex(celldata.textcolor))
                if rgb_to_hex(celldata.background) != '#FFFFFF':
                    cellformat.set_bg_color(rgb_to_hex(celldata.background))
                # cellformat.set_border_color('#F000000')
                cellformat.set_align(celldata.textalign.lower())
                number_format = get_number_format(celldata)
                if number_format:
                    cellformat.set_num_format(number_format)

                try:worksheet.write_number(r, c, celldata.data,cellformat)
                except:worksheet.write(r, c, celldata.data,cellformat)
    worksheet.autofilter(0, 0, int(len(d))-1, int(len(d[0])-1))
    e.close()
# GUI Pläne
class ScheduleUI(forms.WPFWindow):
    def __init__(self):
        self.liste_schedule = Liste_bauteileliste
        self.tempcoll = ObservableCollection[RBLItem]()
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
        try:   self.exportto.Text = exportfolder.folder if System.IO.Directory.Exists(exportfolder.folder) else ''
        except:self.exportto.Text = exportfolder.folder = ""

    def write_config(self):
        exportfolder.folder = self.exportto.Text
        script.save_config()
    
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
                    data = el.get_Data2()
                    write_one_sheet(data,str(exportfolder.folder) \
                        + '\\' + el.name + '.xlsx')

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
ScheduleWPF.ShowDialog()
# try: ScheduleWPF.ShowDialog()
# except Exception as e:
#     script.get_logger().error(e)
#     ScheduleWPF.Close()
#     script.exit()


liste = []
for el in cl:
    _id = el.LookupParameter('IGF_X_Bauteil_ID_Text').AsString()
    if _id not in liste:
        liste.append(_id)
    else:
        print(_id)
print(liste)

liste = ['HK_E7_20', 'HK_E8_02', 'HK_E8_03', 'HK_E6_02', 'HK_E6_03', 'HK_E5_24', 'HK_E5_25', 'HK_E4_02', 'HK_E4_03', 'HK_E3_15', 'HK_E3_16', 'HK_E2_02', 'HK_E2_03', 'HK_E7_16', 'HK_E7_17', 'HK_E03_103', 'HK_E03_104', 'HK_E03_102', 'HK_E03_105', 'HK_E03_101', 'HK_E04_54', 'HK_E04_56', 'HK_E04_55', 'HK_E04_58', 'HK_E04_53', 'HK_E05_20', 'HK_E05_22', 'HK_E05_21', 'HK_E05_19', 'HK_E05_18', 'HK_E02_48', 'HK_E02_50', 'HK_E02_49', 'HK_E02_47', 'HK_E02_46', 'HK_E1_18', 'HK_E1_20', 'HK_E1_19', 'HK_E1_17', 'HK_E1_16', 'HK_E2_11', 'HK_E2_13', 'HK_E2_12', 'HK_E2_10', 'HK_E2_09', 'HK_E3_24', 'HK_E3_26', 'HK_E3_25', 'HK_E3_29', 'HK_E3_23', 'HK_E4_12', 'HK_E4_14', 'HK_E4_13', 'HK_E4_11', 'HK_E4_10', 'HK_E5_20', 'HK_E5_22', 'HK_E5_21', 'HK_E5_19', 'HK_E5_18', 'HK_E6_12', 'HK_E6_14', 'HK_E6_13', 'HK_E6_11', 'HK_E6_10', 'HK_E7_26', 'HK_E7_27', 'HK_E7_25', 'HK_E7_24', 'HK_E8_12', 'HK_E8_14', 'HK_E8_13', 'HK_E8_11', 'HK_E8_10', 'HK_E8_01', 'HK_E8_05', 'HK_E7_15', 'HK_E7_19', 'HK_E6_01', 'HK_E6_05', 'HK_E5_23', 'HK_E5_27', 'HK_E4_01', 'HK_E4_04', 'HK_E3_14', 'HK_E2_01', 'HK_E2_04', 'HK_E04_59', 'HK_E04_60', 'HK_E04_61', 'HK_E04_51', 'HK_E04_52', 'HK_E05_11', 'HK_E05_13', 'HK_E05_14', 'HK_E05_23', 'HK_E05_24', 'HK_E06_03', 'HK_E06_04', 'HK_E06_01', 'HK_E05_06', 'HK_E05_07', 'HK_E05_27', 'HK_E05_26', 'HK_E05_28', 'HK_E05_29', 'HK_E05_25', 'HK_E05_16', 'HK_E05_17', 'HK_E03_99', 'HK_E03_100', 'HK_E02_43', 'HK_E02_44', 'HK_E1_13', 'HK_E1_14', 'HK_E2_06', 'HK_E2_07', 'HK_E3_19', 'HK_E3_20', 'HK_E4_07', 'HK_E4_08', 'HK_E5_15', 'HK_E5_16', 'HK_E6_07', 'HK_E6_08', 'HK_E7_22', 'HK_E7_23', 'HK_E8_07', 'HK_E8_08', 'HK_E05_15', 'HK_E04_50', 'HK_E03_98', 'HK_E02_42', 'HK_E1_12', 'HK_E2_05', 'HK_E3_18', 'HK_E4_06', 'HK_E5_14', 'HK_E6_06', 'HK_E7_21', 'HK_E8_06', 'HK_E9_01', 'HK_E05_30', 'HK_E05_31', 'HK_E05_33', 'HK_E05_34', 'HK_E05_32', 'HK_E03_95', 'HK_E03_96', 'HK_E03_94', 'HK_E03_97', 'HK_E03_91', 'HK_E03_90', 'HK_E03_89', 'HK_E03_88', 'HK_E03_87', 'HK_E03_86', 'HK_E03_92', 'HK_E02_52', 'HK_E02_41', 'HK_E02_45', 'HK_E02_40', 'HK_E1_15', 'HK_E1_11', 'HK_E3_17', 'HK_E8_04', 'HK_E7_18', 'HK_E6_04', 'HK_E5_26', 'HK_E4_05', 'HK_E3_27', 'HK_E2_27', 'HK_E6_09', 'HK_E5_17', 'HK_E4_09', 'HK_E3_21', 'HK_E2_08', 'HK_E02_51', 'HK_E03_93', 'HK_E04_09', 'HK_E04_10', 'HK_E04_07', 'HK_E04_08', 'HK_E04_05', 'HK_E04_06', 'HK_E04_03', 'HK_E04_04', 'HK_E04_01', 'HK_E04_02', 'HK_E05_10', 'HK_E05_12', 'HK_E05_40', 'HK_E05_41', 'HK_E05_42', 'HK_E7_28', 'HK_E8_29', 'HK_E05_39', 'HK_E06_02', 'HK_E05_36', 'HK_E05_38', 'HK_E05_37', 'HK_E05_43', 'HK_E05_44']
for el in cl:
    _id = el.LookupParameter('IGF_X_Bauteil_ID_Text').AsString()
    if _id in liste:
        liste.remove(_id)
    else:
        print(_id)
print(liste)

HK_E7_19
HK_E03_98
['HK_E7_17', 'HK_E03_99']