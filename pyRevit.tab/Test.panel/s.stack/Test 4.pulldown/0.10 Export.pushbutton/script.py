# coding: utf8
from ansicht import Ansicht,ObservableCollection
from customclass._ansicht import RBLItem,UI
from excel import WExcel,System
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
                    WExcel.write_one_sheet(data,str(exportfolder.folder) \
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
try: ScheduleWPF.ShowDialog()
except Exception as e:
    script.get_logger().error(e)
    ScheduleWPF.Close()
    script.exit()