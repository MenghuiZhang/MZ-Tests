# coding: utf8
from IGF_Forms import WPFWindow
import os
XAML_FILES_DIR = os.path.dirname(__file__)

class SelectFromListe(WPFWindow):
    """
    SelectFromListe.Show(Liste,titil = None, handle_esc = True)
    """
    def __init__(self,Liste,titil = None,handle_esc=True):
        WPFWindow.__init__(self,os.path.join(XAML_FILES_DIR, '_selectfromlist.xaml'),handle_esc=handle_esc)
        if titil:
            self.Title = titil
        self.Liste = Liste
        self.ListView.ItemsSource = self.Liste
    
    def suchetextchanged(self,sender,e):
        if sender.Text in [None,'']:
            for el in self.Liste:el.visibility = True
            return
        for el in self.Liste:
            if el.name.upper().find(sender.Text.upper()) != -1:
                el.visibility = True
            else:
                el.visibility = False
    
    def checkedchanged(self,sender,e):
        item = sender.DataContext
        if item:
            checked = sender.IsChecked
            if item in self.ListView.SelectedItems:
                for el in self.ListView.SelectedItems:
                    el.checked = checked                
    
    def checkall(self,sender,e):
        for el in self.Liste:
            if el.visibility:el.checked = True
    
    def uncheckall(self,sender,e):
        for el in self.Liste:
            if el.visibility:el.checked = False
    
    def toggleall(self,sender,e):
        for el in self.Liste:
            if el.visibility:el.checked = not el.checked
    
    def ok(self,sender,e):
        self.Close()
    
    @staticmethod
    def Show(Liste,titil = None, handle_esc = True):
        tmp = SelectFromListe(Liste,titil,handle_esc)
        tmp.ShowDialog()
        # del tmp
        return [el for el in tmp.Liste if el.checked]
        
        