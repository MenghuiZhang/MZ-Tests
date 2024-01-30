# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from IGF_Klasse import ItemTemplateMitName

class TreeViewItem(ItemTemplateMitName):
    def __init__(self,name):
        ItemTemplateMitName.__init__(self, name)
        self.children = ObservableCollection[TreeViewItem]()
        self.parent = None
        self._expand = False
    
    @property
    def expand(self):
        return self._expand
    @expand.setter
    def expand(self,value):
        if value != self._expand:
            self._expand = value
            self.RaisePropertyChanged('expand')
    
    
    def checkedchanged(self):
        self.expand = True
        if self.parent:
            checked = None
            for elem in self.parent.children:
                if checked == None:
                    if elem.checked:
                        checked = True
                    elif elem.checked == False:
                        checked = False
                    elif elem.checked == None:
                        checked = None
                        break
                elif checked == True:
                    if not elem.checked:
                        checked = None
                        break
                else:
                    if elem.checked or elem.checked == None:
                        checked = None
                        break
            self.parent.checked = checked
        if self.checked == True:
            for elem in self.children:
                elem.checked = True
        elif self.checked == False:
            for elem in self.children:
                elem.checked = False

class BeschriftungInAnsicht(TreeViewItem):
    def __init__(self,name):
        TreeViewItem.__init__(self,name)
        self.expand = True
        self.ListeHidden = []
        self.ListeShow = [] 
