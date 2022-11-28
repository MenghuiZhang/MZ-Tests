# coding: utf8
from eventhandler import ExternalEvent,config,script,forms,Liste_Kategorie,FamilieUI,SELECT,SCHREIBEN,SELECTMEP,SCHREIBENMEP,SCHREIBENBSK

__title__ = "Verknüpfung durch Zeichnen (GUI)"
__doc__ = """der markeierten Fläche wird der definierte Slave von Value zugewiesen"""
__authors__ = "Menghui Zhang"
       
class GUI(forms.WPFWindow):
    def __init__(self):
        self.elems = []
        self.Liste_Kategorie = Liste_Kategorie
        self.config = config
        self.script = script
        self.FamilieUI = FamilieUI
        self.select = SELECT()
        self.schreiben = SCHREIBEN()
        self.selectmep = SELECTMEP()
        self.schreibenmep = SCHREIBENMEP()
        self.schreibenbsk = SCHREIBENBSK()

        self.selectmepevent = ExternalEvent.Create(self.selectmep)
        self.schreibenmepevent = ExternalEvent.Create(self.schreibenmep )
        self.schreibenbskevent = ExternalEvent.Create(self.schreibenbsk)
        self.schreibenevent = ExternalEvent.Create(self.schreiben)
        self.selectevent = ExternalEvent.Create(self.select)

        forms.WPFWindow.__init__(self, "GUI.xaml")
        self.read_config()

     
    def familieauswahl(self,sender,args):
        wpfui = self.FamilieUI(self.Liste_Kategorie)
        try:
            wpfui.ShowDialog()
            self.write_config()
        except Exception as e:
            print(e)
            wpfui.Close()
    

        
    def read_config(self):
        try:
            Familien = self.config.Liste_Familien
            for class_cate in self.Liste_Kategorie:
                for familie in class_cate.Familien:
                    if familie.Name in Familien:
                        familie.checked = True
        except:
            pass
            


    def write_config(self):
        try:
            Familien = []
            for class_cate in self.Liste_Kategorie:
                for familie in class_cate.Familien:
                    if familie.checked == True:
                        Familien.append(familie.Name)
            self.config.Liste_Familien = Familien

            self.script.save_config()
        except:pass

    def auswahl(self,sender,args):
        self.selectevent.Raise()
    def writevalue(self,sender,args):
        self.schreibenevent.Raise()
    
    def auswahlmep(self,sender,args):
        self.selectmepevent.Raise()
    def writevaluemep(self,sender,args):
        self.schreibenmepevent.Raise()
    
    def bsk(self,sender,args):
        self.schreibenbskevent.Raise()

gui = GUI()
gui.select.GUI = gui
gui.schreiben.GUI = gui
gui.selectmep.GUI = gui
gui.schreibenmep.GUI = gui
gui.schreibenbsk.GUI = gui
gui.Show()