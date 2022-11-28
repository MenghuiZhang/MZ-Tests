# coding: utf8
from IGF_log import getlog,getloglocal
from pyrevit import forms,script
from rpw import DB,revit,UI
import os
import clr
from System.Windows.Input import Key


__title__ = "4. RV-Nummer aus HLS-Bauteilnummer übertragen"
__doc__ = """

Bsp. Heizkörper: HK_100_101
     in THV : THV_100_101 
[2022.08.03]
Version: 1.0

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

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc
active_view = uidoc.ActiveView
if active_view.ViewType.ToString() != 'Schedule':
    UI.TaskDialog.Show('Fehler','Bitte ein Bauteilliste öffnen!')
    script.exit()

HLSBauteil = DB.FilteredElementCollector(doc,active_view.Id).ToElements()
RVs = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
Families = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family)).ToElements()

Param_HLSBauteil = []
Param_RV = []
RV_Families = []

for el in HLSBauteil:
    params = el.Parameters
    for para in params:
        defi = para.Definition.Name
        Param_HLSBauteil.append(defi)
    break
for el in RVs:
    params = el.Parameters
    for para in params:
        defi = para.Definition.Name
        Param_RV.append(defi)
    break

for el in Families:RV_Families.append(el.Name)

class GUI(forms.WPFWindow):
    def __init__(self):
        self.Key = Key
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
        self.hlsbauteil.ItemsSource = Param_HLSBauteil
        self.regelventil.ItemsSource = Param_RV
        self.rvfamilie.ItemsSource = RV_Families

    def aktualisieren(self, sender, args):
        if self.hlsbauteil.SelectedIndex != -1 and self.regelventil.SelectedIndex != -1 and self.rvfamilie.SelectedIndex != -1 and \
            self.von.Text != '' and self.von.Text != None:
            self.Close() 
        else:
            UI.TaskDialog.Show('Fehler','Einstellungsfelder nicht vollstandig eingegeben')

    def Setkey(self, sender, args):   
        if ((args.Key >= self.Key.D0 and args.Key <= self.Key.D9) or (args.Key >= self.Key.NumPad0 and args.Key <= self.Key.NumPad9) \
            or args.Key == self.Key.Delete or args.Key == self.Key.Back):
            args.Handled = False
        
        else:
            args.Handled = True
       
wind = GUI()
try:wind.ShowDialog()
except:
    wind.Close()
    script.exit()
if wind.Praefix.Text == None:wind.Praefix.Text = ''
if wind.Suffix.Text == None:wind.Suffix.Text = ''
if wind.biszum.Text == None:wind.biszum.Text = ''

class ME:
    def __init__(self,elem):
        self.elem = elem
        self.vsr = ''
        self.liste = []
        self.get_vsr(self.elem)
        self.elemID = self.elem.LookupParameter(wind.hlsbauteil.SelectedItem.ToString()).AsString()
        if wind.biszum.Text == '':wind.biszum.Text = len(self.elemID) - 1
        self.vsrID = wind.Praefix.Text+self.elemID[int(wind.von.Text):int(wind.biszum.Text)]+wind.Suffix.Text

    def get_vsr(self,elem):
        if self.vsr:return
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:

                if conns.Size > 3 and cate != 'HLS-Bauteile':return
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste:
                            if owner.Category.Name == 'Rohrzubehör':
                                faminame = owner.Symbol.FamilyName
                                if faminame == wind.rvfamilie.SelectedItem.ToString():
                                    self.vsr = owner
                                    return
                                                                        
                            self.get_vsr(owner)
    def wert_schreiben(self):
        self.vsr.LookupParameter(wind.regelventil.SelectedItem.ToString()).Set(self.vsrID)
        
Liste = []
for el in HLSBauteil:
    mechanicalequip = ME(el)
    Liste.append(mechanicalequip(el))

t = DB.Transaction(doc,'VSR Id')
t.Start()
for el in Liste:
    if not el.vsr:
        logger.error("kein RV für Element {} gefunden".format(el.elem.Id.ToString()))
        continue
    try:
        el.wert_schreiben()
    except Exception as e:logger.error(e)
t.Commit()