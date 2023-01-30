# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from IGF_log import getlog,getloglocal
from rpw import revit,DB,UI
from pyrevit import script, forms
# from eventhandler import Legend_Normal,ExternalEvent,Legend_Duct,Legend_Color,Legend_Keynote
from System.Collections.Generic import List
from System.Windows.Input import ModifierKeys,Keyboard,Key

__title__ = "Legenden erstellen"
__doc__ = """

Legenden für ausgewählte Ansicht erstellen
[2022.05.31]
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

uidoc = revit.uidoc
doc = revit.doc
logger = script.get_logger()

def Legende_Filter():
    param_equality=DB.FilterStringEquals()
    param_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    param_prov=DB.ParameterValueProvider(param_id)
    param_value_rule=DB.FilterStringRule(param_prov,param_equality,'Legende',True)
    param_filter = DB.ElementParameterFilter(param_value_rule)
    return param_filter

legend = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType().WherePasses(Legende_Filter()).FirstElement()
if legend is None:
    dialog = UI.TaskDialog('Legenden')
    dialog.MainContent = 'Bitte eine Legendeansicht und eine Legendenkomponente erstellen. Es funktioniert nur wenn in Projekt zumindest eine Legende vonhanden ist.\nLegende erstellen: Ansicht -> Erstellen -> Legenden -> Legende\nLegendensymbol: Beschriften -> Detail -> Bauteil -> Legendenbauteil'
    dialog.Show()
    script.exit()


    
legendelement = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_LegendComponents).WhereElementIsNotElementType().FirstElement()
if legendelement is None:
    dialog = UI.TaskDialog('Legendenkomponente')
    dialog.MainContent = 'Bitte eine Legendenkomponente erstellen. Es funktioniert nur wenn in Projekt zumindest eine Legendenkomponente vonhanden ist. \nLegendensymbol: Beschriften -> Detail -> Bauteil -> Legendenbauteil'
    dialog.Show()
    script.exit()

class Ansicht(object):
    def __init__(self,elemid,Family,name):
        self.elemid = elemid
        self.name = name
        self.checked = False
        self.family = Family

class Kategorien(object):
    def __init__(self,elemid,name):
        self.elemid = elemid
        self.name = name
        self.checked = False

# Viewssource
VS = ObservableCollection[Ansicht]()
VS_Grundriss = ObservableCollection[Ansicht]()
VS_Schnitt = ObservableCollection[Ansicht]()
VS_Cate = ObservableCollection[Kategorien]()

def GetAllViews():
    _dict = {}
    views = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).ToElements()
    for v in views:
        if v.IsTemplate:continue
        name = v.Name
        fam = v.ViewType.ToString()
        if fam in ['FloorPlan','CeilingPlan']:
            fam = 'Grundriss'
        elif fam == 'Section':
            fam = 'Schnitt'
        else:
            continue
        _dict[name] = [v.Id,fam]
    for n in sorted(_dict.keys()):
        temp = Ansicht(_dict[n][0],_dict[n][1],n)
        VS.Add(temp)
        if _dict[n][1] == 'Grundriss':VS_Grundriss.Add(temp)
        else:VS_Schnitt.Add(temp)

GetAllViews()

VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_GenericModel,'Allgemeines Modell'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_ElectricalEquipment,'Elektrische Ausstattung'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_MechanicalEquipment,'HLS-Bauteile'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_DuctTerminal,'Luftdurchlässe'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_DuctAccessory,'Luftkanalzubehör'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_PipeAccessory,'Rohrzubehör'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_PlumbingFixtures,'Sanitärinstallationen'))
VS_Cate.Add(Kategorien(DB.BuiltInCategory.OST_Sprinklers,'Sprinkler'))

# Texttyp
all_text_types = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).WhereElementIsElementType().ToElements()
TEXTTYPE = {i.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(): i for i in all_text_types}

all_lines = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
LINES = {i.Name: i.GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id  for i in all_lines.SubCategories if i.Id.ToString() not in ['-2000066','-2000831','-2000079','-2009018','-2000045','-2009019','-2000077','-2000065']}

class LEGENDEN(forms.WPFWindow):
    def __init__(self):
        self.alltyxttyp = TEXTTYPE
        self.alllinetype = LINES
        self.allviews = VS
        self.allcate = VS_Cate
        self.altv = VS
        self.allgrundriss = VS_Grundriss
        self.allschnitt = VS_Schnitt
        self.legend_template = legend
        self.legengcomkonente_template = legendelement
        self.richtung = {'Grundriss':-8,'Vorne':-7,'Hinter':-6,'Links':-10,'Rechts':-9} #  -5: Schnitt
        self.detailgrad = {'Fine':3,'Medium':2,'Coarse':1,'Undefined':3}
        self.tempob = ObservableCollection[Ansicht]()
        self.cate_Liste = List[DB.BuiltInCategory]()

          
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.LB_Views.ItemsSource = self.allviews
        self.LB_Kate.ItemsSource = self.allcate
        self.tn.ItemsSource = sorted(self.alltyxttyp.keys())
        self.linie.ItemsSource = sorted(self.alllinetype.keys())
        self.linie2.ItemsSource = sorted(self.alllinetype.keys())
        self.comauswahlmanuell.ItemsSource = self.richtung.keys()
        self.beschreibung.ItemsSource = ['Familie','Typ','Familie und Typ']

    def suchchanged(self, sender, args):
        self.tempob.Clear()
        text_typ = self.filter.Text.upper()    

        if text_typ in ['',None]:
            self.LB_Views.ItemsSource = self.altv
            text_typ = self.filter.Text = ''

        for item in self.altv:
            if item.name.upper().find(text_typ) != -1:
                self.tempob.Add(item)

            self.LB_Views.ItemsSource = self.tempob
        self.LB_Views.Items.Refresh()

    def checkedchanged(self, sender, args):
        if self.grundcheck.IsChecked and self.schnittcheck.IsChecked:
            self.LB_Views.ItemsSource = self.allviews
            self.altv = self.allviews
        elif self.grundcheck.IsChecked is True and self.schnittcheck.IsChecked is False:
            self.LB_Views.ItemsSource = self.allgrundriss
            self.altv = self.allgrundriss
        elif self.grundcheck.IsChecked is False and self.schnittcheck.IsChecked is True:
            self.LB_Views.ItemsSource = self.allschnitt
            self.altv = self.allschnitt
        else:
            self.LB_Views.ItemsSource = ObservableCollection[Ansicht]()
            self.altv = ObservableCollection[Ansicht]()

        text = self.filter.Text
        self.filter.Text = ''
        self.filter.Text = text

    def viewscheckedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LB_Views.SelectedItem is not None:
            try:
                if sender.DataContext in self.LB_Views.SelectedItems:
                    for item in self.LB_Views.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.LB_Views.Items.Refresh()
                else:
                    pass
            except:
                pass
        return   

    def close(self, sender, args):
        self.Close()
        script.exit()

    def changetoansichtmass(self, sender, args):
        self.Massstabmanuell.IsEnabled = False
        self.ansicht.IsChecked = True
        self.manuell.IsChecked = False

    def changetomanuellMass(self, sender, args):
        self.Massstabmanuell.IsEnabled = True
        self.ansicht.IsChecked = False
        self.manuell.IsChecked = True
    
    def changetoansichtkom(self, sender, args):
        self.comauswahlmanuell.IsEnabled = False
        self.comansicht.IsChecked = True
        self.commanuell.IsChecked = False

    def changetomanuellkom(self, sender, args):
        self.comauswahlmanuell.IsEnabled = True
        self.comansicht.IsChecked = False
        self.commanuell.IsChecked = True

    def movewindow(self, sender, args):
        self.DragMove()

    def catechanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LB_Kate.SelectedItem is not None:
            try:
                if sender.DataContext in self.LB_Kate.SelectedItems:
                    for item in self.LB_Kate.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.LB_Kate.Items.Refresh()
                else:
                    pass
            except:
                pass
        temp_liste = [el.elemid for el in self.allcate if el.checked == True]
        self.cate_Liste = List[DB.BuiltInCategory](temp_liste)
        return  

    def create_text(self, view, x ,y ,text, text_note_type):
        text = '-' if not text else text
        text_note = DB.TextNote.Create(doc, view.Id, DB.XYZ(x, y, 0),text, text_note_type.Id)
        return text_note

    def create_line(self, view, P0, P1):
        line_constructor = DB.Line.CreateBound(P0, P1)
        line = doc.Create.NewDetailCurve(view, line_constructor)

        return line
    
    def Setkey(self, sender, args):   
        if ((args.Key >= Key.D0 and args.Key <= Key.D9) or (args.Key >= Key.NumPad0 and args.Key <= Key.NumPad9) \
            or args.Key == Key.Delete or args.Key == Key.Back or args.Key == Key.Enter):
            args.Handled = False
        else:
            args.Handled = True

    def start(self, sender, args):
        """Legende erstellen"""
        self.Close()
        _liste = []
        for el in self.LB_Views.Items:
            if el.checked:
                _liste.append(el)
        if len(_liste) == 0:
            return
        t = DB.Transaction(doc,'Legenden erstellen')
        t.Start()
        for el_customView in _liste:
            __Liste = []
            _Dict_Neu = {}
            Families = DB.FilteredElementCollector(doc,el_customView.elemid).WherePasses(DB.ElementMulticategoryFilter(self.cate_Liste)).WhereElementIsNotElementType().ToElements()
            if Families.Count == 0:
                continue
            Symbols = []
            fams = []
            for fam in Families:
                if self.family.IsChecked:
                    if fam.Symbol.Family.Id.ToString() not in fams:
                        fams.append(fam.Symbol.Family.Id.ToString())
                        Symbols.append(fam.Symbol)
                    else:
                        continue
                else:
                    if fam.Symbol.Id.ToString() not in fams:
                        fams.append(fam.Symbol.Id.ToString())
                        Symbols.append(fam.Symbol)
                    else:
                        continue
                    
            for fam in Symbols:
                if self.beschreibung.Text == 'Familie':
                    name = fam.FamilyName
                elif self.beschreibung.Text == 'Typ':
                    name = fam.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                elif self.beschreibung.Text == 'Familie und Typ':
                    name = fam.FamilyName + ': ' + fam.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                else:
                    para = fam.LookupParameter(self.beschreibung.Text)
                    if para is not None:
                        if para.StorageType.ToString() in ['Integer','Double','ElementId']:
                            name = para.AsValueString()
                        elif para.StorageType.ToString() == 'String':
                            name = para.AsString()
                    else:
                        name = fam.FamilyName
                __Liste.append([fam,name])
                
            
            # Legend Name    
            l = legend.Duplicate(DB.ViewDuplicateOption.Duplicate)
            _View_RvtLegend = doc.GetElement(l)
            legendname = 'Legend - ' + el_customView.name
            rename = False
            while (not rename):
                try:
                    _View_RvtLegend.Name = legendname
                    rename = True
                except:
                    legendname += ' (1)'
            ########################################

            # Maßstab        
            if self.ansicht.IsChecked:_View_RvtLegend.Scale = doc.GetElement(el_customView.elemid).Scale
            else:
                try:
                    if self.Massstabmanuell.Text != '0':_View_RvtLegend.Scale = int(round(float(self.Massstabmanuell.Text)))
                    else:
                        logger.error('Maßstab kann nicht 1:0 sein!')
                        _View_RvtLegend.Scale = doc.GetElement(el_customView.elemid).Scale
                except:
                    
                    _View_RvtLegend.Scale = doc.GetElement(el_customView.elemid).Scale
            #######################################
            try:
                sourcev = legendelement.OwnerViewId

                # legendenertsellen
                for el_customListe in __Liste:
                    temp = DB.ElementTransformUtils.CopyElements(doc.GetElement(sourcev),List[DB.ElementId]([legendelement.Id]),doc.GetElement(l),None,None)[0]
                    doc.GetElement(temp).get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_DETAIL_LEVEL).Set(self.detailgrad[doc.GetElement(el_customView.elemid).DetailLevel.ToString()])
                    box = doc.GetElement(temp).get_BoundingBox(_View_RvtLegend)
                    x = (box.Max.X + box.Min.X) / 2
                    y = (box.Max.Y + box.Min.Y) / 2
                    doc.GetElement(temp).Location.Move(DB.XYZ(0-x,0-y,0))
                    _Dict_Neu[temp] = el_customListe
                #######################################
                abstand = 0
                maxX = 0
                Liste_TextNote = []
                Liste_TextNoteWithSymbol = []
                for n,elid in enumerate(_Dict_Neu.keys()):
                    el_Rvt_Legendensymbol = doc.GetElement(elid)
                    text = _Dict_Neu[elid][1]
                    el_Rvt_Legendensymbol.LookupParameter('Komponententyp').Set(_Dict_Neu[elid][0].Id)

                    # Legendensymbol ansichtsrichtung
                    if self.comansicht.IsChecked:
                        if el_customView.family == 'Grundriss':
                            try:el_Rvt_Legendensymbol.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(-8)
                            except:pass
                        else:
                            el_RvtView_Richtung = doc.GetElement(el_customView.elemid).ViewDirection
                            if el_RvtView_Richtung.X == 1:
                                try:el_Rvt_Legendensymbol.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(-9)
                                except:pass
                            elif el_RvtView_Richtung.X == -1:
                                try:el_Rvt_Legendensymbol.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(-10)
                                except:pass
                            elif el_RvtView_Richtung.Y == -1:
                                try:el_Rvt_Legendensymbol.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(-7)
                                except:pass
                            elif el_RvtView_Richtung.Y == 1:
                                try:el_Rvt_Legendensymbol.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(-6)
                                except:pass
                            else:
                                pass
                    elif self.commanuell.IsChecked:
                        richtung = self.richtung[self.comauswahlmanuell.SelectedItem.ToString()]
                        try:el_Rvt_Legendensymbol.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(richtung)
                        except:pass
                    
                    ################################

                    # Create Beschreibung 
                    textnote = self.create_text(_View_RvtLegend,0,0,text,self.alltyxttyp[self.tn.SelectedItem.ToString()])
                    Liste_TextNote.append(textnote)
                    Liste_TextNoteWithSymbol.append([textnote,el_Rvt_Legendensymbol])
                    doc.Regenerate()
                    box0 = el_Rvt_Legendensymbol.get_BoundingBox(_View_RvtLegend)
                    box1 = textnote.get_BoundingBox(_View_RvtLegend)
                    if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                    else:abstand -= (box1.Max.Y - box1.Min.Y) / 2
                    maxX = max(box0.Max.X,maxX)

                    if n == 0:
                        textnote.Location.Move(DB.XYZ(0,0-(box1.Max.Y + box1.Min.Y) / 2,0))
                        abstand -= 2
                        continue
                    else:
                        el_Rvt_Legendensymbol.Location.Move(DB.XYZ(0,abstand,0))
                        textnote.Location.Move(DB.XYZ(0,abstand-(box1.Max.Y + box1.Min.Y) / 2,0))
                        if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                        else:abstand -= (box1.Max.Y - box1.Min.Y) / 2
                        abstand -= 2

                for el_Rvt_Textnote in Liste_TextNote:
                    el_Rvt_Textnote.Location.Move(DB.XYZ(maxX+2,0,0))
                    doc.Regenerate()

                maxX_OuterLine = 0     
                minX_OuterLine = 0  
                maxY_OuterLine = 0  
                minY_OuterLine = 0               
                
                for n,el_Liste in enumerate(Liste_TextNoteWithSymbol):
                    box0 = el_Liste[0].get_BoundingBox(_View_RvtLegend)
                    box1 = el_Liste[1].get_BoundingBox(_View_RvtLegend)
                    maxX_OuterLine = max(maxX_OuterLine,box0.Max.X,box1.Max.X)
                    minX_OuterLine = min(minX_OuterLine,box0.Min.X,box1.Min.X)
                    maxY_OuterLine = max(maxY_OuterLine,box0.Max.Y,box1.Max.Y)
                    minY_OuterLine = min(minY_OuterLine,box0.Min.Y,box1.Min.Y)
                
                # Create Linie
                if self.auline.IsChecked:
                    LineStyle_out = self.alllinetype[self.linie.SelectedItem.ToString()]
                    line0 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+1,minY_OuterLine-1,0),DB.XYZ(maxX_OuterLine+1,maxY_OuterLine+1,0))
                    line1 = self.create_line(_View_RvtLegend,DB.XYZ(minX_OuterLine-1,minY_OuterLine-1,0),DB.XYZ(minX_OuterLine-1,maxY_OuterLine+1,0))
                    line2 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+1,maxY_OuterLine+1,0),DB.XYZ(minX_OuterLine-1,maxY_OuterLine+1,0))
                    line3 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+1,minY_OuterLine-1,0),DB.XYZ(minX_OuterLine-1,minY_OuterLine-1,0))
                    try:line0.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line1.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line2.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line3.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                
                if self.inline.IsChecked:
                    LineStyle_in = self.alllinetype[self.linie2.SelectedItem.ToString()]
                    line_temp = self.create_line(_View_RvtLegend,DB.XYZ(maxX+1,minY_OuterLine-1,0),DB.XYZ(maxX+1,maxY_OuterLine+1,0))
                    try:line_temp.LineStyle = doc.GetElement(LineStyle_in)
                    except:pass

                    for n,el_Liste in enumerate(Liste_TextNoteWithSymbol):
                        if n == Liste_TextNoteWithSymbol.Count - 1:
                            continue
                        box0 = el_Liste[0].get_BoundingBox(_View_RvtLegend)
                        box1 = el_Liste[1].get_BoundingBox(_View_RvtLegend)
                        Y_line = min(box0.Min.Y,box1.Min.Y)
                        line_neu = self.create_line(_View_RvtLegend,DB.XYZ(minX_OuterLine-1,Y_line-1,0),DB.XYZ(maxX_OuterLine+1,Y_line-1,0))
                        try:line_neu.LineStyle = doc.GetElement(LineStyle_in)
                        except:pass
                ###########################
            except Exception as e:logger.error(e)

        t.Commit()
        
       
legenderstellen = LEGENDEN()
try:legenderstellen.ShowDialog()
except Exception as e:
    legenderstellen.Close()
    logger.error(e)