# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from IGF_log import getlog,getloglocal
from rpw import revit,DB,UI
from pyrevit import script, forms
# from eventhandler import Legend_Normal,ExternalEvent,Legend_Duct,Legend_Color,Legend_Keynote
from System.Collections.Generic import List
from System.Windows.Input import ModifierKeys,Keyboard,Key
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs

__title__ = "Daten übertragen"
__doc__ = """

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



class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

class Itemtemplate(TemplateItemBase):
    def __init__(self,name,checked = False):
        TemplateItemBase.__init__(self)
        self.name = name
        self._checked = checked

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')

class Ansicht(Itemtemplate):
    def __init__(self,elemid,Family,name):
        Itemtemplate.__init__(self,name)
        self.elemid = elemid
        self.family = Family
        self.Systemtyp = self.get_Systemtyp()
    
    def get_Systemtyp(self):
        Liste = []
        systems = DB.FilteredElementCollector(doc,self.elemid).OfClass(DB.MEPSystem).WhereElementIsNotElementType().ToElements()
        for el in systems:
            if el.GetTypeId().IntegerValue == -1:
                continue
            systemtyp = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()

            if systemtyp not in Liste:
                Liste.append(systemtyp)
        return Liste

class SystemClass(Itemtemplate):
    def __init__(self,name):
        Itemtemplate.__init__(self,name)

# Viewssource
VS = ObservableCollection[Ansicht]()
VS_Grundriss = ObservableCollection[Ansicht]()
VS_Schnitt = ObservableCollection[Ansicht]()
VS_System = ObservableCollection[SystemClass]()

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

# Texttyp

class LEGENDEN(forms.WPFWindow):
    def __init__(self):
        self.alltyxttyp = TEXTTYPE
        self.alllinetype = LINES
        self.allviews = VS
        self.allsystem = VS_System
        self.Liste_System = []
        self.altv = VS
        self.allgrundriss = VS_Grundriss
        self.allschnitt = VS_Schnitt
        self.legend_template = legend
        self.detailgrad = {'Fine':3,'Medium':2,'Coarse':1,'Undefined':3}
        self.tempob = ObservableCollection[Ansicht]()
          
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.LB_Views.ItemsSource = self.allviews
        self.LB_System.ItemsSource = self.allsystem
        self.tn.ItemsSource = sorted(self.alltyxttyp.keys())
        try:self.tn.SelectedIndex = 0
        except:pass
        self.texttyp_tit.ItemsSource = sorted(self.alltyxttyp.keys())
        self.linie.ItemsSource = sorted(self.alllinetype.keys())
        self.linie2.ItemsSource = sorted(self.alllinetype.keys())
        self.ueberauswahl.ItemsSource = ['Legende Lüftung','Legende Heizung/Kälte','Legende GA','Legende Sanitär','Legende LB Medien','Legende Gase', 'Legende Sprinkler']

    def suchchanged(self, sender, args):
        self.tempob.Clear()
        text_typ = self.filter.Text    

        if text_typ in ['',None]:
            self.LB_Views.ItemsSource = self.altv
            text_typ = self.filter.Text = ''

        for item in self.altv:
            if item.name.upper().find(text_typ.upper()) != -1:
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
        Liste = []
        self.allsystem.Clear()

        if self.LB_Views.SelectedItem is not None:
            try:
                if sender.DataContext in self.LB_Views.SelectedItems:
                    for item in self.LB_Views.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                else:
                    pass
            except:
                pass
        for el in self.allviews:
            if el.checked:
                Liste.extend(el.Systemtyp) 
        Liste = set(Liste)
        Liste = list(Liste)
        for el in sorted(Liste):
            self.allsystem.Add(SystemClass(el))
        self.LB_System.ItemsSource = self.allsystem
        return   
    
    def systemcheckedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LB_System.SelectedItem is not None:
            try:
                if sender.DataContext in self.LB_System.SelectedItems:
                    for item in self.LB_System.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                else:
                    pass
            except:
                pass
        temp_liste = [el.name for el in self.allsystem if el.checked == True]
        self.Liste_System = temp_liste
 
    def close(self, sender, args):
        self.Close()
        script.exit()

    def movewindow(self, sender, args):
        self.DragMove()

    def create_text(self, view, x ,y ,text, text_note_type):
        text = '-' if not text else text
        text_note = DB.TextNote.Create(doc, view.Id, DB.XYZ(x, y, 0),text, text_note_type.Id)
        return text_note

    def create_line(self, view, P0, P1):
        line_constructor = DB.Line.CreateBound(P0, P1)
        line = doc.Create.NewDetailCurve(view, line_constructor)

        return line
    
    def titelchange(self, sender, args):
        if not sender.IsChecked:
            self.texttyp_tit.IsEnabled = False
            self.ueberauswahl.IsEnabled = False
        else:
            self.texttyp_tit.IsEnabled = True
            self.ueberauswahl.IsEnabled = True

    def start(self, sender, args):
        """Legende erstellen"""
        self.Close()
        _liste = []
        for el in self.LB_Views.Items:
            if el.checked:
                _liste.append(el)
        if len(_liste) == 0:
            return
        
        textsize = self.alltyxttyp[self.tn.SelectedItem.ToString()].get_Parameter(DB.BuiltInParameter.TEXT_SIZE).AsDouble()
        linglength = 20 * textsize
        v_margin = 2 * textsize
        h_margin = 4 * textsize

        t = DB.Transaction(doc,'Legenden erstellen')
        t.Start()

        if self.einzeln.IsChecked:
            for el_customView in _liste:
                _Dict_Neu = {}
                systemtyps = [i for i in el.Systemtyp if i in self.Liste_System]
                if len(systemtyps) == 0:
                    continue
                                
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
                _View_RvtLegend.Scale = 1    
                                
                #######################################
                # try:
                    # legendenertsellen
                    # Systemlinie erstellen
                for systemtyp in systemtyps:
                    if systemtyp in SYSLINES.keys():
                        line = SYSLINES[systemtyp]
                    else:
                        line = doc.Settings.Categories.NewSubcategory(all_lines,systemtyp)
                        doc.Regenerate()
                    line.SetLineWeight(4,DB.GraphicsStyleType.Projection)
                    try:line.LineColor = SYSTEMTYPE[systemtyp].LineColor
                    except Exception as e:print(e)
                    if SYSTEMTYPE[systemtyp].LinePatternId.IntegerValue != -1:
                        line.SetLinePatternId( SYSTEMTYPE[systemtyp].LinePatternId ,DB.GraphicsStyleType.Projection )
                        
                    SYSLINES[systemtyp] = line
                    doc.Regenerate()
                    neuline = self.create_line(_View_RvtLegend,DB.XYZ(0,0,0),DB.XYZ(linglength,0,0))
                    try:neuline.LineStyle = line.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                    except Exception as e:print(e)
                    _Dict_Neu[systemtyp] = neuline

                #######################################
                abstand = 0
                maxX = 0
                Liste_TextNote = []
                Liste_TextNoteWithSymbol = []
                for n,systemname in enumerate(_Dict_Neu.keys()):
                    systemtyp = SYSTEMTYPE[systemname]
                    el_Rvt_Legendensymbol = _Dict_Neu[systemname]
                    igf_systemname = systemtyp.LookupParameter('IGF_X_SystemName')
                    igf_systemKuerzel = systemtyp.LookupParameter('IGF_X_SystemKürzel')
                    if igf_systemname and igf_systemKuerzel:
                        text = igf_systemname.AsString() + ' ('+igf_systemKuerzel.AsString()+')'
                    else:
                        text = systemname

                    
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
                        abstand -= v_margin
                        continue
                    else:
                        el_Rvt_Legendensymbol.Location.Move(DB.XYZ(0,abstand,0))
                        textnote.Location.Move(DB.XYZ(0,abstand-(box1.Max.Y + box1.Min.Y) / 2,0))
                        if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                        else:abstand -= (box1.Max.Y - box1.Min.Y) / 2
                        abstand -= v_margin

                for el_Rvt_Textnote in Liste_TextNote:
                    el_Rvt_Textnote.Location.Move(DB.XYZ(maxX+h_margin,0,0))
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
                
                maxY_OuterLine_temp = maxY_OuterLine
                if self.tit.IsChecked:
                    ueberschrift = self.create_text(_View_RvtLegend,0,0,self.ueberauswahl.Text,self.alltyxttyp[self.texttyp_tit.SelectedItem.ToString()])
                    doc.Regenerate()
                    box = ueberschrift.get_BoundingBox(_View_RvtLegend)
                    height = (box.Max.Y-box.Min.Y)/2
                    x = (box.Max.X+box.Min.X)/2
                    y = (box.Max.Y+box.Min.Y)/2
                    ueberschrift.Location.Move(DB.XYZ((maxX_OuterLine+minX_OuterLine)/2-x,maxY_OuterLine+v_margin+height-y,0))
                    doc.Regenerate()
                    box_neu = ueberschrift.get_BoundingBox(_View_RvtLegend)
                    maxX_OuterLine = max(maxX_OuterLine,box_neu.Max.X)
                    maxY_OuterLine = max(maxY_OuterLine,box_neu.Max.Y)
                    minX_OuterLine = min(minX_OuterLine,box_neu.Min.X)

                
                # Create Linie
                if self.auline.IsChecked:
                    LineStyle_out = self.alllinetype[self.linie.SelectedItem.ToString()]
                    line0 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(maxX_OuterLine+h_margin/2.0,maxY_OuterLine+h_margin/2.0,0))
                    line1 = self.create_line(_View_RvtLegend,DB.XYZ(minX_OuterLine-h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,maxY_OuterLine+h_margin/2.0,0))
                    line2 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,maxY_OuterLine+h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,maxY_OuterLine+h_margin/2.0,0))
                    line3 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,minY_OuterLine-h_margin/2.0,0))
                    line4 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,maxY_OuterLine_temp+h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,maxY_OuterLine_temp+h_margin/2.0,0))
                    try:line0.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line1.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line2.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line3.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line4.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                
                if self.inline.IsChecked:
                    LineStyle_in = self.alllinetype[self.linie2.SelectedItem.ToString()]
                    line_temp = self.create_line(_View_RvtLegend,DB.XYZ(maxX+h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(maxX+h_margin/2.0,maxY_OuterLine_temp+h_margin/2.0,0))
                    try:line_temp.LineStyle = doc.GetElement(LineStyle_in)
                    except:pass

                    for n,el_Liste in enumerate(Liste_TextNoteWithSymbol):
                        if n == Liste_TextNoteWithSymbol.Count - 1:
                            continue
                        box0 = el_Liste[0].get_BoundingBox(_View_RvtLegend)
                        box1 = el_Liste[1].get_BoundingBox(_View_RvtLegend)
                        Y_line = min(box0.Min.Y,box1.Min.Y)
                        line_neu = self.create_line(_View_RvtLegend,DB.XYZ(minX_OuterLine-v_margin/2.0,Y_line-h_margin/2.0,0),DB.XYZ(maxX_OuterLine+v_margin/2.0,Y_line-h_margin/2.0,0))
                        try:line_neu.LineStyle = doc.GetElement(LineStyle_in)
                        except:pass
                ###########################
                # except Exception as e:logger.error(e)
        else:
            
            _Dict_Neu = {}
            systemtyps = self.Liste_Systemz
            if len(systemtyps) == 0:
                t.Commit()
                return
                            
            # Legend Name    
            l = legend.Duplicate(DB.ViewDuplicateOption.Duplicate)
            _View_RvtLegend = doc.GetElement(l)
            legendname = 'Legend - ' + doc.ProjectInformation.Name
            rename = False
            while (not rename):
                try:
                    _View_RvtLegend.Name = legendname
                    rename = True
                except:
                    legendname += ' (1)'
            ########################################

            # Maßstab    
            _View_RvtLegend.Scale = 1    
                            
            #######################################
            try:
                # legendenertsellen
                # Systemlinie erstellen
                for systemtyp in systemtyps:
                    if systemtyp in SYSLINES.keys():
                        line = SYSLINES[systemtyp]
                    else:
                        line = doc.Settings.Categories.NewSubcategory(all_lines,systemtyp)
                        doc.Regenerate()
                    line.SetLineWeight(4,DB.GraphicsStyleType.Projection)
                    try:line.LineColor = SYSTEMTYPE[systemtyp].LineColor
                    except Exception as e:print(e)
                    if SYSTEMTYPE[systemtyp].LinePatternId.IntegerValue != -1:
                        line.SetLinePatternId( SYSTEMTYPE[systemtyp].LinePatternId ,DB.GraphicsStyleType.Projection )
                        break
                    SYSLINES[systemtyp] = line
                    doc.Regenerate()
                    neuline = self.create_line(_View_RvtLegend,DB.XYZ(0,0,0),DB.XYZ(0.5,0,0))
                    try:neuline.LineStyle = line.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                    except Exception as e:print(e)
                    _Dict_Neu[systemtyp] = neuline

                #######################################
                abstand = 0
                maxX = 0
                Liste_TextNote = []
                Liste_TextNoteWithSymbol = []
                for n,systemname in enumerate(_Dict_Neu.keys()):
                    systemtyp = SYSTEMTYPE[systemname]
                    el_Rvt_Legendensymbol = _Dict_Neu[systemname]
                    igf_systemname = systemtyp.LookupParameter('IGF_X_SystemName')
                    igf_systemKuerzel = systemtyp.LookupParameter('IGF_X_SystemKürzel')
                    if igf_systemname and igf_systemKuerzel:
                        text = igf_systemname.AsString() + ' ()'.format(igf_systemKuerzel.AsString())
                    else:
                        text = systemname

                    
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
                        abstand -= 0.06
                        continue
                    else:
                        el_Rvt_Legendensymbol.Location.Move(DB.XYZ(0,abstand,0))
                        textnote.Location.Move(DB.XYZ(0,abstand-(box1.Max.Y + box1.Min.Y) / 2,0))
                        if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                        else:abstand -= (box1.Max.Y - box1.Min.Y) / 2
                        abstand -= 0.06

                for el_Rvt_Textnote in Liste_TextNote:
                    el_Rvt_Textnote.Location.Move(DB.XYZ(maxX+0.5,0,0))
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
                
                maxY_OuterLine_temp = maxY_OuterLine
                if self.tit.IsChecked:
                    ueberschrift = self.create_text(_View_RvtLegend,0,0,self.ueberauswahl.Text,self.alltyxttyp[self.texttyp_tit.SelectedItem.ToString()])
                    doc.Regenerate()
                    box = ueberschrift.get_BoundingBox(_View_RvtLegend)
                    height = (box.Max.Y-box.Min.Y)/2
                    x = (box.Max.X+box.Min.X)/2
                    y = (box.Max.Y+box.Min.Y)/2
                    ueberschrift.Location.Move(DB.XYZ((maxX_OuterLine+minX_OuterLine)/2-x,maxY_OuterLine+2+height-y,0))
                    doc.Regenerate()
                    box_neu = ueberschrift.get_BoundingBox(_View_RvtLegend)
                    maxX_OuterLine = max(maxX_OuterLine,box_neu.Max.X)
                    maxY_OuterLine = max(maxY_OuterLine,box_neu.Max.Y)
                    minX_OuterLine = min(minX_OuterLine,box_neu.Min.X)

                
                # Create Linie
                if self.auline.IsChecked:
                    LineStyle_out = self.alllinetype[self.linie.SelectedItem.ToString()]
                    line0 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(maxX_OuterLine+h_margin/2.0,maxY_OuterLine+h_margin/2.0,0))
                    line1 = self.create_line(_View_RvtLegend,DB.XYZ(minX_OuterLine-h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,maxY_OuterLine+h_margin/2.0,0))
                    line2 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,maxY_OuterLine+h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,maxY_OuterLine+h_margin/2.0,0))
                    line3 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,minY_OuterLine-h_margin/2.0,0))
                    line4 = self.create_line(_View_RvtLegend,DB.XYZ(maxX_OuterLine+h_margin/2.0,maxY_OuterLine_temp+h_margin/2.0,0),DB.XYZ(minX_OuterLine-h_margin/2.0,maxY_OuterLine_temp+h_margin/2.0,0))
                    try:line0.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line1.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line2.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line3.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                    try:line4.LineStyle = doc.GetElement(LineStyle_out)
                    except:pass
                
                if self.inline.IsChecked:
                    LineStyle_in = self.alllinetype[self.linie2.SelectedItem.ToString()]
                    line_temp = self.create_line(_View_RvtLegend,DB.XYZ(maxX+h_margin/2.0,minY_OuterLine-h_margin/2.0,0),DB.XYZ(maxX+h_margin/2.0,maxY_OuterLine_temp+h_margin/2.0,0))
                    try:line_temp.LineStyle = doc.GetElement(LineStyle_in)
                    except:pass

                    for n,el_Liste in enumerate(Liste_TextNoteWithSymbol):
                        if n == Liste_TextNoteWithSymbol.Count - 1:
                            continue
                        box0 = el_Liste[0].get_BoundingBox(_View_RvtLegend)
                        box1 = el_Liste[1].get_BoundingBox(_View_RvtLegend)
                        Y_line = min(box0.Min.Y,box1.Min.Y)
                        line_neu = self.create_line(_View_RvtLegend,DB.XYZ(minX_OuterLine-h_margin/2.0,Y_line-h_margin/2.0,0),DB.XYZ(maxX_OuterLine+h_margin/2.0,Y_line-h_margin/2.0,0))
                        try:line_neu.LineStyle = doc.GetElement(LineStyle_in)
                        except:pass
                ###########################
            except Exception as e:logger.error(e)

        t.Commit()
        
       
legenderstellen = LEGENDEN()
legenderstellen.ShowDialog()
# try:legenderstellen.ShowDialog()
# except Exception as e:
#     legenderstellen.Close()
#     logger.error(e)