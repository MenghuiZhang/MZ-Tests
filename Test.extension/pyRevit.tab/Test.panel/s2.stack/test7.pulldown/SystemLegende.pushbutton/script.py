# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
from IGF_log import getlog,getloglocal
from rpw import revit,DB,UI
from pyrevit import script, forms
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs

__title__ = "Systemlegenden erstellen"
__doc__ = """

Legenden für ausgewählte Ansicht erstellen
[2023.02.02]
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



SYSLINES = {}
LINES = {}

systemtyp = DB.FilteredElementCollector(doc).OfClass(DB.MEPSystemType).ToElements()
SYSTEMTYPE = {i.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():i for i in systemtyp}


all_lines = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
for i in all_lines.SubCategories:
    if i.Id.ToString() not in ['-2000066','-2000831','-2000079','-2009018','-2000045','-2009019','-2000077','-2000065']:
        SYSLINES[i.Name] = i
        LINES[i.Name] = i.GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id

all_text_types = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).WhereElementIsElementType().ToElements()
TEXTTYPE = {i.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(): i for i in all_text_types}

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
        self.ueberauswahl.ItemsSource = ['Legende Lüftung','Legende Heizung/Kälte','Legende GA','Legende Sanitär','Legende LB Medien','Legende Gase', 'Legende Sprinkler','Legende Dampf','Legende Heizung','Legende Kälte',]

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

    def gewerkchecked(self,sender,e):
        checked = sender.IsChecked
        if checked:
            self.ueberauswahl.IsEnabled = False
        else:
            if self.tit.IsChecked:
                self.ueberauswahl.IsEnabled = True
            else:
                self.ueberauswahl.IsEnabled = False

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
            if self.systemgetrennt.IsChecked:
                self.ueberauswahl.IsEnabled = True
    
    def create_header(self,_View_RvtLegend,header):
        ueberschrift = self.create_text(_View_RvtLegend,0,0,header,self.alltyxttyp[self.texttyp_tit.SelectedItem.ToString()])
        doc.Regenerate()
        return ueberschrift
    
    def modifyLine(self,systemtyp):
        linename = 'Legende_'+ systemtyp
        if linename in SYSLINES.keys():
            line = SYSLINES[linename]
        else:
            line = doc.Settings.Categories.NewSubcategory(all_lines,linename)
            doc.Regenerate()
        line.SetLineWeight(4,DB.GraphicsStyleType.Projection)
        try:line.LineColor = SYSTEMTYPE[systemtyp].LineColor
        except Exception as e:logger.error(e)
        if SYSTEMTYPE[systemtyp].LinePatternId.IntegerValue != -1:
            line.SetLinePatternId( SYSTEMTYPE[systemtyp].LinePatternId ,DB.GraphicsStyleType.Projection )
            
        SYSLINES[systemtyp] = line
        doc.Regenerate()
        return line
    
    def create_line_and_text(self,_View_RvtLegend,linglength,line,systemname,h_margin):
        neuline = self.create_line(_View_RvtLegend,DB.XYZ(0,0,0),DB.XYZ(linglength,0,0))
        try:neuline.LineStyle = line.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        except Exception as e:logger.error(e)
        systemtyp = SYSTEMTYPE[systemname]
        igf_systemname = systemtyp.LookupParameter('IGF_X_SystemName')
        igf_systemKuerzel = systemtyp.LookupParameter('IGF_X_SystemKürzel')
        if igf_systemname and igf_systemKuerzel:
            text0 =  igf_systemname.AsString()
            text1 = igf_systemKuerzel.AsString()
            if not text0:text0 = ''
            if not text1:text1 = ''
            text = text0 + ' ('+text1+')'
        else:
            text = systemname

        textnote = self.create_text(_View_RvtLegend,linglength+h_margin,0,text,self.alltyxttyp[self.tn.SelectedItem.ToString()])
        doc.Regenerate()
        return [neuline,textnote]
    
    def create_aussen_kanten_y(self,max_x_textnote,max_x_header,last_textnote,uebershcrift,_View_RvtLegend,h_margin,v_margin):
        if max_x_textnote >= max_x_header:
            min_x = 0
            max_x = max_x_textnote
        else:
            max_x = max_x_header
            min_x = max_x_textnote - max_x_header
        
        box0 = uebershcrift.get_BoundingBox(_View_RvtLegend)
        max_y = box0.Max.Y
        box0.Dispose()
        box1 = last_textnote.get_BoundingBox(_View_RvtLegend)
        min_y = box1.Min.Y
        box1.Dispose()    

        LineStyle_out = self.alllinetype[self.linie.SelectedItem.ToString()]
        line0 = self.create_line(_View_RvtLegend,DB.XYZ(max_x+h_margin/2.0,max_y+v_margin/2.0,0),DB.XYZ(max_x+h_margin/2.0,min_y-v_margin/2.0,0))
        line1 = self.create_line(_View_RvtLegend,DB.XYZ(min_x-h_margin/2.0,min_y-v_margin/2.0,0),DB.XYZ(min_x-h_margin/2.0,max_y+v_margin/2.0,0))
        try:line0.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        try:line1.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        line0.Dispose()
        line1.Dispose()
    
    def create_aussen_kanten_x(self,max_x_textnote,max_x_header,ueberschrift,_View_RvtLegend,h_margin,v_margin):
        if max_x_textnote >= max_x_header:
            min_x = 0
            max_x = max_x_textnote
        else:
            max_x = max_x_header
            min_x = max_x_textnote - max_x_header
        
        box = ueberschrift.get_BoundingBox(_View_RvtLegend)
        min_y = box.Min.Y
        max_y = box.Max.Y
        

        LineStyle_out = self.alllinetype[self.linie.SelectedItem.ToString()]
        line0 = self.create_line(_View_RvtLegend,DB.XYZ(max_x+h_margin/2.0,max_y+v_margin/2.0,0),DB.XYZ(min_x-h_margin/2.0,max_y+v_margin/2.0,0))
        line1 = self.create_line(_View_RvtLegend,DB.XYZ(max_x+h_margin/2.0,min_y-v_margin/2.0,0),DB.XYZ(min_x-h_margin/2.0,min_y-v_margin/2.0,0))
        try:line0.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        try:line1.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        box.Dispose()
        line0.Dispose()
        line1.Dispose()
    
    def create_mittel_linie_y(self,linglength,h_margin,_View_RvtLegend,v_margin,ueberschrift_0,ueberschrift_1 = None,last_item = None):
        mitte_x = linglength + h_margin/2
        box0 = ueberschrift_0.get_BoundingBox(_View_RvtLegend)
        start_y = box0.Min.Y - v_margin/2.0
        if ueberschrift_1:
            box1 = ueberschrift_1.get_BoundingBox(_View_RvtLegend)
            end_y = box1.Max.Y + v_margin/2.0
        else:
            box1 = last_item.get_BoundingBox(_View_RvtLegend)
            end_y = box1.Min.Y - v_margin/2.0

        LineStyle_out = self.alllinetype[self.linie2.SelectedItem.ToString()]
        line0 = self.create_line(_View_RvtLegend,DB.XYZ(mitte_x,start_y,0),DB.XYZ(mitte_x,end_y,0))
        try:line0.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        line0.Dispose()
    
    def create_mittel_linie_x(self,max_x_textnote,max_x_header,_View_RvtLegend,v_margin,textnote,h_marin):
        if max_x_textnote >= max_x_header:
            min_x = 0
            max_x = max_x_textnote
        else:
            max_x = max_x_header
            min_x = max_x_textnote - max_x_header
        
    
        box0 = textnote.get_BoundingBox(_View_RvtLegend)
        y = box0.Min.Y - v_margin/2.0
        box0.Dispose()

        LineStyle_out = self.alllinetype[self.linie2.SelectedItem.ToString()]
        line0 = self.create_line(_View_RvtLegend,DB.XYZ(min_x-h_marin/2,y,0),DB.XYZ(max_x+h_marin/2,y,0))
        try:line0.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        line0.Dispose()

    def create_last_aussen_kanten_x(self,max_x_textnote,max_x_header,_View_RvtLegend,v_margin,textnote,h_marin):
        if max_x_textnote >= max_x_header:
            min_x = 0
            max_x = max_x_textnote
        else:
            max_x = max_x_header
            min_x = max_x_textnote - max_x_header
        
    
        box0 = textnote.get_BoundingBox(_View_RvtLegend)
        y = box0.Min.Y - v_margin/2.0
        box0.Dispose()

        LineStyle_out = self.alllinetype[self.linie.SelectedItem.ToString()]
        line0 = self.create_line(_View_RvtLegend,DB.XYZ(min_x-h_marin/2,y,0),DB.XYZ(max_x+h_marin/2,y,0))
        try:line0.LineStyle = doc.GetElement(LineStyle_out)
        except:pass
        line0.Dispose()

    def start(self, sender, args):
        """Legende erstellen"""
        if len(self.Liste_System) == 0:
            UI.TaskDialog.Show('Error','Kein System ausgewählt!')
            return

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
                systemtyps = [i for i in el.Systemtyp if i in self.Liste_System]
                dict_Systemtyp = {}
                if self.tit.IsChecked:
                    if self.systemgetrennt.IsChecked:
                        dict_Systemtyp[self.ueberauswahl.Text] = systemtyps
                    else:
                        for systemtyp in systemtyps:
                            rvt_gk = SYSTEMTYPE[systemtyp].LookupParameter('IGF_X_Gewerkkürzel').AsString()
                            if rvt_gk == 'B':
                                if 'Legende Sprinkler' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Sprinkler'] = []
                                dict_Systemtyp['Legende Sprinkler'].append(systemtyp)
                            elif rvt_gk == 'D':
                                if 'Legende Dampf' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Dampf'] = []
                                dict_Systemtyp['Legende Dampf'].append(systemtyp)
                            elif rvt_gk == 'G':
                                if 'Legende Gase' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Gase'] = []
                                dict_Systemtyp['Legende Gase'].append(systemtyp)
                            elif rvt_gk == 'H':
                                if 'Legende Heizung' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Heizung'] = []
                                dict_Systemtyp['Legende Heizung'].append(systemtyp)
                            elif rvt_gk == 'K':
                                if 'Legende Kälte' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Kälte'] = []
                                dict_Systemtyp['Legende Kälte'].append(systemtyp)
                            elif rvt_gk == 'M':
                                if 'Legende LB Medien' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende LB Medien'] = []
                                dict_Systemtyp['Legende LB Medien'].append(systemtyp)
                            elif rvt_gk == 'S':
                                if 'Legende Sanitär' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Sanitär'] = []
                                dict_Systemtyp['Legende Sanitär'].append(systemtyp)
                            elif rvt_gk == 'L' or rvt_gk == 'R':
                                if 'Legende Lüftung' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende Lüftung'] = []
                                dict_Systemtyp['Legende Lüftung'].append(systemtyp)
                            else:
                                if 'Legende' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['Legende'] = []
                                dict_Systemtyp['Legende'].append(systemtyp)
                
                else:
                    if self.systemgetrennt.IsChecked:
                        dict_Systemtyp[' '] = systemtyps
                    else:
                        for systemtyp in systemtyps:
                            rvt_gk = SYSTEMTYPE[systemtyp].LookupParameter('IGF_X_Gewerkkürzel').AsString()
                            if rvt_gk == 'B':
                                if 'B' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['B'] = []
                                dict_Systemtyp['B'].append(systemtyp)
                            elif rvt_gk == 'D':
                                if 'D' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['D'] = []
                                dict_Systemtyp['D'].append(systemtyp)
                            elif rvt_gk == 'G':
                                if 'G' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['G'] = []
                                dict_Systemtyp['G'].append(systemtyp)
                            elif rvt_gk == 'H':
                                if 'H' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['H'] = []
                                dict_Systemtyp['H'].append(systemtyp)
                            elif rvt_gk == 'K':
                                if 'K' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['K'] = []
                                dict_Systemtyp['K'].append(systemtyp)
                            elif rvt_gk == 'M':
                                if 'M' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['M'] = []
                                dict_Systemtyp['M'].append(systemtyp)
                            elif rvt_gk == 'S':
                                if 'S' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['S'] = []
                                dict_Systemtyp['S'].append(systemtyp)
                            elif rvt_gk == 'L' or rvt_gk == 'R':
                                if 'L' not in dict_Systemtyp.keys():
                                    dict_Systemtyp['L'] = []
                                dict_Systemtyp['L'].append(systemtyp)
                            else:
                                if ' ' not in dict_Systemtyp.keys():
                                    dict_Systemtyp[' '] = []
                                dict_Systemtyp[' '].append(systemtyp)

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

                # Maßstab    
                _View_RvtLegend.Scale = 1    
                                
                dict_neu = {}

                for header in dict_Systemtyp.keys():
                    ueberschrift = header
                    if self.tit.IsChecked:
                        ueberschrift = self.create_header(_View_RvtLegend,header)                            
                    dict_neu[ueberschrift] = []

                    for systemtyp in dict_Systemtyp[header]:
                        line = self.modifyLine(systemtyp)
                        lineandtext = self.create_line_and_text(_View_RvtLegend,linglength,line,systemtyp,h_margin)
                        dict_neu[ueberschrift].append(lineandtext)
                
                dict_abstand = {}
                max_x = 0
                for ueberschrift in dict_neu.keys():
                    abstand = 0
                    textheight = 0
                    for n,set_line_and_text in enumerate(dict_neu[ueberschrift]):
                        box0 = set_line_and_text[0].get_BoundingBox(_View_RvtLegend)
                        box1 = set_line_and_text[1].get_BoundingBox(_View_RvtLegend)
                        max_x = max(max_x,box1.Max.X)
                        textheight = box1.Max.Y - box1.Min.Y

                        if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                        else:abstand -= (box1.Max.Y - box1.Min.Y) / 2

                        
                        if n == 0:
                            set_line_and_text[0].Location.Move(DB.XYZ(0,0-(box0.Max.Y + box0.Min.Y) / 2,0))
                            set_line_and_text[1].Location.Move(DB.XYZ(0,0-(box1.Max.Y + box1.Min.Y) / 2,0))
                            abstand -= v_margin
                        else:
                            set_line_and_text[0].Location.Move(DB.XYZ(0,abstand,0))
                            set_line_and_text[1].Location.Move(DB.XYZ(0,abstand-(box1.Max.Y + box1.Min.Y) / 2,0))
                            if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                            else:abstand -= (box1.Max.Y - box1.Min.Y) / 2
                            abstand -= v_margin

                        box0.Dispose()
                        box1.Dispose()
                    
                    if ueberschrift.GetType().Name == 'TextNote':
                        box = ueberschrift.get_BoundingBox(_View_RvtLegend)
                        height = box.Max.Y-box.Min.Y
                        y = (box.Max.Y+box.Min.Y)/2
                        max_x = max(max_x,box.Max.X)
                        ueberschrift.Location.Move(DB.XYZ(0,textheight/2+v_margin+height/2-y,0))
                        dict_abstand[ueberschrift] = abstand - height - v_margin - textheight/2
                        box.Dispose()
                    else:
                        dict_abstand[ueberschrift] = abstand  - textheight/2
                ueberschrift_temp = None
                
                abstand = 0
                max_x_header = 0
                last_text = None

                for n,ueberschrift in enumerate(dict_abstand.keys()):
                    if ueberschrift.GetType().Name == 'TextNote':
                        box = ueberschrift.get_BoundingBox(_View_RvtLegend)
                        x = (box.Max.X + box.Min.X)/2
                        ueberschrift.Location.Move(DB.XYZ(max_x/2-x,0,0))
                        doc.Regenerate()
                        max_x_header = max(box.Max.X ,max_x_header)
                        box.Dispose()
                    if n  == 0:
                        ueberschrift_temp = ueberschrift
                        for text_and_line in dict_neu[ueberschrift]:
                            last_text = text_and_line[1]
                        continue

                    abstand += dict_abstand[ueberschrift_temp]
                    ueberschrift_temp = ueberschrift

                    ueberschrift.Location.Move(DB.XYZ(0,abstand,0))
                    for text_and_line in dict_neu[ueberschrift]:
                        text_and_line[0].Location.Move(DB.XYZ(0,abstand,0))
                        text_and_line[1].Location.Move(DB.XYZ(0,abstand,0))
                        last_text = text_and_line[1]

                if self.auline.IsChecked:
                    for n,ueberschrift in enumerate(dict_abstand.keys()):
                        if n == 0:
                            self.create_aussen_kanten_y(max_x,max_x_header,last_text,ueberschrift,_View_RvtLegend,h_margin,v_margin)
                            self.create_last_aussen_kanten_x(max_x,max_x_header,_View_RvtLegend,v_margin,last_text,h_margin)
                        self.create_aussen_kanten_x(max_x,max_x_header,ueberschrift,_View_RvtLegend,h_margin,v_margin)
                
                if self.inline.IsChecked:
                    liste = dict_abstand.keys()[:]
                    for n,ueberschrift in enumerate(liste):
                        if n == len(dict_abstand)-1:
                            self.create_mittel_linie_y(linglength,h_margin,_View_RvtLegend,v_margin,ueberschrift,None,last_text)

                        else:
                            self.create_mittel_linie_y(linglength,h_margin,_View_RvtLegend,v_margin,ueberschrift,liste[n+1])
                    
                    if self.systemgetrennt.IsChecked:
                        for ueberschrift in dict_neu.keys():
                            for n,item in enumerate(dict_neu[ueberschrift]):
                                if n == len(dict_neu[ueberschrift]) - 1:
                                    continue
                                self.create_mittel_linie_x(max_x,max_x_header,_View_RvtLegend,v_margin,item[1],h_margin)



                    print('{} erstellt'.format(legendname))
        
        else:
            
            systemtyps = self.Liste_System
            dict_Systemtyp = {}
            if self.tit.IsChecked:
                if self.systemgetrennt.IsChecked:
                    dict_Systemtyp[self.ueberauswahl.Text] = systemtyps
                else:
                    for systemtyp in systemtyps:
                        rvt_gk = SYSTEMTYPE[systemtyp].LookupParameter('IGF_X_Gewerkkürzel').AsString()
                        if rvt_gk == 'B':
                            if 'Legende Sprinkler' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Sprinkler'] = []
                            dict_Systemtyp['Legende Sprinkler'].append(systemtyp)
                        elif rvt_gk == 'D':
                            if 'Legende Dampf' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Dampf'] = []
                            dict_Systemtyp['Legende Dampf'].append(systemtyp)
                        elif rvt_gk == 'G':
                            if 'Legende Gase' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Gase'] = []
                            dict_Systemtyp['Legende Gase'].append(systemtyp)
                        elif rvt_gk == 'H':
                            if 'Legende Heizung' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Heizung'] = []
                            dict_Systemtyp['Legende Heizung'].append(systemtyp)
                        elif rvt_gk == 'K':
                            if 'Legende Kälte' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Kälte'] = []
                            dict_Systemtyp['Legende Kälte'].append(systemtyp)
                        elif rvt_gk == 'M':
                            if 'Legende LB Medien' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende LB Medien'] = []
                            dict_Systemtyp['Legende LB Medien'].append(systemtyp)
                        elif rvt_gk == 'S':
                            if 'Legende Sanitär' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Sanitär'] = []
                            dict_Systemtyp['Legende Sanitär'].append(systemtyp)
                        elif rvt_gk == 'L' or rvt_gk == 'R':
                            if 'Legende Lüftung' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende Lüftung'] = []
                            dict_Systemtyp['Legende Lüftung'].append(systemtyp)
                        else:
                            if 'Legende' not in dict_Systemtyp.keys():
                                dict_Systemtyp['Legende'] = []
                            dict_Systemtyp['Legende'].append(systemtyp)
            
            else:
                if self.systemgetrennt.IsChecked:
                    dict_Systemtyp[' '] = systemtyps
                else:
                    for systemtyp in systemtyps:
                        rvt_gk = SYSTEMTYPE[systemtyp].LookupParameter('IGF_X_Gewerkkürzel').AsString()
                        if rvt_gk == 'B':
                            if 'B' not in dict_Systemtyp.keys():
                                dict_Systemtyp['B'] = []
                            dict_Systemtyp['B'].append(systemtyp)
                        elif rvt_gk == 'D':
                            if 'D' not in dict_Systemtyp.keys():
                                dict_Systemtyp['D'] = []
                            dict_Systemtyp['D'].append(systemtyp)
                        elif rvt_gk == 'G':
                            if 'G' not in dict_Systemtyp.keys():
                                dict_Systemtyp['G'] = []
                            dict_Systemtyp['G'].append(systemtyp)
                        elif rvt_gk == 'H':
                            if 'H' not in dict_Systemtyp.keys():
                                dict_Systemtyp['H'] = []
                            dict_Systemtyp['H'].append(systemtyp)
                        elif rvt_gk == 'K':
                            if 'K' not in dict_Systemtyp.keys():
                                dict_Systemtyp['K'] = []
                            dict_Systemtyp['K'].append(systemtyp)
                        elif rvt_gk == 'M':
                            if 'M' not in dict_Systemtyp.keys():
                                dict_Systemtyp['M'] = []
                            dict_Systemtyp['M'].append(systemtyp)
                        elif rvt_gk == 'S':
                            if 'S' not in dict_Systemtyp.keys():
                                dict_Systemtyp['S'] = []
                            dict_Systemtyp['S'].append(systemtyp)
                        elif rvt_gk == 'L' or rvt_gk == 'R':
                            if 'L' not in dict_Systemtyp.keys():
                                dict_Systemtyp['L'] = []
                            dict_Systemtyp['L'].append(systemtyp)
                        else:
                            if ' ' not in dict_Systemtyp.keys():
                                dict_Systemtyp[' '] = []
                            dict_Systemtyp[' '].append(systemtyp)

            if len(systemtyps) == 0:
                t.RollBack()
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

            # Maßstab    
            _View_RvtLegend.Scale = 1    
                            
            dict_neu = {}

            for header in dict_Systemtyp.keys():
                ueberschrift = header
                if self.tit.IsChecked:
                    ueberschrift = self.create_header(_View_RvtLegend,header)                            
                dict_neu[ueberschrift] = []

                for systemtyp in dict_Systemtyp[header]:
                    line = self.modifyLine(systemtyp)
                    lineandtext = self.create_line_and_text(_View_RvtLegend,linglength,line,systemtyp,h_margin)
                    dict_neu[ueberschrift].append(lineandtext)
            
            dict_abstand = {}
            max_x = 0
            for ueberschrift in dict_neu.keys():
                abstand = 0
                textheight = 0
                for n,set_line_and_text in enumerate(dict_neu[ueberschrift]):
                    box0 = set_line_and_text[0].get_BoundingBox(_View_RvtLegend)
                    box1 = set_line_and_text[1].get_BoundingBox(_View_RvtLegend)
                    max_x = max(max_x,box1.Max.X)
                    textheight = box1.Max.Y - box1.Min.Y

                    if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                    else:abstand -= (box1.Max.Y - box1.Min.Y) / 2

                    
                    if n == 0:
                        set_line_and_text[0].Location.Move(DB.XYZ(0,0-(box0.Max.Y + box0.Min.Y) / 2,0))
                        set_line_and_text[1].Location.Move(DB.XYZ(0,0-(box1.Max.Y + box1.Min.Y) / 2,0))
                        abstand -= v_margin
                    else:
                        set_line_and_text[0].Location.Move(DB.XYZ(0,abstand,0))
                        set_line_and_text[1].Location.Move(DB.XYZ(0,abstand-(box1.Max.Y + box1.Min.Y) / 2,0))
                        if (box0.Max.Y - box0.Min.Y) >= (box1.Max.Y - box1.Min.Y):abstand -= (box0.Max.Y - box0.Min.Y) / 2
                        else:abstand -= (box1.Max.Y - box1.Min.Y) / 2
                        abstand -= v_margin

                    box0.Dispose()
                    box1.Dispose()
                
                if ueberschrift.GetType().Name == 'TextNote':
                    box = ueberschrift.get_BoundingBox(_View_RvtLegend)
                    height = box.Max.Y-box.Min.Y
                    y = (box.Max.Y+box.Min.Y)/2
                    max_x = max(max_x,box.Max.X)
                    ueberschrift.Location.Move(DB.XYZ(0,textheight/2+v_margin+height/2-y,0))
                    dict_abstand[ueberschrift] = abstand - height - v_margin - textheight/2
                    box.Dispose()
                else:
                    dict_abstand[ueberschrift] = abstand  - textheight/2
            ueberschrift_temp = None
            
            abstand = 0
            max_x_header = 0
            last_text = None

            for n,ueberschrift in enumerate(dict_abstand.keys()):
                if ueberschrift.GetType().Name == 'TextNote':
                    box = ueberschrift.get_BoundingBox(_View_RvtLegend)
                    x = (box.Max.X + box.Min.X)/2
                    ueberschrift.Location.Move(DB.XYZ(max_x/2-x,0,0))
                    doc.Regenerate()
                    max_x_header = max(box.Max.X ,max_x_header)
                    box.Dispose()
                if n  == 0:
                    ueberschrift_temp = ueberschrift
                    for text_and_line in dict_neu[ueberschrift]:
                        last_text = text_and_line[1]
                    continue

                abstand += dict_abstand[ueberschrift_temp]
                ueberschrift_temp = ueberschrift

                ueberschrift.Location.Move(DB.XYZ(0,abstand,0))
                for text_and_line in dict_neu[ueberschrift]:
                    text_and_line[0].Location.Move(DB.XYZ(0,abstand,0))
                    text_and_line[1].Location.Move(DB.XYZ(0,abstand,0))
                    last_text = text_and_line[1]

            if self.auline.IsChecked:
                for n,ueberschrift in enumerate(dict_abstand.keys()):
                    if n == 0:
                        self.create_aussen_kanten_y(max_x,max_x_header,last_text,ueberschrift,_View_RvtLegend,h_margin,v_margin)
                        self.create_last_aussen_kanten_x(max_x,max_x_header,_View_RvtLegend,v_margin,last_text,h_margin)
                    self.create_aussen_kanten_x(max_x,max_x_header,ueberschrift,_View_RvtLegend,h_margin,v_margin)
            
            if self.inline.IsChecked:
                liste = dict_abstand.keys()[:]
                for n,ueberschrift in enumerate(liste):
                    if n == len(dict_abstand)-1:
                        self.create_mittel_linie_y(linglength,h_margin,_View_RvtLegend,v_margin,ueberschrift,None,last_text)

                    else:
                        self.create_mittel_linie_y(linglength,h_margin,_View_RvtLegend,v_margin,ueberschrift,liste[n+1])
                
                if self.systemgetrennt.IsChecked:
                    for ueberschrift in dict_neu.keys():
                        for n,item in enumerate(dict_neu[ueberschrift]):
                            if n == len(dict_neu[ueberschrift]) - 1:
                                continue
                            self.create_mittel_linie_x(max_x,max_x_header,_View_RvtLegend,v_margin,item[1],h_margin)



                print('{} erstellt'.format(legendname))
        
        t.Commit()
        
       
legenderstellen = LEGENDEN()

try:legenderstellen.ShowDialog()
except Exception as e:
    legenderstellen.Close()
    logger.error(e)