# coding: utf8
from pyrevit import revit, UI, DB, script, forms, HOST_APP
import clr
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from System.Text.RegularExpressions import Regex


__title__ = "Pläne anpassen"
__doc__ = """

Pläne anpassen
Legende nachhaltig ergänzen

[2022.12.24]
version: 2.0
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config_user = script.get_config('Plaene')
version = str(HOST_APP.app.VersionNumber)

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
active_view = uidoc.ActiveView

config = script.get_config(version+name+number+'HK-Familie')

Filterplankopf = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_TitleBlocks)

plans = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()

if len(plans) == 0:
    logger.error('Keine Pläne in Projekt gefunden')
    script.exit()

coll_ansichtsfenster = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ElementType)).ToElements()
viewport_dict = {}
for el in coll_ansichtsfenster:
    if el.FamilyName == 'Ansichtsfenster':
        name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        viewport_dict[name] = el.Id

coll_Plankopf = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family)).ToElements()
Plankopf_dict = {}
for el in coll_Plankopf:
    if el.FamilyCategoryId.ToString() == '-2000280':
        symids = el.GetFamilySymbolIds()
        for Id in symids:
            name = doc.GetElement(Id).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
            Plankopf_dict[name] = Id

views = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElementIds()
Legenden_dict_Name  = {}

for el in views:
    elem = doc.GetElement(el)
    typ = elem.ViewType.ToString()
    if elem.IsTemplate:
        continue
    if typ == 'Legend':
        Legenden_dict_Name[elem.Name] = el

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

class Plan(TemplateItemBase):
    def __init__(self,elem):
        TemplateItemBase.__init__(self)
        self._checked = False
        self._checked1 = False
        self.WPF = None
        self._legendeselectedindex = -1
        self._Plankopfindex_neu = -1
        self.elem = elem
        self.plannr = self.elem.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
        self.planna = self.elem.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString()
        self.plannummer = self.plannr + ' - ' + self.planna
        self.Plankopf = ''
        self.lenegdeDict = Legenden_dict_Name
        self.plankopfDict = Plankopf_dict
        self.legendeitemssource = sorted(self.lenegdeDict.keys())
        self.plankopfitemssource = sorted(self.plankopfDict.keys())
        self.Revit_Plankopf = None
        self.Liste_Grundriss = []
        self.Liste_Schnitte = []
        self.Liste_Legende = []
        self.get_Plankopf()
        self.get_Viewports()
    
    def add_Legende(self):
        if self.legendeselectedindex != -1:
            location = DB.XYZ((self.elem.Outline.Max.U+self.elem.Outline.Min.U)/2,(self.elem.Outline.Max.V+self.elem.Outline.Min.V)/2,0)
            
            try:
                legende = DB.Viewport.Create(doc,self.elem.Id,self.lenegdeDict[self.legendeitemssource[self.legendeselectedindex]],location)
                self.Liste_Legende.append(legende)
            except:
                logger.error('Fehler beim Hinzufügen der Legende von Plan {}.'.format(self.plannummer))
          
    def change_plankopf(self):
        if self.Plankopfindex_neu != -1:
            try:self.Revit_Plankopf.ChangeTypeId(self.plankopfDict[self.plankopfitemssource[self.Plankopfindex_neu]])
            except:logger.error('Fehler beim Ändern des Plankopfs von Plan {}.'.format(self.plannummer))
    
    def changeviewtype(self):
        if self.WPF.HA_Ansicht_anpassen.SelectedIndex != -1:
            for grundriss in self.Liste_Grundriss:
                try:grundriss.ChangeTypeId(viewport_dict[self.WPF.HA_Ansicht_anpassen.SelectedItem])
                except:logger.error('Fehler beim Ändern des Ansichtsfenstertypes des Grundrisses von Plan {}.'.format(self.plannummer))
            for schnitt in self.Liste_Schnitte:
                try:schnitt.ChangeTypeId(viewport_dict[self.WPF.HA_Ansicht_anpassen.SelectedItem])
                except:logger.error('Fehler beim Ändern des Ansichtsfenstertypes des Schnittes von Plan {}.'.format(self.plannummer))
        if self.WPF.LG_Ansicht_anpassen.SelectedIndex != -1:
            for legend in self.Liste_Legende:
                try:legend.ChangeTypeId(viewport_dict[self.WPF.LG_Ansicht_anpassen.SelectedItem])
                except:logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Legende von Plan {}.'.format(self.plannummer))
            
    def changebeschriftungszuschnitt(self):
        def changeviewbeschriftungszuschnitt(view):
            viewID = doc.GetElement(view.ViewId)
            cropbox = viewID.GetCropRegionShapeManager()
            cropbox.TopAnnotationCropOffset = float(self.WPF.bz_o_anpassen.Text) / 304.8
            cropbox.BottomAnnotationCropOffset = float(self.WPF.bz_u_anpassen.Text) / 304.8
            cropbox.RightAnnotationCropOffset = float(self.WPF.bz_r_anpassen.Text) / 304.8
            cropbox.LeftAnnotationCropOffset = float(self.WPF.bz_l_anpassen.Text) / 304.8
        try:
            for grundriss in self.Liste_Grundriss:
                try:changeviewbeschriftungszuschnitt(grundriss)
                except:logger.error('Fehler beim Ändern des Versatz von Beschriftungszuschnitt des Grundrisses von Plan {}.'.format(self.plannummer))
            for schnitt in self.Liste_Schnitte:
                try:changeviewbeschriftungszuschnitt(schnitt)
                except:logger.error('Fehler beim Ändern des Versatz von Beschriftungszuschnitt des Schnittes von Plan {}.'.format(self.plannummer))
        except:pass

    def plan_raster_anpassen(self):
        def raster_view_anpassen(view):
            viewID = doc.GetElement(view.ViewId)
            rasters = DB.FilteredElementCollector(doc,view.ViewId).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
            box = viewID.get_BoundingBox(viewID)
            max_X = box.Max.X
            max_Y = box.Max.Y
            max_Z = box.Max.Z
            min_X = box.Min.X
            min_Y = box.Min.Y
            min_Z = box.Min.Z
            for raster in rasters:
                raster.Pinned = False
                gridCurves = raster.GetCurvesInView(DB.DatumExtentType.ViewSpecific, viewID)
                if not gridCurves:
                    continue
                for gridCurve in gridCurves:
                    start = gridCurve.GetEndPoint( 0 )
                    end = gridCurve.GetEndPoint( 1 )
                    X1 = start.X
                    Y1 = start.Y
                    Z1 = start.Z
                    X2 = end.X
                    Y2 = end.Y
                    Z2 = end.Z
                    newStart = None
                    newEnd = None
                    newLine = None
                    if abs(X1-X2) > 1:
                        if X1>X2:
                            newStart = DB.XYZ(max_X,Y1,Z1)
                            newEnd = DB.XYZ(min_X,Y2,Z2)
                        else:
                            newStart = DB.XYZ(min_X,Y1,Z1)
                            newEnd = DB.XYZ(max_X,Y2,Z2)
                    elif abs(Y1-Y2) > 1:
                        if Y1>Y2:
                            newStart = DB.XYZ(X1,max_Y,Z1)
                            newEnd = DB.XYZ(X2,min_Y,Z2)
                        else:
                            newStart = DB.XYZ(X1,min_Y,Z1)
                            newEnd = DB.XYZ(X2,max_Y,Z2)
                    else:
                        if Z1>Z2:
                            newStart = DB.XYZ(X1,Y1,max_Z)
                            newEnd = DB.XYZ(X2,Y2,min_Z)
                        else:
                            newStart = DB.XYZ(X1,Y1,min_Z)
                            newEnd = DB.XYZ(X2,Y2,max_Z)

                    if all([newStart,newEnd]):
                        newLine = DB.Line.CreateBound( newStart, newEnd )
                    if newLine:
                        raster.SetCurveInView(DB.DatumExtentType.ViewSpecific, viewID, newLine )
                raster.Pinned = True
        if self.WPF.raster_anpassen.IsChecked:
            for grundriss in self.Liste_Grundriss:
                try:raster_view_anpassen(grundriss)
                except:logger.error('Fehler beim Anpassen der Raster des Grundrisses von Plan {}.'.format(self.plannummer))
            for schnitt in self.Liste_Schnitte:
                try:raster_view_anpassen(schnitt)
                except:logger.error('Fehler beim Anpassen der Raster des Schnittes von Plan {}.'.format(self.plannummer))

    def get_Plankopf(self):
        Liste = self.elem.GetDependentElements(Filterplankopf)
        if len(Liste) == 1:
            self.Revit_Plankopf = doc.GetElement(Liste[0])
            self.Plankopf = self.Revit_Plankopf.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()    
    
    def get_Viewports(self):
        Viewports = self.elem.GetAllViewports()
        for elem_id in Viewports:
            elem_viewport = doc.GetElement(elem_id)
            typ = elem_viewport.get_Parameter(DB.BuiltInParameter.VIEW_FAMILY).AsString()
            if typ == 'Grundrisse':
                self.Liste_Grundriss.append(elem_viewport)
            elif typ == 'Legenden':
                self.Liste_Legende.append(elem_viewport)
            elif typ == 'Schnitte':
                self.Liste_Schnitte.append(elem_viewport)

    def move_views(self):
        if self.WPF.Haupt_anpassen.IsChecked:
            if len(self.Liste_Grundriss) + len(self.Liste_Schnitte) < 2:
                if self.WPF.linksoben.IsChecked:
                    for grundriss in self.Liste_Grundriss:
                        grundriss.Pinned = False
                        try:
                            x_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Min.X - grundriss.get_BoundingBox(self.elem).Min.X + float(self.WPF.pk_l_anpassen.Text) / 304.8 
                            y_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Max.Y - grundriss.get_BoundingBox(self.elem).Max.Y - float(self.WPF.pk_o_anpassen.Text) / 304.8
                            xyz_move = DB.XYZ(x_move,y_move,0)
                            grundriss.Location.Move(xyz_move)
                        except:
                            logger.error('Fehler beim Verschieben des Grudnrisses von Plan {}.'.format(self.plannummer))
                        grundriss.Pinned = True
                    for schnitt in self.Liste_Schnitte:
                        schnitt.Pinned = False
                        try:
                            x_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Min.X - schnitt.get_BoundingBox(self.elem).Min.X + float(self.WPF.pk_l_anpassen.Text) / 304.8 
                            y_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Max.Y - schnitt.get_BoundingBox(self.elem).Max.Y - float(self.WPF.pk_o_anpassen.Text) / 304.8
                            xyz_move = DB.XYZ(x_move,y_move,0)
                            schnitt.Location.Move(xyz_move)
                        except:
                            logger.error('Fehler beim Verschieben des Schnittes von Plan {}.'.format(self.plannummer))
                        schnitt.Pinned = True
                else:
                    for grundriss in self.Liste_Grundriss:
                        grundriss.Pinned = False
                        try:
                            x_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Min.X - grundriss.get_BoundingBox(self.elem).Min.X + float(self.WPF.pk_l_anpassen.Text) / 304.8 
                            y_move = (self.Revit_Plankopf.get_BoundingBox(self.elem).Max.Y + self.Revit_Plankopf.get_BoundingBox(self.elem).Min.Y) / 2.0 - (grundriss.get_BoundingBox(self.elem).Max.Y + grundriss.get_BoundingBox(self.elem).Min.Y) / 2.0
                            xyz_move = DB.XYZ(x_move,y_move,0)
                            grundriss.Location.Move(xyz_move)
                        except:
                            logger.error('Fehler beim Verschieben des Grudnrisses von Plan {}.'.format(self.plannummer))
                        grundriss.Pinned = True
                    for schnitt in self.Liste_Schnitte:
                        schnitt.Pinned = False
                        try:
                            x_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Min.X - schnitt.get_BoundingBox(self.elem).Min.X + float(self.WPF.pk_l_anpassen.Text) / 304.8 
                            y_move = (self.Revit_Plankopf.get_BoundingBox(self.elem).Max.Y + self.Revit_Plankopf.get_BoundingBox(self.elem).Min.Y) / 2.0 - (schnitt.get_BoundingBox(self.elem).Max.Y + schnitt.get_BoundingBox(self.elem).Min.Y) / 2.0
                            xyz_move = DB.XYZ(x_move,y_move,0)
                            schnitt.Location.Move(xyz_move)
                        except:
                            logger.error('Fehler beim Verschieben des Schnittes von Plan {}.'.format(self.plannummer))
                        schnitt.Pinned = True


        if self.WPF.legend_anpassen.IsChecked:
            if len(self.Liste_Legende) < 2:
                for legende in self.Liste_Legende:
                    legende.Pinned = False
                    try:
                        x_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Max.X - legende.get_BoundingBox(self.elem).Max.X - float(self.WPF.pk_r_anpassen.Text) / 304.8 
                        y_move = self.Revit_Plankopf.get_BoundingBox(self.elem).Max.Y - legende.get_BoundingBox(self.elem).Max.Y - float(self.WPF.pk_o_anpassen.Text) / 304.8
                        xyz_move = DB.XYZ(x_move,y_move,0)
                        legende.Location.Move(xyz_move)
                    except:
                        logger.error('Fehler beim Verschieben der Legende von Plan {}.'.format(self.plannummer))
                    legende.Pinned = True
            else:
                logger.error('{} Legenden in Plan {}, Verschieben von Legende unmöglich.'.format(len(self.Liste_Legende),self.plannummer))
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    @property
    def checked1(self):
        return self._checked1
    @checked1.setter
    def checked1(self,value):
        if value != self._checked1:
            self._checked1 = value
            self.RaisePropertyChanged('checked1')
    @property
    def legendeselectedindex(self):
        return self._legendeselectedindex
    @legendeselectedindex.setter
    def legendeselectedindex(self,value):
        if value != self._legendeselectedindex:
            self._legendeselectedindex = value
            self.RaisePropertyChanged('legendeselectedindex')
    @property
    def Plankopfindex_neu(self):
        return self._Plankopfindex_neu
    @Plankopfindex_neu.setter
    def Plankopfindex_neu(self,value):
        if value != self._Plankopfindex_neu:
            self._Plankopfindex_neu = value
            self.RaisePropertyChanged('Plankopfindex_neu')

Liste_Plan = ObservableCollection[Plan]()
Liste_Plan1 = ObservableCollection[Plan]()

for elem in plans:
    plan = Plan(elem)
    if not plan.Revit_Plankopf:
        continue
    
    if plan.elem.Id.IntegerValue == active_view.Id.IntegerValue:
        plan.checked = True
        plan.checked1 = True

    Liste_Plan.Add(plan)
    if len(plan.Liste_Legende) == 0:
        Liste_Plan1.Add(plan)

class PlanUI(WPFWindow):
    def __init__(self):
        self.Liste_Plan = Liste_Plan
        self.Liste_Plan1 = Liste_Plan1
        self.regex2 = Regex("[^0-9]+")
        WPFWindow.__init__(self, 'window.xaml',handle_esc=False)
        self.read_config()
        self.HA_Ansicht_anpassen.ItemsSource = sorted(viewport_dict.keys())
        self.LG_Ansicht_anpassen.ItemsSource = sorted(viewport_dict.keys())
        self.ListPlan.ItemsSource = self.Liste_Plan
        self.ListPlan_legende.ItemsSource = self.Liste_Plan1
        self.tempcoll = ObservableCollection[Plan]()
        self.tempcoll1 = ObservableCollection[Plan]()
        self.start = False

    def textinput(self, sender, args):
        try:
            args.Handled = self.regex2.IsMatch(args.Text)
        except:
            args.Handled = True

    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.ListPlan.SelectedItem is not None:
            try:
                if sender.DataContext in self.ListPlan.SelectedItems:
                    for item in self.ListPlan.SelectedItems:
                        try:item.checked = Checked
                        except:pass
                else:pass
            except:pass
    
    def checkedchanged1(self, sender, args):
        Checked = sender.IsChecked
        if self.ListPlan_legende.SelectedItem is not None:
            try:
                if sender.DataContext in self.ListPlan_legende.SelectedItems:
                    for item in self.ListPlan_legende.SelectedItems:
                        try:item.checked1 = Checked
                        except:pass
                else:pass
            except:pass

    def textchanged(self,sender,Liste0,Liste1,WPFitem):
        text = sender.Text
        if not text:
            WPFitem.ItemsSource = Liste0
            return
        else:
            Liste1.Clear()
            for elem in Liste0:
                if elem.plannummer.upper().find(text.upper()) != -1:
                    Liste1.Add(elem)
            WPFitem.ItemsSource = Liste1
            return
    
    def plankopfchanged(self,sender,e):
        sindex = sender.SelectedIndex
        if self.ListPlan.SelectedItem is not None:
            try:
                if sender.DataContext in self.ListPlan.SelectedItems:
                    for item in self.ListPlan.SelectedItems:
                        try:item.Plankopfindex_neu = sindex
                        except:pass
                else:pass
            except:pass
    
    def legendechanged(self,sender,e):
        sindex = sender.SelectedIndex
        if self.ListPlan_legende.SelectedItem is not None:
            try:
                if sender.DataContext in self.ListPlan_legende.SelectedItems:
                    for item in self.ListPlan_legende.SelectedItems:
                        try:item.legendeselectedindex = sindex
                        except:pass
                else:pass
            except:pass

    
    def suche_plan_changed(self,sender,e):
        self.textchanged(sender,self.Liste_Plan,self.tempcoll,self.ListPlan)
    
    def suche_legende_changed(self,sender,e):
        self.textchanged(sender,self.Liste_Plan1,self.tempcoll1,self.ListPlan_legende)
    
    def read_config(self):
        try:self.bz_l_anpassen.Text = config.bz_l_anpassen
        except:pass
        try:self.bz_r_anpassen.Text = config.bz_r_anpassen
        except:pass
        try:self.bz_o_anpassen.Text = config.bz_o_anpassen
        except:pass
        try:self.bz_u_anpassen.Text = config.bz_u_anpassen
        except:pass
        try:self.pk_l_anpassen.Text = config.pk_l_anpassen
        except:pass
        try:self.pk_r_anpassen.Text = config.pk_r_anpassen
        except:pass
        try:self.pk_o_anpassen.Text = config.pk_o_anpassen
        except:pass
        try:self.pk_u_anpassen.Text = config.pk_u_anpassen
        except:pass
        
        try:self.raster_anpassen.IsChecked = config.raster_anpassen
        except:pass
        try:self.Haupt_anpassen.IsChecked = config.Haupt_anpassen
        except:pass
        try:self.legend_anpassen.IsChecked = config.legend_anpassen
        except:pass
        try:self.linkscenter.IsChecked = config.linkscenter
        except:pass
        try:self.linksoben.IsChecked = config.linksoben
        except:pass

        try:
            if config.HA_Ansicht_anpassen in viewport_dict.keys():
                self.HA_Ansicht_anpassen.SelectedItem = config.HA_Ansicht_anpassen
            else:pass
        except:pass
        
        try:
            if config.LG_Ansicht_anpassen in viewport_dict.keys():
                self.LG_Ansicht_anpassen.SelectedItem = config.LG_Ansicht_anpassen
            else:pass
        except:pass

    def write_config(self):
        try:config.bz_l_anpassen = self.bz_l_anpassen.Text
        except:pass
        try:config.bz_r_anpassen = self.bz_r_anpassen.Text
        except:pass
        try:config.bz_o_anpassen = self.bz_o_anpassen.Text
        except:pass
        try:config.bz_u_anpassen = self.bz_u_anpassen.Text
        except:pass
        try:config.pk_u_anpassen = self.pk_u_anpassen.Text
        except:pass
        try:config.pk_o_anpassen = self.pk_o_anpassen.Text
        except:pass
        try:config.pk_l_anpassen = self.pk_l_anpassen.Text
        except:pass
        try:config.pk_r_anpassen = self.pk_r_anpassen.Text
        except:pass
        try:config.raster_anpassen = self.raster_anpassen.IsChecked
        except:pass
        try:config.Haupt_anpassen = self.Haupt_anpassen.IsChecked
        except:pass
        try:config.legend_anpassen = self.legend_anpassen.IsChecked
        except:pass
        try:config.HA_Ansicht_anpassen = self.HA_Ansicht_anpassen.SelectedItem
        except:pass
        try:config.LG_Ansicht_anpassen = self.LG_Ansicht_anpassen.SelectedItem
        except:pass 

        script.save_config()

    def ok(self,sender,args):
        self.start = True
        self.write_config()
        self.Close()
            
    def close(self,sender,args):
        self.Close()

    def hauptcheckedchanged(self,sender,e):
        checked = sender.IsChecked
        if checked:
            self.linkscenter.IsEnabled = True
            self.linksoben.IsEnabled = True
        else:
            self.linkscenter.IsEnabled = False
            self.linksoben.IsEnabled = False
            
Planfenster = PlanUI()
try:
    Planfenster.ShowDialog()
except Exception as e:
    Planfenster.Close()
    logger.error(e)
    script.exit()

if not Planfenster.start:
    script.exit()

Liste_Plan_Neu = []
for el in Liste_Plan:
    if el.checked or el.checked1:
        Liste_Plan_Neu.append(el)
        el.WPF = Planfenster

if len(Liste_Plan_Neu) == 0:
    UI.TaskDialog.Show('Info','Kein Plan ausgewählt!')
    script.exit()

t = DB.Transaction(doc, 'Pläne anpassen')
t.Start()

with forms.ProgressBar(title='{value}/{max_value} Pläne ausgewählt',cancellable=True, step=1) as pb:
    for n, plan in enumerate(Liste_Plan_Neu):
        if pb.cancelled:
            t.RollBack()
            script.exit()
        pb.update_progress(n+1, len(Liste_Plan_Neu))
        if plan.checked:
            plan.change_plankopf()
            doc.Regenerate()
        if plan.checked1:
            plan.add_Legende()
            doc.Regenerate()

        plan.changeviewtype()
        doc.Regenerate()

        plan.changebeschriftungszuschnitt()
        doc.Regenerate()

        plan.plan_raster_anpassen()
        doc.Regenerate()

        plan.move_views()
        doc.Regenerate()
t.Commit()
