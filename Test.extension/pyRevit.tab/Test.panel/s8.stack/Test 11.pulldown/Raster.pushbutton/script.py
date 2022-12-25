# coding: utf8
from pyrevit import revit, UI, DB, script, forms
import clr
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from pyIGF_config import Server_config


__title__ = "Raster anpassen"
__doc__ = """Raster anpassen"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config_user = script.get_config('Plaene')
# server = Server_config()
# config_server = server.get_config('Plaene')
config_temp = script.get_config('Raster_anpassen')

global Start 
Start = False


uidoc = revit.uidoc
doc = revit.doc
active_view = doc.ActiveView


config = config_user


plan_coll = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_Sheets) \
    .WhereElementIsNotElementType()
planids = plan_coll.ToElementIds()
plan_coll.Dispose()


if not planids:
    logger.error('Keine Pläne in Projekt')
    script.exit()

coll_ansichtsfenster = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ElementType))
viewport = []
viewport_dict = {}
for el in coll_ansichtsfenster:
    if el.FamilyName == 'Ansichtsfenster':
        name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        viewport.append(name)
        viewport_dict[name] = el.Id
viewport.sort()
coll_ansichtsfenster.Dispose()

class Plan(object):
    def __init__(self):
        self.checked = False
        self.plannummer = ''
        self.plankopf = ''

   
    @property
    def plannummer(self):
        return self._plannummer
    @plannummer.setter
    def plannummer(self, value):
        self._plannummer = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    
    @property
    def plankopf(self):
        return self._plankopf
    @plankopf.setter
    def plankopf(self, value):
        self._plankopf = value
    
    @property
    def planid(self):
        return self._planid
    @planid.setter
    def planid(self, value):
        self._planid = value

Liste_Plan = ObservableCollection[Plan]()

Filterplankopf = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_TitleBlocks)
Auswahldict = {}

for planid in planids:
    elem = doc.GetElement(planid)
    plankopf = elem.GetDependentElements(Filterplankopf)
    if plankopf.Count != 1:
       continue
    tempclass = Plan()
    plannr = elem.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
    planna = elem.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString()
    tempclass.plannummer = plannr+' - '+planna
    tempclass.planid = planid
    if planid == active_view.Id:
        tempclass.checked = True
    tempclass.plankopf = doc.GetElement(plankopf[0])
    Liste_Plan.Add(tempclass)

class PlanUI(WPFWindow):
    def __init__(self, xaml_file_name,liste_views):
        self.liste_views = liste_views
        WPFWindow.__init__(self, xaml_file_name)
        self.read_config()
        self.HA_Ansicht_anpassen.ItemsSource = viewport
        self.LG_Ansicht_anpassen.ItemsSource = viewport
        self.ListPlan.ItemsSource = liste_views
        self.tempcoll = ObservableCollection[Plan]()

        self.plansuche.TextChanged += self.auswahl_txt_changed


    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.ListPlan.SelectedItem is not None:
            try:
                if sender.DataContext in self.ListPlan.SelectedItems:
                    for item in self.ListPlan.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.ListPlan.Items.Refresh()
                else:
                    pass
            except:
                pass
        
    def auswahl_txt_changed(self, sender, args):
        self.tempcoll.Clear()
        text_typ = self.plansuche.Text.upper()

        if text_typ in ['',None]:
            self.ListPlan.ItemsSource = self.liste_views
            text_typ = self.plansuche.Text = ''

        for item in self.liste_views:
            if item.plannummer.upper().find(text_typ) != -1:
                self.tempcoll.Add(item)
            self.ListPlan.ItemsSource = self.tempcoll
        self.ListPlan.Items.Refresh()

    
    def read_config(self):
        try:
            self.bz_l_anpassen.Text = config.bz_l_anpassen
        except:
            self.bz_l_anpassen.Text = config.bz_l_anpassen = '10'
        try:
            self.bz_r_anpassen.Text = config.bz_r_anpassen
        except:
            self.bz_r_anpassen.Text = config.bz_r_anpassen = '10'
        try:
            self.bz_o_anpassen.Text = config.bz_o_anpassen
        except:
            self.bz_o_anpassen.Text = config.bz_o_anpassen = '10'
        try:
            self.bz_u_anpassen.Text = config.bz_u_anpassen
        except:
            self.bz_u_anpassen.Text = config.bz_u_anpassen = '10'

        try:
            self.pk_l_anpassen.Text = config.pk_l_anpassen
        except:
            self.pk_l_anpassen.Text = config.pk_l_anpassen = '20'
        try:
            self.pk_r_anpassen.Text = config.pk_r_anpassen
        except:
            self.pk_r_anpassen.Text = config.pk_r_anpassen = '5'
        try:
            self.pk_o_anpassen.Text = config.pk_o_anpassen
        except:
            self.pk_o_anpassen.Text = config.pk_o_anpassen = '5'
        try:
            self.pk_u_anpassen.Text = config.pk_u_anpassen
        except:
            self.pk_u_anpassen.Text = config.pk_u_anpassen = '5'
        
        try:
            self.raster_anpassen.IsChecked = config.raster_anpassen
        except:
            self.raster_anpassen.IsChecked = config.raster_anpassen = False
        try:
            self.Haupt_anpassen.IsChecked = config.Haupt_anpassen
        except:
            self.Haupt_anpassen.IsChecked = config.Haupt_anpassen = False
        try:
            self.legend_anpassen.IsChecked = config.legend_anpassen
        except:
            self.legend_anpassen.IsChecked = config.legend_anpassen = False
    
        try:
            if config.HA_Ansicht_anpassen in viewport:
                self.HA_Ansicht_anpassen.Text = config.HA_Ansicht_anpassen
            else:
                self.HA_Ansicht_anpassen.Text = config.HA_Ansicht_anpassen = ''
        except:
            self.HA_Ansicht_anpassen.Text = config.HA_Ansicht_anpassen = ''
        
        try:
            if config.LG_Ansicht_anpassen in viewport:
                self.LG_Ansicht_anpassen.Text = config.LG_Ansicht_anpassen
            else:
                self.LG_Ansicht_anpassen.Text = config.LG_Ansicht_anpassen = ''
        except:
            self.LG_Ansicht_anpassen.Text = config.LG_Ansicht_anpassen = ''

    def write_config(self):
        try:
            config_temp.bz_l_anpassen = self.bz_l_anpassen.Text
        except:
            pass

        try:
            config_temp.bz_r_anpassen = self.bz_r_anpassen.Text
        except:
            pass

        try:
            config_temp.bz_o_anpassen = self.bz_o_anpassen.Text
        except:
            pass

        try:
            config_temp.bz_u_anpassen = self.bz_u_anpassen.Text
        except:
            pass

        try:
            config_temp.pk_u_anpassen = self.pk_u_anpassen.Text
        except:
            pass

        try:
            config_temp.pk_o_anpassen = self.pk_o_anpassen.Text
        except:
            pass

        try:
            config_temp.pk_l_anpassen = self.pk_l_anpassen.Text
        except:
            pass

        try:
            config_temp.pk_r_anpassen = self.pk_r_anpassen.Text
        except:
            pass
        try:
            config_temp.raster_anpassen = self.raster_anpassen.IsChecked
        except:
            pass
        try:
            config_temp.Haupt_anpassen = self.Haupt_anpassen.IsChecked
        except:
            pass
        try:
            config_temp.legend_anpassen = self.legend_anpassen.IsChecked
        except:
            pass

        try:
            config_temp.HA_Ansicht_anpassen = self.HA_Ansicht_anpassen.SelectedItem.ToString()
        except:
            config_temp.HA_Ansicht_anpassen = self.HA_Ansicht_anpassen.Text
        try:
            config_temp.LG_Ansicht_anpassen = self.LG_Ansicht_anpassen.SelectedItem.ToString()
        except:
            config_temp.LG_Ansicht_anpassen = self.LG_Ansicht_anpassen.Text 

        script.save_config()

    def check(self,sender,args):
        for item in self.ListPlan.Items:
            item.checked = True
        self.ListPlan.Items.Refresh()


    def uncheck(self,sender,args):
        for item in self.ListPlan.Items:
            item.checked = False
        self.ListPlan.Items.Refresh()


    def toggle(self,sender,args):
        for item in self.ListPlan.Items:
            value = item.checked
            item.checked = not value
        self.ListPlan.Items.Refresh()

    def ok(self,sender,args):
        global Start 
        Start = True
        self.write_config()
        self.Close()
            
    def close(self,sender,args):
        self.Close()

Planfenster = PlanUI("window.xaml",Liste_Plan)
try:
    Planfenster.ShowDialog()
except Exception as e:
    Planfenster.Close()
    logger.error(e)
    script.exit()

if not Start:
    script.exit()

Liste_Plaene_checked = []
for ele in Liste_Plan:
    if ele.checked:
        Liste_Plaene_checked.append(ele)

if len(Liste_Plaene_checked) == 0:
    UI.TaskDialog.Show('Info','Kein Plan ausgewählt!')
    script.exit()

t = DB.Transaction(doc, 'Raster anpassen')
t.Start()

with forms.ProgressBar(title='{value}/{max_value} Pläne ausgewählt',cancellable=True, step=1) as pb:
    for n, elem in enumerate(Liste_Plaene_checked):
        if pb.cancelled:
            t.RollBack()
            script.exit()
        pb.update_progress(n+1, len(Liste_Plaene_checked))
        plan = doc.GetElement(elem.planid)
        plankopf = elem.plankopf
        Viewport_dict = {'Grundrisse':[],'Legenden':[],'Schnitte':[]}
        Viewports = plan.GetAllViewports()
        for elem_id in Viewports:
            elem_viewport = doc.GetElement(elem_id)
            typ = elem_viewport.get_Parameter(DB.BuiltInParameter.VIEW_FAMILY).AsString()
            if typ in ['Grundrisse','Legenden','Schnitte']:
                Viewport_dict[typ].append(elem_viewport)
        grundrisse = Viewport_dict['Grundrisse']
        legenden = Viewport_dict['Legenden']
        schnitte = Viewport_dict['Schnitte']
        for grundriss in grundrisse:
            grundriss.Pinned = False
            if config_temp.HA_Ansicht_anpassen:
                try:
                   grundriss.ChangeTypeId(viewport_dict[config_temp.HA_Ansicht_anpassen])
                except:
                    logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Hauptansciht von Plan {}.'.format(elem.plannummer))
            try:
                viewID = doc.GetElement(grundriss.ViewId)
                cropbox = viewID.GetCropRegionShapeManager()
                cropbox.TopAnnotationCropOffset = float(config_temp.bz_o_anpassen) / 304.8
                cropbox.BottomAnnotationCropOffset = float(config_temp.bz_u_anpassen) / 304.8
                cropbox.RightAnnotationCropOffset = float(config_temp.bz_r_anpassen) / 304.8
                cropbox.LeftAnnotationCropOffset = float(config_temp.bz_l_anpassen) / 304.8
            except:
                logger.error('Fehler beim Ändern des Versatz von Beschriftungszuschnitt von Plan {}.'.format(elem.plannummer))
            doc.Regenerate()
            if config_temp.raster_anpassen:
                try:
                    rasters_collector = DB.FilteredElementCollector(doc,viewID.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType()
                    rasters = rasters_collector.ToElementIds()
                    rasters_collector.Dispose()
                    box = viewID.get_BoundingBox(viewID)
                    max_X = box.Max.X
                    max_Y = box.Max.Y
                    min_X = box.Min.X
                    min_Y = box.Min.Y
                    for rasid in rasters:
                        raster = doc.GetElement(rasid)
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

                            if all([newStart,newEnd]):
                                newLine = DB.Line.CreateBound( newStart, newEnd )
                            if newLine:
                                raster.SetCurveInView(DB.DatumExtentType.ViewSpecific, viewID, newLine )
                        raster.Pinned = True
                except:
                    logger.error('Fehler beim Anpassen der Raster der Hauptansciht von Plan {}.'.format(elem.plannummer))
            doc.Regenerate()
            if config_temp.Haupt_anpassen:
                if len(grundrisse) < 2:
                    try:
                        x_move = plankopf.get_BoundingBox(plan).Min.X - grundriss.get_BoundingBox(plan).Min.X + float(config_temp.pk_l_anpassen) / 304.8 
                        y_move = plankopf.get_BoundingBox(plan).Max.Y - grundriss.get_BoundingBox(plan).Max.Y - float(config_temp.pk_o_anpassen) / 304.8
                        xyz_move = DB.XYZ(x_move,y_move,0)
                        grundriss.Location.Move(xyz_move)
                    except:
                        logger.error('Fehler beim Verschieben der Hauptansicht von Plan {}.'.format(elem.plannummer))
                else:
                    logger.error('mehr als 1 Grundrisse in Plan {}, Verschieben von Grundrisse unmöglich.'.format(elem.plannummer))

            grundriss.Pinned = True

        for schnitt in schnitte:
            schnitt.Pinned = False
            if config_temp.HA_Ansicht_anpassen:
                try:
                   schnitt.ChangeTypeId(viewport_dict[config_temp.HA_Ansicht_anpassen])
                except:
                    logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Hauptansciht von Plan {}.'.format(elem.plannummer))
            try:
                viewID = doc.GetElement(schnitt.ViewId)
                cropbox = viewID.GetCropRegionShapeManager()
                cropbox.TopAnnotationCropOffset = float(config_temp.bz_o_anpassen) / 304.8
                cropbox.BottomAnnotationCropOffset = float(config_temp.bz_u_anpassen) / 304.8
                cropbox.RightAnnotationCropOffset = float(config_temp.bz_r_anpassen) / 304.8
                cropbox.LeftAnnotationCropOffset = float(config_temp.bz_l_anpassen) / 304.8
            except:
                logger.error('Fehler beim Ändern des Versatz von Beschriftungszuschnitt von Plan {}.'.format(elem.plannummer))
            doc.Regenerate()
            if config_temp.raster_anpassen:
                try:
                    rasters_collector = DB.FilteredElementCollector(doc,viewID.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType()
                    rasters = rasters_collector.ToElementIds()
                    rasters_collector.Dispose()
                    view = viewID.GetDependentElements(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Viewers))
                    box = doc.GetElement(view[0]).get_BoundingBox(viewID)
                    max_X = box.Max.X
                    max_Y = box.Max.Y
                    min_X = box.Min.X
                    min_Y = box.Min.Y
                    max_Z = box.Max.Z
                    min_Z = box.Min.Z
                    for rasid in rasters:
                        raster = doc.GetElement(rasid)
                        raster.Pinned = False
                        gridCurves = raster.GetCurvesInView(DB.DatumExtentType.ViewSpecific, viewID)
                        visible0 = raster.IsBubbleVisibleInView(DB.DatumEnds.End0,viewID)
                        visible1 = raster.IsBubbleVisibleInView(DB.DatumEnds.End1,viewID)
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

                            elif abs(Z1-Z2) > 1:
                                if Z1 > Z2:
                                    newStart = DB.XYZ(X1,Y1,max_Z)
                                    newEnd = DB.XYZ(X2,Y2,min_Z)
                                else:
                                    newStart = DB.XYZ(X1,Y1,min_Z)
                                    newEnd = DB.XYZ(X2,Y2,max_Z)
                            if all([newStart,newEnd]):
                                newLine = DB.Line.CreateBound( newStart, newEnd )
                            if newLine:
                                raster.SetCurveInView(DB.DatumExtentType.ViewSpecific, viewID, newLine )
                        if visible0:raster.ShowBubbleInView(DB.DatumEnds.End0, viewID)
                        else:raster.HideBubbleInView(DB.DatumEnds.End0, viewID)
                        if visible1:raster.ShowBubbleInView(DB.DatumEnds.End1, viewID)
                        else:raster.HideBubbleInView(DB.DatumEnds.End1, viewID)
                        raster.Pinned = True
                except:
                    logger.error('Fehler beim Anpassen der Raster der Hauptansciht von Plan {}.'.format(elem.plannummer))
            doc.Regenerate()
            if config_temp.Haupt_anpassen:
                if len(schnitte) < 2:
                    try:
                        x_move = plankopf.get_BoundingBox(plan).Min.X - schnitt.get_BoundingBox(plan).Min.X + float(config_temp.pk_l_anpassen) / 304.8 
                        y_move = plankopf.get_BoundingBox(plan).Max.Y - schnitt.get_BoundingBox(plan).Max.Y - float(config_temp.pk_o_anpassen) / 304.8
                        xyz_move = DB.XYZ(x_move,y_move,0)
                        schnitt.Location.Move(xyz_move)
                    except:
                        logger.error('Fehler beim Verschieben der Hauptansicht von Plan {}.'.format(elem.plannummer))
                else:
                    logger.error('mehr als 1 Schnitte in Plan {}, Verschieben von Schnitt unmöglich.'.format(elem.plannummer))

            schnitt.Pinned = True

        for legend in legenden:
            legend.Pinned = False
            if config_temp.LG_Ansicht_anpassen:
                try:
                    legend.ChangeTypeId(viewport_dict[config_temp.LG_Ansicht_anpassen])
                except:
                    logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Legende von Plan {}.'.format(elem.plannummer))
            if config_temp.legend_anpassen:
                if len(legenden) < 2:
                    try:
                        x_move = plankopf.get_BoundingBox(plan).Max.X - legend.get_BoundingBox(plan).Max.X - float(config_temp.pk_r_anpassen) / 304.8 
                        y_move = plankopf.get_BoundingBox(plan).Max.Y - legend.get_BoundingBox(plan).Max.Y - float(config_temp.pk_o_anpassen) / 304.8
                        xyz_move = DB.XYZ(x_move,y_move,0)
                        legend.Location.Move(xyz_move)
                    except:
                        logger.error('Fehler beim Verschieben der Hauptansicht von Plan {}.'.format(elem.plannummer))
                else:
                    logger.error('mehr als 1 Legenden in Plan {}, Verschieben von Legende unmöglich.'.format(elem.plannummer))

            legend.Pinned = True
            
t.Commit()