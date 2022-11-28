# coding: utf8
from rpw import revit,DB,UI
from pyrevit import script, forms
import clr
from System.Collections.ObjectModel import ObservableCollection

__title__ = "Pläne anpassen"
__doc__ = """Pläne anpassen"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
config = script.get_config('Plaene_anpassen')

uidoc = revit.uidoc
doc = revit.doc
active_view = uidoc.ActiveView

planids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElementIds()


if planids.Count == 0:
    logger.error('Keine Pläne in Projekt gefunden')
    script.exit()

coll_ansichtsfenster = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ElementType))
viewport_dict = {}
for el in coll_ansichtsfenster:
    if el.FamilyName == 'Ansichtsfenster':
        name = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        viewport_dict[name] = el.Id
coll_ansichtsfenster.Dispose()

Filterplankopf = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_TitleBlocks)

class Plan(object):
    def __init__(self,elemid):
        self.checked = False
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.plannummer = self.elem.SheetNumber + ' - ' + self.elem.Name
        self.plankopf = doc.GetElement(self.elem.GetDependentElements(Filterplankopf)[0])
        self.plankopfname = self.plankopf.Name

Liste_Plan = ObservableCollection[Plan]()
Dict_Plan = {}

for planid in planids:
    plan = Plan(planid)
    if planid == active_view.Id:
        plan.checked = True
    Dict_Plan[plan.plannummer] = plan

for el in Dict_Plan.values():Liste_Plan.Add(el)

class PlanUI(forms.WPFWindow):
    def __init__(self):
        self.liste_views = Liste_Plan
        forms.WPFWindow.__init__(self, 'window.xaml')
        self.read_config()
        self.HA_Ansicht_anpassen.ItemsSource = sorted(viewport_dict.keys())
        self.LG_Ansicht_anpassen.ItemsSource = sorted(viewport_dict.keys())
        self.ListPlan.ItemsSource = Liste_Plan
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
        try:
            text_typ = self.plansuche.Text.upper()

            if text_typ in ['',None]:
                self.ListPlan.ItemsSource = self.liste_views
                text_typ = self.plansuche.Text = ''

            for item in self.liste_views:
                if item.plannummer.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
                self.ListPlan.ItemsSource = self.tempcoll
            self.ListPlan.Items.Refresh()
        except:
            self.ListPlan.ItemsSource = self.liste_views

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
            if config.HA_Ansicht_anpassen in viewport_dict.keys():
                self.HA_Ansicht_anpassen.Text = config.HA_Ansicht_anpassen
            else:
                self.HA_Ansicht_anpassen.Text = config.HA_Ansicht_anpassen = ''
        except:
            self.HA_Ansicht_anpassen.Text = config.HA_Ansicht_anpassen = ''
        
        try:
            if config.LG_Ansicht_anpassen in viewport_dict.keys():
                self.LG_Ansicht_anpassen.Text = config.LG_Ansicht_anpassen
            else:
                self.LG_Ansicht_anpassen.Text = config.LG_Ansicht_anpassen = ''
        except:
            self.LG_Ansicht_anpassen.Text = config.LG_Ansicht_anpassen = ''

    def write_config(self):
        try:
            config.bz_l_anpassen = self.bz_l_anpassen.Text
        except:
            pass

        try:
            config.bz_r_anpassen = self.bz_r_anpassen.Text
        except:
            pass

        try:
            config.bz_o_anpassen = self.bz_o_anpassen.Text
        except:
            pass

        try:
            config.bz_u_anpassen = self.bz_u_anpassen.Text
        except:
            pass

        try:
            config.pk_u_anpassen = self.pk_u_anpassen.Text
        except:
            pass

        try:
            config.pk_o_anpassen = self.pk_o_anpassen.Text
        except:
            pass

        try:
            config.pk_l_anpassen = self.pk_l_anpassen.Text
        except:
            pass

        try:
            config.pk_r_anpassen = self.pk_r_anpassen.Text
        except:
            pass
        try:
            config.raster_anpassen = self.raster_anpassen.IsChecked
        except:
            pass
        try:
            config.Haupt_anpassen = self.Haupt_anpassen.IsChecked
        except:
            pass
        try:
            config.legend_anpassen = self.legend_anpassen.IsChecked
        except:
            pass

        try:
            config.HA_Ansicht_anpassen = self.HA_Ansicht_anpassen.SelectedItem.ToString()
        except:
            config.HA_Ansicht_anpassen = self.HA_Ansicht_anpassen.Text
        try:
            config.LG_Ansicht_anpassen = self.LG_Ansicht_anpassen.SelectedItem.ToString()
        except:
            config.LG_Ansicht_anpassen = self.LG_Ansicht_anpassen.Text 

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

    def get_alle_Ansicht(self,Plan):
        plan = Plan.elem
        Viewports = plan.GetAllViewports()
        Viewport_dict = {'Grundrisse':[],'Legenden':[],'Schnitte':[]}
        for elem_id in Viewports:
            elem_viewport = doc.GetElement(elem_id)
            typ = elem_viewport.get_Parameter(DB.BuiltInParameter.VIEW_FAMILY).AsString()
            if typ in Viewport_dict.keys():
                Viewport_dict[typ].append(elem_viewport)
        return Viewport_dict
    
    def change_Beschriftungszuschnitt(self,Ansicht,plannummer):
        try:
            viewID = doc.GetElement(Ansicht.ViewId)
            cropbox = viewID.GetCropRegionShapeManager()
            cropbox.TopAnnotationCropOffset = float(config.bz_o_anpassen) / 304.8
            cropbox.BottomAnnotationCropOffset = float(config.bz_u_anpassen) / 304.8
            cropbox.RightAnnotationCropOffset = float(config.bz_r_anpassen) / 304.8
            cropbox.LeftAnnotationCropOffset = float(config.bz_l_anpassen) / 304.8
        except:logger.error('Fehler beim Ändern des Versatz von Beschriftungszuschnitt von Plan {}.'.format(plannummer))
        doc.Regenerate()
    
    def Raster_anpassen(self,Ansicht,plannummer):
        try:
            viewID = doc.GetElement(Ansicht.ViewId)
            rasters = DB.FilteredElementCollector(doc,viewID.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElementIds()
            
            box = viewID.get_BoundingBox(viewID)
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
                if not gridCurves:
                    continue
                for gridCurve in gridCurves:
                    start = gridCurve.GetEndPoint(0)
                    end = gridCurve.GetEndPoint(1)
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
                raster.Pinned = True
        except:
            logger.error('Fehler beim Anpassen der Raster der Hauptansciht von Plan {}.'.format(plannummer))
        doc.Regenerate()
    
    def change_Ansichtsfenstertype(self,Ansicht,Type,plannummer):
        try:Ansicht.ChangeTypeId(Type)
        except:logger.error('Fehler beim Ändern des Ansichtsfenstertypes der Hauptansciht/Legende von Plan {}.'.format(plannummer))

    def Ansicht_verschieben(self,Plan,Ansicht):
        try:
            x_move = Plan.plankopf.get_BoundingBox(Plan.elem).Min.X - Ansicht.get_BoundingBox(Plan.elem).Min.X + float(config.pk_l_anpassen) / 304.8 
            y_move = Plan.plankopf.get_BoundingBox(Plan.elem).Max.Y - Ansicht.get_BoundingBox(Plan.elem).Max.Y - float(config.pk_o_anpassen) / 304.8
            xyz_move = DB.XYZ(x_move,y_move,0)
            Ansicht.Location.Move(xyz_move)
        except:
            logger.error('Fehler beim Verschieben der Hauptansicht von Plan {}.'.format(Plan.plannummer))
    
    def Legend_verschieben(self,Plan,Ansicht):
        try:
            x_move = Plan.plankopf.get_BoundingBox(Plan.elem).Max.X - Ansicht.get_BoundingBox(Plan.elem).Max.X - float(config.pk_r_anpassen) / 304.8 
            y_move = Plan.plankopf.get_BoundingBox(Plan.elem).Max.Y - Ansicht.get_BoundingBox(Plan.elem).Max.Y - float(config.pk_o_anpassen) / 304.8
            xyz_move = DB.XYZ(x_move,y_move,0)
            Ansicht.Location.Move(xyz_move)
        except:
            logger.error('Fehler beim Verschieben der Legende von Plan {}.'.format(Plan.plannummer))
    
    def bearbeiten(self):
        Liste_Plaene_checked = [ele for ele in self.liste_views if ele.checked]
        if len(Liste_Plaene_checked) == 0:
            UI.TaskDialog.Show('Info','Kein Plan ausgewählt!')
            return
        t = DB.Transaction(doc, 'Raster anpassen')
        t.Start()

        with forms.ProgressBar(title='{value}/{max_value} Pläne ausgewählt',cancellable=True, step=1) as pb:
            for n, elem in enumerate(Liste_Plaene_checked):
                if pb.cancelled:
                    t.RollBack()
                    return
                pb.update_progress(n+1, len(Liste_Plaene_checked))
                Viewport_dict = self.get_alle_Ansicht(elem)
                grundrisse = Viewport_dict['Grundrisse']
                legenden = Viewport_dict['Legenden']
                schnitte = Viewport_dict['Schnitte']
                plannummer = elem.plannummer

                for grundriss in grundrisse:
                    grundriss.Pinned = False
                    if config.HA_Ansicht_anpassen:
                        self.change_Ansichtsfenstertype(grundriss,viewport_dict[config.HA_Ansicht_anpassen],plannummer)
                    self.change_Beschriftungszuschnitt(grundriss,plannummer)

                    if config.raster_anpassen:
                        self.Raster_anpassen(grundriss,plannummer)
                        
                    if config.Haupt_anpassen:
                        if len(grundrisse) == 1 and len(schnitte) == 0:
                            self.Ansicht_verschieben(elem,grundriss)

                    grundriss.Pinned = True

                for schnitt in schnitte:
                    schnitt.Pinned = False
                    if config.HA_Ansicht_anpassen:
                        self.change_Ansichtsfenstertype(schnitt,viewport_dict[config.HA_Ansicht_anpassen],plannummer)
                    self.change_Beschriftungszuschnitt(schnitt,plannummer)
                    if config.raster_anpassen:
                       self.Raster_anpassen(schnitt,plannummer)
                    if config.Haupt_anpassen:
                        if len(schnitte) == 1 and len(grundrisse) == 0:
                            self.Ansicht_verschieben(elem,schnitt)

                    schnitt.Pinned = True

                for legend in legenden:
                    legend.Pinned = False
                    if config.LG_Ansicht_anpassen:
                        self.change_Ansichtsfenstertype(legend,viewport_dict[config.HA_Ansicht_anpassen],plannummer)
                    if config.legend_anpassen:
                        if len(legenden) == 1:
                            self.Legend_verschieben(elem,legend)

                    legend.Pinned = True
                    
        t.Commit()
        t.Dispose()

   

    def ok(self,sender,args):
        self.Close()
        self.write_config()
        self.bearbeiten()
        
    def close(self,sender,args):
        self.Close()

Planfenster = PlanUI()
try:
    Planfenster.ShowDialog()
except Exception as e:
    Planfenster.Close()
    logger.error(e)