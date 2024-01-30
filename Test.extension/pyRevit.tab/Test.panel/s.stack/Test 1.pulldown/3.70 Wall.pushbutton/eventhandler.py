# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit

WALLS = {}
LEVELS = {}

def Get_Wandtyp():
    FamilySymbols = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsElementType().ToElements()
    for el in FamilySymbols:
        Typname = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if Typname not in WALLS.keys():
            WALLS[Typname] = el.Id

Get_Wandtyp()

def Get_LevelIds():
    FamilySymbols = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
    for el in FamilySymbols:
        Typname = el.Name
        if Typname not in LEVELS.keys():
            LEVELS[Typname] = el.Id

Get_LevelIds()


class CHANGEFAMILY(IExternalEventHandler):
    def __init__(self):
        self.class_GUI = None

        
    def Execute(self,uiapp):
        if self.class_GUI.wall.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Keinen Wandtyp ausgewählt!')
            return  
        if self.class_GUI.EU.SelectedIndex == -1:
            TaskDialog.Show('Fehler','untere Ebene nicht ausgewählt!')
            return  
        if self.class_GUI.EO.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Obene Ebene nicht ausgewählt!')
            return  

        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        eu = doc.GetElement(self.class_GUI.levels[self.class_GUI.EU.SelectedItem.ToString()])
        eo = doc.GetElement(self.class_GUI.levels[self.class_GUI.EO.SelectedItem.ToString()])
        wall =  doc.GetElement(self.class_GUI.walls[self.class_GUI.wall.SelectedItem.ToString()])

        if eo.Elevation <= eu.Elevation:
            TaskDialog.Show('Fehler','Obere Ebene liegt unter der unteren Ebene!')
            return  


        if self.class_GUI.modell.IsChecked:
            elems = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaceSeparationLines).WhereElementIsNotElementType().ToElementIds()
        elif self.class_GUI.ansicht.IsChecked:
            elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaceSeparationLines).WhereElementIsNotElementType().ToElementIds()
        elif self.class_GUI.auswahl.IsChecked:
            Auswahl = uidoc.Selection.GetElementIds()
            Liste = []
            for el in Auswahl:
                elem = doc.GetElement(el)
                try:
                    if elem.Category.Id.ToString() == '-2000831':
                        Liste.append(el)
                except:pass
            elems = Liste
        if elems.Count == 0:
            TaskDialog.Show('Info.','Keine Raumtrennung gefunden!')
            return
        # for item in elems:
        #     print(item)
        
        t = DB.Transaction(doc,'change Bauteiltyp')
        t.Start()
        for item in elems:        
            try:
            
                elem_item = doc.GetElement(item)
                curve = elem_item.GeometryCurve

                w = DB.Wall.Create(doc,curve,wall.Id,eu.Id,10,0,True,True)
                w.get_Parameter(DB.BuiltInParameter.WALL_HEIGHT_TYPE).Set(eo.Id)
                if self.class_GUI.delete.IsChecked:
                    doc.Delete(item)

            except Exception as e:
                t.RollBack()
                t.Dispose()
                print(e)
                return

        t.Commit()
        t.Dispose()
        TaskDialog.Show('Fertig','Erledigt!')

    def GetName(self):
        return "Bauteil austauschen"