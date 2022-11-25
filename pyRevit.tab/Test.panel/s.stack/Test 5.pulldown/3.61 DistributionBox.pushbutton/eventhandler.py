# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List

class Spiegeln(IExternalEventHandler):
    def __init__(self):        
        self.Familien_Name = 'DistributionBox_'
        self.elems = None
        self.modell = True
        self.ansicht = False
        self.auswahl = False
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        param_equality = DB.FilterStringContains()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,self.Familien_Name,True)
        Filter = DB.ElementParameterFilter(Fam_name_value_rule)
        if self.modell:
            self.elems = DB.FilteredElementCollector(doc).WherePasses(Filter).WhereElementIsNotElementType().ToElementIds()
        elif self.ansicht:
            self.elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WherePasses(Filter).WhereElementIsNotElementType().ToElementIds()
        elif self.auswahl:
            Auswahl = uidoc.Selection.GetElementIds()
            Liste = []
            for el in Auswahl:
                elem = doc.GetElement(el)
                try:
                    if elem.Symbol.FamilyName.find(self.Familien_Name) != -1:
                        Liste.append(el)
                except:pass
            self.elems = Liste
        if self.elems.Count == 0:
            TaskDialog.Show('Info.','Keine DistributionBox gefunden!')
            return

        for elid in self.elems:
            el = doc.GetElement(elid)
            if el.Pinned:
                el.Pinned = False
            if el.Mirrored == False:
                continue
            t = DB.Transaction(doc,'Spiegeln')
            t.Start()
            transform = el.GetTransform()
            plane_XZ = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisX,transform.BasisZ)
            plane_XY = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisX,transform.BasisY)
            plane_YZ = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisY,transform.BasisZ)
            try:
                DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([el.Id]), plane_XZ,False)
                doc.Regenerate()
                if el.Mirrored != False:
                    t.RollBack()
                else:
                    t.Commit()
                    continue
                t.Start()
                DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([el.Id]), plane_XY,False)
                doc.Regenerate()
                if el.Mirrored != False:
                    t.RollBack()
                else:
                    t.Commit()
                    continue
                t.Start()
                DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([el.Id]), plane_YZ,False)
                doc.Regenerate()
                if el.Mirrored != False:
                    t.RollBack()
                    print("Spiegeln für DistributionBox: {} gescheitert".fromat(el.Id.ToString()))     
                else:
                    t.Commit()
                    continue
            except:
                t.RollBack()
                print("Spiegeln für DistributionBox: {} gescheitert".fromat(el.Id.ToString()))     

        TaskDialog.Show('Fertig','Fertig!')

    def GetName(self):
        return "Bauteil spiegeln"