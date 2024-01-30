# coding: utf8
from System.Collections.ObjectModel import ObservableCollection
import Autodesk.Revit.DB as DB
from IGF_ExternalEventHandler import TaskDialog,TaskDialogCommonButtons 
from System.Collections.Generic import List
from IGF_Klasse._labormedien import LaboranschlussType
from IGF_Funktionen._Parameter import get_value_Family
from IGF_Funktionen._Geometrie import LinkedPointTransform,LinkedVectorTransform
from IGF_Forms._treeview import BeschriftungInAnsicht
import System
from IGF_Filters._selectionfilter import SelectionFilterFactory
from IGF_Filters._parameterFilter import ElementParameterFilterFactory

class ExternalEventHandlerFactory_READONLY:
    
    @staticmethod
    def ElementAnzeigenIFCID(uiapp,IFC_ID):
        '''IFC ID'''
        if not IFC_ID:
            TaskDialog.Show('Fehler','keine IFC-GUID') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _filter = ElementParameterFilterFactory.createEqualFilter(DB.ElementId(DB.BuiltInParameter.IFC_GUID),IFC_ID)
        
        try:
            elemids = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ContainedInDesignOption(DB.ElementId(-1)).WherePasses(_filter).ToElementIds()
            if len(elemids) == 0:
                TaskDialog.Show('Fehler','ungültige IFC-GUID') 
                uidoc.Dispose()
                doc.Dispose()
                del uidoc
                del doc
                return
        except:
            TaskDialog.Show('Fehler','ungültige IFC-GUID') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetElementIds(elemids)
        uidoc.ShowElements(elemids)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def ElementAuswaehlenIFCID(uiapp,IFC_ID):
        '''IFC ID'''
        if not IFC_ID:
            TaskDialog.Show('Fehler','keine IFC-GUID') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _filter = ElementParameterFilterFactory.createEqualFilter(DB.ElementId(DB.BuiltInParameter.IFC_GUID),IFC_ID)
        # _filter = DB.ElementParameterFilter\
        #     (DB.FilterStringRule\
        #     (DB.ParameterValueProvider\
        #     (DB.ElementId\
        #     (DB.BuiltInParameter.IFC_GUID)),\
        #     DB.FilterStringEquals(),IFC_ID,False))
        
        try:
            elemids = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ContainedInDesignOption(DB.ElementId(-1)).WherePasses(_filter).ToElementIds()
            if len(elemids) == 0:
                TaskDialog.Show('Fehler','ungültige IFC-GUID') 
                uidoc.Dispose()
                doc.Dispose()
                del uidoc
                del doc
                return
        except:
            TaskDialog.Show('Fehler','ungültige IFC-GUID') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetElementIds(elemids)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def ElementAnzeigenElementId(uiapp,ElementId):
        '''ElementId '''
        if not ElementId:
            TaskDialog.Show('Fehler','keine ElementId') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        try:
            elem = doc.GetElement(DB.ElementId(int(ElementId)))
        except:
            TaskDialog.Show('Fehler','ungültige ElementId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return

        if not elem:
            TaskDialog.Show('Fehler','falsche ElementId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetElementIds(List[DB.ElementId]([DB.ElementId(int(ElementId))]))
        uidoc.ShowElements(elem)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def ElementAuswaehlenElementId(uiapp,ElementId):
        '''ElementId '''
        if not ElementId:
            TaskDialog.Show('Fehler','keine ElementId') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        try:
            elem = doc.GetElement(DB.ElementId(int(ElementId)))
        except:
            TaskDialog.Show('Fehler','ungültige ElementId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return

        if not elem:
            TaskDialog.Show('Fehler','falsche ElementId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetElementIds(List[DB.ElementId]([DB.ElementId(int(ElementId))]))
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def ElementAnzeigenRevitId(uiapp,RevitId):
        '''RevitId '''
        if not RevitId:
            TaskDialog.Show('Fehler','keine Revit-UniqueId') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        try:
            elem = doc.GetElement(RevitId)
        except:
            TaskDialog.Show('Fehler','ungültige Revit-UniqueId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return

        if not elem:
            TaskDialog.Show('Fehler','falsche Revit-UniqueId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetRevitIds(List[DB.RevitId]([elem.Id]))
        uidoc.ShowElements(elem)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def ElementAuswaehlenRevitId(uiapp,ElementId):
        '''RevitId '''
        if not RevitId:
            TaskDialog.Show('Fehler','keine Revit-UniqueId') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        try:
            elem = doc.GetElement(RevitId)
        except:
            TaskDialog.Show('Fehler','ungültige Revit-UniqueId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return

        if not elem:
            TaskDialog.Show('Fehler','falsche Revit-UniqueId') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetRevitIds(List[DB.RevitId]([elem.Id]))
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def PlansuchenPlannummer(uiapp,Plannummer):
        '''Plannummer '''
        if not Plannummer:
            TaskDialog.Show('Fehler','keine Plannummer') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _filter = ElementParameterFilterFactory.createEqualFilter(DB.ElementId(DB.BuiltInParameter.SHEET_NUMBER),Plannummer)
        
        # _filter = DB.ElementParameterFilter\
        #     (DB.FilterStringRule\
        #     (DB.ParameterValueProvider\
        #     (DB.ElementId\
        #     (DB.BuiltInParameter.SHEET_NUMBER)),\
        #     DB.FilterStringEquals(),Plannummer,False))
        
        try:
            elemids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ContainedInDesignOption(DB.ElementId(-1)).WherePasses(_filter).ToElementIds()
            _filter.Dispose()
            if len(elemids) == 0:
                TaskDialog.Show('Fehler','ungültige Plannummer') 
                uidoc.Dispose()
                doc.Dispose()
                del uidoc
                del doc
                return
        except:
            TaskDialog.Show('Fehler','ungültige Plannummer') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.ActiveView = doc.GetElement(elemids[0])
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc
    
    @staticmethod
    def RaumAuswaehlenRaumnummer(uiapp,Raumummer):
        '''Raumnummer '''
        if not Raumummer:
            TaskDialog.Show('Fehler','keine Raumnummer') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _filter = ElementParameterFilterFactory.createEqualFilter(DB.ElementId(DB.BuiltInParameter.ROOM_NUMBER),Raumummer)
        
        try:
            elemids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ContainedInDesignOption(DB.ElementId(-1)).WherePasses(_filter).ToElementIds()
            _filter.Dispose()
            if len(elemids) == 0:
                TaskDialog.Show('Fehler','ungültige Raumnummer') 
                uidoc.Dispose()
                doc.Dispose()
                del uidoc
                del doc
                return
        except Exception as e:
            print(e)
            TaskDialog.Show('Fehler','ungültige Raumnummer') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetElementIds(elemids)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

    @staticmethod
    def RaumanzeigenRaumnummer(uiapp,Raumummer):
        '''Raumnummer '''
        if not Raumummer:
            TaskDialog.Show('Fehler','keine Raumnummer') 
            return
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        _filter = ElementParameterFilterFactory.createEqualFilter(DB.ElementId(DB.BuiltInParameter.ROOM_NUMBER),Raumummer)
        
        try:
            elemids = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ContainedInDesignOption(DB.ElementId(-1)).WherePasses(_filter).ToElementIds()
            _filter.Dispose()
            if len(elemids) == 0:
                TaskDialog.Show('Fehler','ungültige Raumnummer') 
                uidoc.Dispose()
                doc.Dispose()
                del uidoc
                del doc
                return
        except Exception as e:
            print(e)
            TaskDialog.Show('Fehler','ungültige Raumnummer') 
            uidoc.Dispose()
            doc.Dispose()
            del uidoc
            del doc
            return
            
        uidoc.Selection.SetElementIds(elemids)
        uidoc.ShowElements(elemids)
        uidoc.Dispose()
        doc.Dispose()
        del uidoc
        del doc

