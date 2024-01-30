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
from IGF_Extensions._super_pickElements import SelectionFilterFactory

def ElementAnzeigenIFCID(uiapp,IFC_ID):
    '''IFC ID'''
    if not IFC_ID:
        TaskDialog.Show('Fehler','keine IFC-GUID') 
        return
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _filter = DB.ElementParameterFilter\
        (DB.FilterStringRule\
        (DB.ParameterValueProvider\
        (DB.ElementId\
        (DB.BuiltInParameter.IFC_GUID)),\
        DB.FilterStringEquals(),IFC_ID,False))
    
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

def ElementAuswaehlenIFCID(uiapp,IFC_ID):
    '''IFC ID'''
    if not IFC_ID:
        TaskDialog.Show('Fehler','keine IFC-GUID') 
        return
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _filter = DB.ElementParameterFilter\
        (DB.FilterStringRule\
        (DB.ParameterValueProvider\
        (DB.ElementId\
        (DB.BuiltInParameter.IFC_GUID)),\
        DB.FilterStringEquals(),IFC_ID,False))
    
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

def PlansuchenPlannummer(uiapp,Plannummer):
    '''Plannummer '''
    if not Plannummer:
        TaskDialog.Show('Fehler','keine Plannummer') 
        return
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _filter = DB.ElementParameterFilter\
        (DB.FilterStringRule\
        (DB.ParameterValueProvider\
        (DB.ElementId\
        (DB.BuiltInParameter.SHEET_NUMBER)),\
        DB.FilterStringEquals(),Plannummer,False))
    
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

def RaumAuswaehlenRaumnummer(uiapp,Raumummer):
    '''Raumnummer '''
    if not Raumummer:
        TaskDialog.Show('Fehler','keine Raumnummer') 
        return
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _filter = DB.ElementParameterFilter\
        (DB.FilterStringRule\
        (DB.ParameterValueProvider\
        (DB.ElementId\
        (DB.BuiltInParameter.ROOM_NUMBER)),\
        DB.FilterStringEquals(),Raumummer,False))
    
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
    except:
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

def RaumanzeigenRaumnummer(uiapp,Raumummer):
    '''Raumnummer '''
    if not Raumummer:
        TaskDialog.Show('Fehler','keine Raumnummer') 
        return
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _filter = DB.ElementParameterFilter\
        (DB.FilterStringRule\
        (DB.ParameterValueProvider\
        (DB.ElementId\
        (DB.BuiltInParameter.ROOM_NUMBER)),\
        DB.FilterStringEquals(),Raumummer,False))
    
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
    except:
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

def FamilyParameterErmitteln(uiapp):
    doc = uiapp.ActiveUIDocument.Document
    if not doc.IsFamilyDocument:
        doc.Dispose()
        return
    Parameters = doc.FamilyManager.Parameters
    _dict = {param.Definition.Name:param for param in Parameters}
    Parameters.Dispose()
    doc.Dispose()
    del doc
    # del Parameters
    return _dict

def FamilyParameterwechseln(uiapp,alt_parameter,neu_parameter):
    doc = uiapp.ActiveUIDocument.Document
    familymanager = doc.FamilyManager
    t = DB.Transaction(doc,'Parameter ändern')
    t.Start()
    if alt_parameter.IsDeterminedByFormula:
        familymanager.SetFormula(neu_parameter,alt_parameter.Formula)
    else:
        for familytyp in familymanager.Types:
            familymanager.CurrentType  = familytyp
            doc.Regenerate()
            altvwert = get_value_Family(familytyp, alt_parameter)
            if altvwert is not None:
                storagetype = neu_parameter.StorageType.ToString()
                if storagetype == 'String':
                    try:
                        familymanager.Set(neu_parameter,str(altvwert))
                    except Exception as e:
                        print('String',e.ToString())
                elif storagetype == 'Integer':
                    try:
                        familymanager.Set(neu_parameter,int(altvwert))
                    except Exception as e:
                        print('Integer',e.ToString())
                elif storagetype == 'Double':
                    try:
                        familymanager.SetValueString(neu_parameter,str(altvwert))
                    except Exception as e:
                        print('Double',e.ToString())
                elif storagetype == 'ElementId':
                    try:
                        familymanager.SetValueString(neu_parameter,str(altvwert))
                    except Exception as e:
                        print('ElementId',e.ToString())
                    

    for assoparam in alt_parameter.AssociatedParameters:
        familymanager.AssociateElementParameterToFamilyParameter(assoparam,neu_parameter)
    params = familymanager.Parameters
    fehler = False
    for param in params:
        if param not in [alt_parameter,neu_parameter]:
          
            formula = param.Formula
            if formula:
                try:
                    
                    if formula.find(alt_parameter.Definition.Name) != -1:
                        neuformula = formula.replace(alt_parameter.Definition.Name,neu_parameter.Definition.Name)
                        familymanager.SetFormula(param,neuformula)
                except:
                    fehler = True
                    print('Änderung der Formel von Parameter {} fehlschlagen.'.format(param.Definition.Name))
    dimensions = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Dimensions).ToElements()
    for dimension in dimensions:
        if dimension.FamilyLabel and dimension.FamilyLabel.Definition.Name == alt_parameter.Definition.Name:
            try:dimension.FamilyLabel = neu_parameter
            except Exception as e:
                print(e)
                fehler = True
    if not fehler:
        result = TaskDialog.Show('Abfrage','den alten Parameter löschen?',TaskDialogCommonButtons.Yes|TaskDialogCommonButtons.No) 
        if result.ToString() == 'Yes':
            try:familymanager.RemoveParameter(alt_parameter)
            except Exception as e:print(e)

    t.Commit()
    t.Dispose()
    doc.Dispose()
    familymanager.Dispose()
    params.Dispose()
    del t
    del doc
    del familymanager
    del params
    TaskDialog.Show('Info','Parameter erfolgreich geändert!')

def Get_Beschriftung_In_Ansicht(uiapp):
    doc = uiapp.ActiveUIDocument.Document
    view = uiapp.ActiveUIDocument.ActiveView
    liste = List[System.Type]()
    liste.Add(DB.IndependentTag)
    liste.Add(DB.SpatialElementTag)
    TagsFilter = DB.ElementMulticlassFilter(liste)

    AllTags = {}
    coll_Ansicht = DB.FilteredElementCollector(doc,view.Id).WhereElementIsNotElementType().WherePasses(TagsFilter).ToElements()

    for elem in coll_Ansicht:
        cate = elem.Category.Name
        name = elem.Name
        fam = doc.GetElement(elem.GetTypeId()).FamilyName
        if cate not in AllTags.keys():
            AllTags[cate] = {}
        if fam not in AllTags[cate].keys():
            AllTags[cate][fam] = {}
        if name not in AllTags[cate][fam].keys():
            AllTags[cate][fam][name] = [List[DB.ElementId](),List[DB.ElementId]()]
    
        AllTags[cate][fam][name][0].Add(elem.Id)

    coll_doc = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(TagsFilter).ToElements()
    for elem in coll_doc:
        if not elem.IsHidden(view):
            continue
        cate = elem.Category.Name
        name = elem.Name
        fam = doc.GetElement(elem.GetTypeId()).FamilyName
        if cate not in AllTags.keys():
            AllTags[cate] = {}
        if fam not in AllTags[cate].keys():
            AllTags[cate][fam] = {}
        if name not in AllTags[cate][fam].keys():
            AllTags[cate][fam][name] = [List[DB.ElementId](),List[DB.ElementId]()]
        
        AllTags[cate][fam][name][1].Add(elem.Id)

    TagsFilter.Dispose()
    del TagsFilter

    doc.Dispose()
    del doc

    view.Dispose()
    del view

    ItemsSource = ObservableCollection[BeschriftungInAnsicht]()

    for cate in sorted(AllTags.keys()):
        kate = BeschriftungInAnsicht(cate)
        ItemsSource.Add(kate)
        for fam in sorted(AllTags[cate].keys()):
            Fam = BeschriftungInAnsicht(fam)
            Fam.parent = kate
            kate.children.Add(Fam)
            for typ in sorted(AllTags[cate][fam].keys()):
                Typ = BeschriftungInAnsicht(typ)
                Typ.parent = Fam
                Fam.children.Add(Typ)
                Typ.ListeShow = AllTags[cate][fam][typ][0]
                Typ.ListeHidden = AllTags[cate][fam][typ][1]
                if Typ.ListeShow.Count == 0:
                    Typ.checked = False
                elif Typ.ListeHidden.Count == 0:
                    Typ.checked = True
                else:
                    Typ.checked = None
                Typ.checkedchanged()
            Fam.checkedchanged()
    
    return ItemsSource

def HiddenUnhidden_Element_In_Ansicht(uiapp,elems_Hidden = [],elems_Unhidden = []):
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    view = uidoc.ActiveView
    t = DB.Transaction(doc,'Ein-/Ausblenden')
    t.Start()

    try:
        if len(elems_Hidden) > 0:view.HideElements(elems_Hidden)
    except Exception as e:print(e)
    try:
        if len(elems_Unhidden) > 0:view.UnhideElements(elems_Unhidden)
    except Exception as e:print(e)

    t.Commit()
    t.Dispose()
    del t 
    doc.Dispose()
    del doc
    uidoc.Dispose()
    del uidoc
    view.Dispose()
    del view
    
    TaskDialog.Show('Info','Erledigt!') 

def SelectRaumPickElementsByRectangle(uiapp):
    uidoc = uiapp.ActiveUIDocument
    filters = SelectionFilterFactory.CreateElementSelectionFilter(lambda x: x.Category.Name == 'MEP-Räume')
    elems = uidoc.Selection.PickElementsByRectangle(filters,'MEP-Räume auswählen')
    uidoc.Dispose()
    return List[DB.ElementId]([elem.Id for elem in elems])

def Get_Selected_MEPRaum(uiapp):
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    elems = [doc.GetElement(e) for e in uidoc.Selection.GetElementIds() if doc.GetElement(e).Category.Id.ToString() == '-2003600']
    uidoc.Dspose()
    doc.Dispose()
    return elems

def Laboranschluss_anpassen(uiapp,Liste_Familie,Liste_Parameter):
    """param: LaboranschlussItemBase,
       Liste_Familie: IGF_Klasse.labormedien.LaboranschlussType
    """
    doc = uiapp.ActiveUIDocument.Document
    for Param in Liste_Parameter:
        Param.Liste_type = []
        if Param.durchmesser:
            if Param.art == 'LAB':
                for Familie in Liste_Familie:
                    if Familie.Art == 'Abluft' and Familie.Shape == 'RU' and Familie.Ist24h == 'not 24h':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        break
            elif Param.art == 'LZU':
                for Familie in Liste_Familie:
                    if Familie.Art == 'Zuluft' and Familie.Shape == 'RU':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        break
            elif Param.art == '24h':
                for Familie in Liste_Familie:
                    if Familie.Art == 'Abluft' and Familie.Shape == 'RU' and Familie.Ist24h == '24h':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        break
            else:
                for Familie in Liste_Familie:
                    if Familie.Shape == 'RU':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        
        else:
            if Param.art == 'LAB':
                for Familie in Liste_Familie:
                    if Familie.Art == 'Abluft' and Familie.Shape == 'RE' and Familie.Ist24h == 'not 24h':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        break
            elif Param.art == '24h':
                for Familie in Liste_Familie:
                    if Familie.Art == 'Abluft' and Familie.Shape == 'RE' and Familie.Ist24h == '24h':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        break
            elif Param.art == 'LZU':
                for Familie in Liste_Familie:
                    if Familie.Art == 'Zuluft' and Familie.Shape == 'RE':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        break
            else:
                for Familie in Liste_Familie:
                    if Familie.Shape == 'RE':
                        if Param.parametername in Familie.Dict_Types.keys():
                            Param.Liste_type.append(Familie.Dict_Types[Param.parametername])
                        else:Param.Liste_type.append(Familie.defaulttype)
                        
    
    t = DB.Transaction(doc,'Laboranschlusstyp erstellen & anpassen')
    t.Start()
    for Param in Liste_Parameter:
        for familietyp in Param.Liste_type:
            if familietyp.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == Param.parametername:
                LaboranschlussType.Anpassen(familietyp, Param)
            else:
                newfamilietyp = familietyp.Duplicate(Param.parametername)
                doc.Regenerate()
                LaboranschlussType.Anpassen(newfamilietyp, Param)
    t.Commit()
    t.Dispose()
    doc.Dispose()
    del t
    del doc

