# coding: utf8
import Autodesk.Revit.DB as DB

doc = __revit__.ActiveUIDocument.Document

def ElementFamilyEquals_Filter(fam_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringEquals()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementFamilyContains_Filter(fam_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringContains()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WhereElementIsNotElementType()
    if Category:
        coll = coll.OfCategory(Category)
    
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementTypeEquals_Filter(fam_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringEquals()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementTypeContains_Filter(fam_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringContains()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementFamilyTypeMergeEquals_Filter(fam_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringEquals()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementFamilyTypeMergeContains_Filter(fam_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringContains()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementFamilyTypeSeparatelyEquals_Filter(fam_name = None, type_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringEquals()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    type_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    type_prov=DB.ParameterValueProvider(type_id)
    type_value_rule=DB.FilterStringRule(type_prov,param_equality,type_name,True)
    type_filter = DB.ElementParameterFilter(type_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WherePasses(type_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WherePasses(type_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()

def ElementFamilyTypeSeparatelyContains_Filter(fam_name = None, type_name = None, ansicht = False, Category = None, EntwurfsOption = True):
    param_equality = DB.FilterStringContains()
    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    type_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    type_prov=DB.ParameterValueProvider(type_id)
    type_value_rule=DB.FilterStringRule(type_prov,param_equality,type_name,True)
    type_filter = DB.ElementParameterFilter(type_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).WherePasses(fam_filter).WherePasses(type_filter).WhereElementIsNotElementType()
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).WherePasses(fam_filter).WherePasses(type_filter).WhereElementIsNotElementType()
    if Category:
        coll.OfCategory(Category)
    if EntwurfsOption:
        coll.ContainedInDesignOption(DB.ElementId(-1))
    return coll.ToElements()