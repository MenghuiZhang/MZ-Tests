# coding: utf8
import System
from System.IO import Path, File
from rpw import revit,DB,UI
from System.Collections.Generic import List
from color_conversion import rgb2hsl,hsl2rgb
from System import Byte
import clr
clr.AddReference('System.Drawing')
from System.Drawing import Color

######################################################################################
def get_value(param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if param.StorageType.ToString() == 'ElementId':
        return param.AsValueString()
    elif param.StorageType.ToString() == 'Integer':
        value = param.AsInteger()
    elif param.StorageType.ToString() == 'Double':
        value = param.AsDouble()
    elif param.StorageType.ToString() == 'String':
        value = param.AsString()
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value
######################################################################################

def get_value_Family(familytyp,param):
    """gibt den gesuchten Wert ohne Einheit zurück"""
    if not familytyp.HasValue(param):return None

    if param.StorageType.ToString() == 'ElementId':
        return familytyp.AsValueString(param)
    elif param.StorageType.ToString() == 'Integer':
        value = familytyp.AsInteger(param)
    elif param.StorageType.ToString() == 'Double':
        value = familytyp.AsDouble(param)
    elif param.StorageType.ToString() == 'String':
        value = familytyp.AsString(param)
        return value

    try:
        # in Revit 2020
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
    except:
        try:
            # in Revit 2021/2022
            unit = param.GetUnitTypeId()
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            pass

    return value

# def convert_value_einheit(param,einheit):
#     """gibt den gesuchten Wert ohne Einheit zurück"""
#     if param.StorageType.ToString() == 'ElementId':
#         return param.AsValueString()
#     elif param.StorageType.ToString() == 'Integer':
#         value = param.AsInteger()
#     elif param.StorageType.ToString() == 'Double':
#         value = param.AsDouble()
#     elif param.StorageType.ToString() == 'String':
#         value = param.AsString()
#         return value

#     try:
#         # in Revit 2020
#         unit = param.DisplayUnitType
#         for el in System.Enum.GetValues(DB.BuiltInParameterGroup().GetType()):
#             if DB.LabelUtils.GetLabelFor(el).upper() == einheit.upper():
#                 unit = el
#                 break
#         value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
#     except:
#         try:
#             # in Revit 2021/2022
#             unit = param.GetUnitTypeId()
#             for el in dir(DB.UnitTypeId):
#                 if DB.LabelUtils.GetLabelForUnit(getattr(DB.UnitTypeId, el)).upper() == einheit.upper():
#                     unit = getattr(DB.UnitTypeId, el)
#                     break
#             value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
#         except:
#             pass

#     return value


def wert_schreibenbase(param,wert):
    if param:
        if param.IsReadOnly:
            return
        if wert is not None:
            try:
                if param.StorageType.ToString() == 'ElementId':
                    param.Set(wert)
                elif param.StorageType.ToString() == 'Integer':
                    param.Set(int(wert))
                elif param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                elif param.StorageType.ToString() == 'String':
                    param.Set(str(wert))
            except:pass

def get_Parameter(elem,GUID):
    return elem.get_Parameter(GUID)
