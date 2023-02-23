# coding: utf8
import Autodesk.Revit.DB as DB
from pyrevit import script

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'Raumluftbilanz')
