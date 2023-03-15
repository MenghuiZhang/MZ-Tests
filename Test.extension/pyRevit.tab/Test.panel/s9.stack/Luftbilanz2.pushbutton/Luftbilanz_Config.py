# coding: utf8
import Autodesk.Revit.DB as DB
from pyrevit import script
from config import ProjektConfig

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = ProjektConfig.get_config_static(number + ' - ' + name,'Raumluftbilanz')
