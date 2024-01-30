# coding: utf8
from Autodesk.Revit.UI.Selection import ISelectionFilter,ObjectType
import Autodesk.Revit.DB as DB

class BaseSelectionFilter(ISelectionFilter):
    def __init__(self,func):
        self._func = func
    
    def AllowElement(self,elem):pass
    def AllowReference(self,ref,position):pass
