from Autodesk.Revit.UI.Selection import ISelectionFilter,ObjectType
import Autodesk.Revit.DB as DB


class BaseSelectionFilter(ISelectionFilter):
    def __init__(self,func):
        self._func = func
    
    def AllowElement(self,elem):pass
    def AllowReference(self,ref,position):pass

class LinkSelectionFilter(BaseSelectionFilter):
    def __init__(self, doc, func):
        super().__init__(func)
        self._doc = doc
    def AllowElement(self, elem):
        return True
    def AllowReference(self, ref, position):
        revitlink = self._doc.GetElement(ref.ElementId)
        if revitlink is not DB.RevitLinkInsatnce:
            return self._func(revitlink)
        return self._func(revitlink.GetLinkDocument().GetElement(ref.LinkedElementId))

class ElementSectionFilter(BaseSelectionFilter):
    def __init__(self, func, func_ref = None):
        super().__init__(func)
        self._func_ref = func_ref
    
    def AllowElement(self, elem):
        return self._func(elem)
    
    def AllowReference(self, ref, position):
        if self._func_ref == None:
            return True
        else:
            return self._func_ref(ref)

class SelectionFilterFactory:
    @staticmethod
    def CreateElementSelectionFIlter(func):
        return ElementSectionFilter(func)
    @staticmethod
    def CreateLinkSelectionFIlter(doc,func):
        return LinkSelectionFilter(doc,func)

# Interface in c# 
class IPickElementOption:
    def PickElemens(self,uidoc,func):
        pass

class CurrentDocumentOption(IPickElementOption):
    def PickElemens(self, uidoc, func):
        return [uidoc.Document.GetElement(ref.ElementId) \
                for ref in uidoc.Selection.PickObjects(\
                ObjectType.Element,\
                SelectionFilterFactory.CreateElementSelectionFIlter(func))]

class LinkDocumentOption(IPickElementOption):
    
    def PickElemens(self, uidoc, func):
        doc = uidoc.Document
        refs = uidoc.Selection.PickObjects(\
                ObjectType.LinkElement,\
                SelectionFilterFactory.CreateLinkSelectionFIlter(doc,func))
        
        elems = []

        for ref in refs:
            RVTLink = doc.GetElement(ref.ElementId)
            if RVTLink:
                elem = RVTLink.GetLinkDocument().GetElement(ref.LinkedElementId)
                elems.append(elem)
        
        doc.Dispose()
        return elems

class BothDocumentOption(IPickElementOption):
    def PickElemens(self, uidoc, func):
        doc = uidoc.Document
        refs = uidoc.Selection.PickObjects(\
                ObjectType.PointOnElement,\
                SelectionFilterFactory.CreateLinkSelectionFIlter(doc,func))
        
        elems = []

        for ref in refs:
            RVTLink = doc.GetElement(ref.ElementId)
            if RVTLink:
                elem = RVTLink.GetLinkDocument().GetElement(ref.LinkedElementId)
                elems.append(elem)
            else:
                elems.append(RVTLink)
        
        doc.Dispose()
        return elems

class PickElementsOptionFactory:
    @staticmethod
    def CreateCurrentDocumentOption():
        return CurrentDocumentOption()
    
    @staticmethod
    def CreateLinkDocumentOption():
        return LinkDocumentOption()

    @staticmethod
    def CreateBothDocumentOption():
        return BothDocumentOption()
    
def PickElements(uidoc,func,pickelementoption):
    return pickelementoption.PickElemens(uidoc,func)
        