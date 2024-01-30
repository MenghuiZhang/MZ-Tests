# coding: utf8
from IGF_Filters import BaseSelectionFilter,ObjectType,DB

class LinkSelectionFilter(BaseSelectionFilter):
    def __init__(self, doc, func):
        BaseSelectionFilter.__init__(self,func)
        self._doc = doc
    def AllowElement(self, elem):
        return True
    def AllowReference(self, ref, position):
        revitlink = self._doc.GetElement(ref.ElementId)
        if not isinstance(revitlink, DB.RevitLinkInstance) :
            return self._func(revitlink)
        return self._func(revitlink.GetLinkDocument().GetElement(ref.LinkedElementId))

class ElementSelectionFilter(BaseSelectionFilter):
    def __init__(self, func, func_ref = None):
        BaseSelectionFilter.__init__(self,func)
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
    def CreateElementSelectionFilter(func,func_ref = None):
        return ElementSelectionFilter(func,func_ref)
    @staticmethod
    def CreateLinkElementSelectionFilter(doc,func):
        return LinkSelectionFilter(doc,func)

class CurrentDocumentOption:
    @staticmethod
    def PickElements(uidoc, func, func_ref = None, Text = 'Elemente auswählen'):
        return [uidoc.Document.GetElement(ref.ElementId) \
                for ref in uidoc.Selection.PickObjects(\
                ObjectType.Element,\
                SelectionFilterFactory.CreateElementSelectionFilter(func,func_ref),\
                Text)]
    
    @staticmethod
    def PickElement(uidoc, func, func_ref = None, Text = 'Elemente auswählen'):
        return  uidoc.Document.GetElement(uidoc.Selection.PickObject(\
                ObjectType.Element,\
                SelectionFilterFactory.CreateElementSelectionFilter(func,func_ref),\
                Text)) \
                
class LinkDocumentOption:
    @staticmethod
    def PickElements(uidoc, func, Text = 'Elemente auswählen'):
        doc = uidoc.Document
        refs = uidoc.Selection.PickObjects(\
                ObjectType.LinkedElement,\
                SelectionFilterFactory.CreateLinkElementSelectionFilter(doc,func),\
                Text)
        
        
        elem = doc.GetElement(ref.ElementId).GetLinkDocument().GetElement(ref.LinkedElementId)
        
        doc.Dispose()
        return elem

    @staticmethod
    def PickElement(uidoc, func, Text = 'Elemente auswählen'):
        doc = uidoc.Document
        ref = uidoc.Selection.PickObject(\
                ObjectType.LinkedElement,\
                SelectionFilterFactory.CreateLinkElementSelectionFilter(doc,func),\
                Text)
        
        elems = []

        RVTLink = doc.GetElement(ref.ElementId)
        if RVTLink:
            elem = RVTLink.GetLinkDocument().GetElement(ref.LinkedElementId)
            elems.append(elem)
        
        doc.Dispose()
        return elems

class BothDocumentOption:
    @staticmethod
    def PickElements(uidoc, func,Text = 'Elemente auswählen'):
        doc = uidoc.Document
        refs = uidoc.Selection.PickObjects(\
                ObjectType.PointOnElement,\
                SelectionFilterFactory.CreateLinkElementSelectionFilter(doc,func),\
                Text)
        
        elems = []

        for ref in refs:
            RVTLink = doc.GetElement(ref.ElementId)
            if isinstance(RVTLink, DB.RevitLinkInsatnce):
                elem = RVTLink.GetLinkDocument().GetElement(ref.LinkedElementId)
                elems.append(elem)
            else:
                elems.append(RVTLink)
        
        doc.Dispose()
        return elems
    
    @staticmethod
    def PickElement(uidoc, func,Text = 'Elemente auswählen'):
        doc = uidoc.Document
        ref = uidoc.Selection.PickObject(\
                ObjectType.PointOnElement,\
                SelectionFilterFactory.CreateLinkElementSelectionFilter(doc,func),\
                Text)
        
      
        RVTLink = doc.GetElement(ref.ElementId)
        if isinstance(RVTLink, DB.RevitLinkInsatnce):
            elem = RVTLink.GetLinkDocument().GetElement(ref.LinkedElementId)
        else:
            elem = RVTLink
        
        doc.Dispose()
        return elem

class PickElementsOptionFactory:
    @staticmethod
    def CreateCurrentDocumentOption(uidoc, func, func_ref = None, Text = 'Elemente auswählen'):
        return CurrentDocumentOption.PickElements(uidoc, func, func_ref, Text)
    
    @staticmethod
    def CreateLinkDocumentOption(uidoc, func,Text = 'Elemente auswählen'):
        return LinkDocumentOption.PickElements(uidoc, func, Text)

    @staticmethod
    def CreateBothDocumentOption(uidoc, func,Text = 'Elemente auswählen'):
        return BothDocumentOption.PickElements(uidoc, func, Text)
    
class PickElementOptionFactory:
    @staticmethod
    def CreateCurrentDocumentOption(uidoc, func, func_ref = None, Text = 'Elemente auswählen'):
        return CurrentDocumentOption.PickElement(uidoc, func, func_ref, Text)
    
    @staticmethod
    def CreateLinkDocumentOption(uidoc, func,Text = 'Elemente auswählen'):
        return LinkDocumentOption.PickElement(uidoc, func, Text)

    @staticmethod
    def CreateBothDocumentOption(uidoc, func,Text = 'Elemente auswählen'):
        return BothDocumentOption.PickElement(uidoc, func, Text)
    
# def PickElements(uidoc,func,pickelementoption):
    # return pickelementoption.PickElemens(uidoc,func)
        