# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import clr


class Filters(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if  element.Category.Name in ['Rohrzubehör','Luftkanalformteile' ]:
            return True
        # if element.Category.Id.ToString() == '-2000485':
        #     return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False
class Rohrfilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if  element.Category.Id.IntegerValue == -2008044:
            return True
        # if element.Category.Id.ToString() == '-2000485':
        #     return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Rohrzubehorfilter(Selection.ISelectionFilter):
    def AllowElement(self,element):
        if  element.Category.Id.IntegerValue == -2008055:
            return True
        # if element.Category.Id.ToString() == '-2000485':
        #     return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter1(Selection.ISelectionFilter):
    def __init__(self,elemid):
        self.elemid = elemid
    def AllowElement(self,element):
        try:
            if element.Category.Id == self.elemid:
                return True
            else:
                return False
        except:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class BESCHIFTUNG2(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        while(True):
            try:
                ref_rohe = uidoc.Selection.PickObject(Selection.ObjectType.Element,Rohrfilter())
                ref_Rohrzubehoer = uidoc.Selection.PickObject(Selection.ObjectType.Element,Rohrzubehorfilter())
                 
                t = DB.Transaction(doc,'anpassen')
                t.Start()
                try:
                    zube = doc.GetElement(ref_Rohrzubehoer)
                    symbols = zube.Symbol.Family.GetFamilySymbolIds()
                    durchmesser = doc.GetElement(ref_rohe).LookupParameter('Durchmesser').AsValueString()
                    _dict = {doc.GetElement(el).get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString():el for el in symbols }
                   
                    for el in sorted(_dict.keys()):
                        if el.find(durchmesser) != -1:
                            zube.ChangeTypeId(_dict[el])
                except Exception as e:
                    print(e)
                    t.Commit()
                    break


                t.Commit()
                t.Dispose()
            except Exception as e:
                print(e)
                return
        

    def GetName(self):
        return "anpassen"