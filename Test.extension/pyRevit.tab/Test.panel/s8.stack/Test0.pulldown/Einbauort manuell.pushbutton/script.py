# coding: utf8
from rpw import revit, UI, DB


__title__ = "Einbauort manuell"
__doc__ = """Exportiert eine AK-Liste. Verbesserte Filterfunktion"""
__authors__ = "Maximilian Prachtel"

doc = revit.doc
uidoc = revit.uidoc

class MEPRaumFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2003600':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class AllgemeinFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2001140':
            return True
        else:
            return False
     

    def AllowReference(self,reference,XYZ):
        return False


while(True):
    try:
        re0 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,AllgemeinFilter(),'Bauteil auswählen')
        re1 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element,MEPRaumFilter(),'MEP Raum auswählen')
        bauteil = doc.GetElement(re0)
        mep = doc.GetElement(re1)
        t = DB.Transaction(doc,'Einbauort')
        t.Start()
        try:
            bauteil.LookupParameter('IGF_X_Einbauort').Set(mep.Number)
        except:
            UI.TaskDialog.Show('Fehler','nicht klappt!')
        t.Commit()
        t.Dispose()
    except Exception as e:
        # print(e)
   
        break

