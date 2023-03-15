# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from rpw import revit
from System import Type

doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveView

class Children(object):
    def __init__(self,name,art):
        self.name = name
        self.checked = False
        self.Art = art
        self.ListeHidden = []
        self.ListeShow = [] 
        self.children = ObservableCollection[Children]()
        self.parent = None
        self.expand = False

liste_class = List[Type]()
liste_class.Add(DB.IndependentTag)
liste_class.Add(DB.SpatialElementTag)
TagsFilter = DB.ElementMulticlassFilter(liste_class)

AllTags = {}
coll_Ansicht = DB.FilteredElementCollector(doc,view.Id).WhereElementIsNotElementType().WherePasses(TagsFilter).ToElementIds()
for el in coll_Ansicht:
    elem = doc.GetElement(el)
    cate = elem.Category.Name
    name = elem.Name
    fam = doc.GetElement(elem.GetTypeId()).FamilyName
    if cate not in AllTags.keys():
        AllTags[cate] = {}
    if fam not in AllTags[cate].keys():
        AllTags[cate][fam] = {}
    if name not in AllTags[cate][fam].keys():
        AllTags[cate][fam][name] = [List[DB.ElementId](),List[DB.ElementId]()]
  
    AllTags[cate][fam][name][0].Add(el)

coll_doc = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(TagsFilter).ToElementIds()
for el in coll_doc:
    elem = doc.GetElement(el)
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
    
    AllTags[cate][fam][name][1].Add(el)

ItemsSource = ObservableCollection[Children]()

for cate in sorted(AllTags.keys()):
    kate = Children(cate,'Kategorie')
    ItemsSource.Add(kate)
    Liste_fam = []
    for fam in sorted(AllTags[cate].keys()):
        Fam = Children(fam,'Familie')
        Fam.parent = kate
        kate.children.Add(Fam)
        Liste_typ = []
        for typ in sorted(AllTags[cate][fam].keys()):
            Typ = Children(typ,'Typ')
            Typ.parent = Fam
            Fam.children.Add(Typ)
            Typ.ListeShow = AllTags[cate][fam][typ][0]
            Typ.ListeHidden = AllTags[cate][fam][typ][1]
            if Typ.ListeShow.Count == 0:
                if 0 not in Liste_typ:
                    Liste_typ.append(0)
                Typ.checked = False
            elif Typ.ListeHidden.Count == 0:
                if 1 not in Liste_typ:
                    Liste_typ.append(1)
                Typ.checked = True
            else:
                if 2 not in Liste_typ:
                    Liste_typ.append(2)
                Typ.checked = None
        if len(Liste_typ) == 1:
            if Liste_typ[0] == 0:
                Fam.checked = False
                if 0 not in Liste_fam:
                    Liste_fam.append(0)
            
            elif Liste_typ[0] == 1:
                Fam.checked = True
                if 1 not in Liste_fam:
                    Liste_fam.append(1)
            else:
                Fam.checked = None
                if 2 not in Liste_fam:
                    Liste_fam.append(2)
        else:
            Fam.checked = None
            if 2 not in Liste_fam:
                Liste_fam.append(2)
    if len(Liste_fam) == 1:
        if Liste_fam[0] == 0:
            kate.checked = False
        elif Liste_fam[0] == 1:
            kate.checked = True
        else:
            kate.checked = None
    else:
        kate.checked = None            

class EINAUS(IExternalEventHandler):
    def __init__(self):
        self.Liste = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        view = uidoc.ActiveView
        t = DB.Transaction(doc,'Ein-/Ausblenden')
        t.Start()
        for cate in self.Liste:
            for fam in cate.children:
                for typ in fam.children:
                    if typ.checked is True:
                        try:view.UnhideElements(typ.ListeHidden)
                        except:pass
                        try:view.UnhideElements(typ.ListeShow)
                        except:pass
                

                    elif typ.checked is False:
                        try:view.HideElements(typ.ListeHidden)
                        except:pass
                        try:view.HideElements(typ.ListeShow)
                        except:pass

        t.Commit()
        t.Dispose()

        TaskDialog.Show('Info','Erledigt!') 


    def GetName(self):
        return "Beschriftungen ein-/ausblenden"