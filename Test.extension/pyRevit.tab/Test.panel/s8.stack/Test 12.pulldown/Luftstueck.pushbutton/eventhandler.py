# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB

class ErsteElement(Selection.ISelectionFilter):
    # Formteil
    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008016' or element.Category.Id.ToString() == '-2008010' or element.Category.Id.ToString() == '-2008020':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class ZweiteElement(Selection.ISelectionFilter):
    def __init__(self,elemids):
        self.elemids = elemids
    # Formteil
    def AllowElement(self,element):
        if element.Id.IntegerValue in self.elemids:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class ExternalEventListe(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.name = 'Erstellen ein Luftkanal'
        self.executeapp = self.verschieben
    
    def GetName(self):
        return self.name
    
    def Execute(self,uiapp):
        try:
            self.executeapp(uiapp)
        except:
            pass
    
    def getconnectedElems(self,elem):
        conns = elem.MEPModel.ConnectorManager.Connectors
        _dict = {}
        for conn in conns:
            if conn.IsConnected:
                refs = conn.AllRefs
                for ref in refs:
                    if ref.Owner.Category.Name not in ['Luftkanal Systeme','Luftkanaldämmung außen','Luftkanaldämmung innen','Rohrsysteme','Rohrdämmung']:
                        _dict[ref.Owner.Id.IntegerValue] = conn
        
        return _dict

    def verschieben(self,uiapp):
        if self.GUI.abstand.Text and self.GUI.ducttype.SelectedIndex != -1:

            while (True):
                try:
                    uidoc = uiapp.ActiveUIDocument
                    doc = uidoc.Document
                    ref1 = uidoc.Selection.PickObject(Selection.ObjectType.Element,ErsteElement(),'Wählt den ersten Teil aus')
                    elem1 = doc.GetElement(ref1.ElementId)

                    levelid = elem1.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId()

                    _dict = self.getconnectedElems(elem1)
                    ref2 = uidoc.Selection.PickObject(Selection.ObjectType.Element,ZweiteElement(_dict.keys()),'Wählt den ersten Teil aus')
                    conn0 = _dict[ref2.ElementId.IntegerValue]
                    conn1 = None

                    elem2 = doc.GetElement(ref2.ElementId)

                    try:
                        conns = elem2.MEPModel.ConnectorManager.Connectors
                    except:
                        conns = elem2.ConnectorManager.Connectors
                    
                    for conn in conns:
                        if conn.IsConnectedTo(conn0):
                            conn1 = conn
                            break
                                

                    t = DB.Transaction(doc,self.name)
                    t.Start()
                    conn0.DisconnectFrom(conn1)
                    elem2.Location.Move(conn0.CoordinateSystem.BasisZ * float(self.GUI.abstand.Text) / 304.8)
                    DB.Mechanical.Duct.Create(doc,self.GUI.DUCT_Dict[self.GUI.ducttype.SelectedItem],levelid,conn0,conn1)



                    t.Commit()
                    t.Dispose()
                except Exception as e:
                    print(e)
                    break
    
