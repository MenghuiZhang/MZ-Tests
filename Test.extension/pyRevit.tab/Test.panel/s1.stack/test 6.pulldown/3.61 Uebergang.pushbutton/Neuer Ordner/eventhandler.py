# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB

class TransitionFitting(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Liste = []
        for elid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(elid)
            if el.Category.Name == 'Luftkanäle':
                 Liste.append(el)
        if len(Liste) != 2:
            TaskDialog.Show('Fehler','Bitte zwei Luftkanäle auswählen!')
            return
       
        el0 = Liste[0]
        el1 = Liste[1]
        conns0 = list(el0.ConnectorManager.Connectors)
        conns1 = list(el1.ConnectorManager.Connectors)
        
        co0 = None
        co1 = None
        distance = 1000
        l0 = DB.Line.CreateBound(conns0[0].Origin,conns0[1].Origin)
        l1 = DB.Line.CreateBound(conns1[0].Origin,conns1[1].Origin)
        if not ((l0.Direction+l1.Direction).GetLength() < 0.00001 or (l0.Direction-l1.Direction).GetLength() < 0.00001):
            TaskDialog.Show('Fehler','Die Luftkanäle sind nicht Parallel!')
            return
        for con in conns0:
            for con1 in conns1:
                if con.IsConnected == False and con1.IsConnected == False:
                    dis = con.Origin.DistanceTo(con1.Origin)
                    if dis < distance:
                        distance = dis
                        co0 = con
                        co1 = con1
        if not (co0 and co1):
            TaskDialog.Show('Fehler','Kein oder nur ein offener Anschluss gefunden!')
            return
        
        o0 = co0.Origin
        o = co0.Origin-co1.Origin
        ln = l0.Direction.Normalize()
        length = o.DotProduct(ln)
        t = DB.Transaction(doc,'Übergangsstück')
        t.Start()
        try:
            fi = doc.Create.NewTransitionFitting(co0, co1)
            doc.Regenerate()
        except:
            fi = doc.Create.NewTransitionFitting(co1, co0)
            doc.Regenerate()  
        cs = fi.MEPModel.ConnectorManager.Connectors
        for con in cs:
            if con.Origin.DistanceTo(co0.Origin) < 0.001:
                
                if not con.IsConnected:
                    con.ConnectTo(co0)
                    con.Width = co0.Width                                    
                    con.Height = co0.Height
            else:

                if not con.IsConnected:
                    con.ConnectTo(co1)
                    con.Width = co1.Width                                    
                    con.Height = co1.Height
        try:fi.LookupParameter('L').Set(length)
        except Exception as e:print(e)
        doc.Regenerate()
        try:fi.Location.Move(o0-co0.Origin)
        except Exception as e:print(e)
        

        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Übergansstück erstellen"