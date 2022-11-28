# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from IGF_lib import get_value

class ListeExternalEvent(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.name = ''
        self.ExcuteApp = ''
        
    def Execute(self,uiapp):
        if self.ExcuteApp:
            try:
                self.ExcuteApp(uiapp)
            except:
                pass
        
    def GetName(self):
        return self.name

    def Dimension(self,uiapp):
        self.name = 'Kanalrechner'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        if self.GUI.Kanalfitting.IsChecked:
            for elid in uidoc.Selection.GetElementIds():
                el = doc.GetElement(elid)
                if el.Category.Id.IntegerValue in [-2008000,-2008010]:
                    elems.append(el)
            if len(elems) == 0:
                TaskDialog.Show('Info','Kein Luftkanal/Luftkanalformteile ausgewählt!')
                return
        else:
            for elid in uidoc.Selection.GetElementIds():
                el = doc.GetElement(elid)
                if el.Category.Id.IntegerValue in [-2008000]:
                    elems.append(el)
            if len(elems) == 0:
                TaskDialog.Show('Info','Kein Luftkanal ausgewählt!')
                return

        elemids = [e.Id.ToString() for e in elems]
        t = DB.Transaction(doc,'Kanäle anpassen')
        t.Start()

            
           
    
        for el in elems:
            if el.Category.Id.IntegerValue == -2008000:
                try:
                    if self.GUI.eingabeb.Text not in ['',None]:
                        if el.LookupParameter('Breite'):
                            el.LookupParameter('Breite').SetValueString(self.GUI.eingabeb.Text.replace(',','.'))
                        
                    if self.GUI.eingabeh.Text not in ['',None]:
                        if el.LookupParameter('Höhe'):
                            el.LookupParameter('Höhe').SetValueString(self.GUI.eingabeh.Text.replace(',','.'))
                    if self.GUI.eingabed.Text not in ['',None]:
                        if el.LookupParameter('Durchmesser'):
                            el.LookupParameter('Durchmesser').SetValueString(self.GUI.eingabed.Text.replace(',','.'))
                        
                except Exception as e:print(e)
                    
            else:
                try:
                    conns = el.MEPModel.ConnectorManager.Connectors
                    
                    for conn in conns:
                        refs = conn.AllRefs
                        for ref in refs:
                            if conns.Size > 2:
                                if ref.Owner.Id.ToString() not in elemids:continue
                            else:pass
                       
                            mepinfo = conn.GetMEPConnectorInfo()
                            d = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_DIAMETER))
                            r = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_RADIUS))
                            h = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_HEIGHT))
                            b = mepinfo.GetAssociateFamilyParameterId(DB.ElementId(DB.BuiltInParameter.CONNECTOR_WIDTH))
                            if d != DB.ElementId.InvalidElementId:
                                try:
                                    param = doc.GetElement(d)
                                    if self.GUI.eingabed.Text not in ['',None]:
                                        el.get_Parameter(param.GetDefinition()).SetValueString(self.GUI.eingabed.Text.replace(',','.'))
                                except:pass
                            if r != DB.ElementId.InvalidElementId:
                                try:
                                    param = doc.GetElement(r)
                                    if self.GUI.eingabed.Text not in ['',None]:
                                        el.get_Parameter(param.GetDefinition()).SetValueString(str(round(float(self.GUI.eingabed.Text.replace(',','.'))/2,1)))
                                except:pass
                            if h != DB.ElementId.InvalidElementId:
                                try:
                                    param = doc.GetElement(h)
                                    if self.GUI.eingabeh.Text not in ['',None]:
                                        el.get_Parameter(param.GetDefinition()).SetValueString(self.GUI.eingabeh.Text.replace(',','.'))
                                except:pass
                            if b != DB.ElementId.InvalidElementId:
                                try:
                                    param = doc.GetElement(b)
                                    if self.GUI.eingabeb.Text not in ['',None]:
                                        el.get_Parameter(param.GetDefinition()).SetValueString(self.GUI.eingabeb.Text.replace(',','.'))
                                except:pass
                except:pass

        doc.Regenerate()
        t.Commit()
    
    def GUIAktualiesieren(self,uiapp):
        self.name = 'GUI aktualisieren'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        for elid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(elid)
            if el.Category.Id.IntegerValue in [-2008000]:
                elems.append(el)
        if len(elems) != 1:
            TaskDialog.Show('Info','Bitte nur ein Luftkanal auswählen!')
            return
        for el in elems:
            try:
                b_p = el.get_Parameter(DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                h_p = el.get_Parameter(DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                d_p = el.get_Parameter(DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)

                mh_p = el.get_Parameter(DB.BuiltInParameter.RBS_OFFSET_PARAM)
                oh_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_TOP_ELEVATION)
                uh_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION)

                re_p = el.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)

                v_p = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM)

                t_p = el.get_Parameter(DB.BuiltInParameter.RBS_VELOCITY)
                
                if b_p:
                    b = str(int(round(get_value(b_p),1)))
                    self.GUI.eingabeb.Text = b
                if h_p:
                    h = str(int(round(get_value(h_p),1)))
                    self.GUI.eingabeh.Text = h
      
                if d_p:
                    d = str(int(round(get_value(d_p),1)))
                    self.GUI.eingabed.Text = d
        
                if oh_p:
                    oh = str(round(get_value(oh_p),1))
                    self.GUI.obentext.Text = oh.replace('.',',')
                
                if mh_p:
                    mh = str(round(get_value(mh_p),1))
                    self.GUI.mittetext.Text = mh.replace('.',',')
                
                if uh_p:
                    uh = str(round(get_value(uh_p),1))
                    self.GUI.untentext.Text = uh.replace('.',',')
                
                if re_p:
                    re = re_p.AsValueString()
                    self.GUI.refe.SelectedItem = re
                
                if v_p:
                    v = str(int(round(get_value(v_p),1)))
                    self.GUI.eingabevol.Text = v
                
                if t_p:
                    t = str(round(get_value(t_p),1))
                    self.GUI.eingabetem.Text = t.replace('.',',')
           

            except Exception as e:
                print(e)

    def HEIGHTAnpassen(self,uiapp):
        self.name = 'Höhe anpassen'
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        elems = []
        for elid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(elid)
            if el.Category.Id.IntegerValue in [-2008000]:
                elems.append(el)
        if len(elems) != 1:
            TaskDialog.Show('Info','Bitte nur ein Luftkanal auswählen!')
            return

        t = DB.Transaction(doc,'Höhe übertragen')
        t.Start()
        if self.GUI.refe.SelectedIndex != -1:elems[0].get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM).Set(self.GUI.ebenedict[self.GUI.refe.SelectedItem.ToString()])
        if self.GUI.oben.IsChecked:elems[0].get_Parameter(DB.BuiltInParameter.RBS_DUCT_TOP_ELEVATION).SetValueString(self.GUI.obentext.Text.replace(',','.'))
        elif self.GUI.mitte.IsChecked:elems[0].get_Parameter(DB.BuiltInParameter.RBS_OFFSET_PARAM).SetValueString(self.GUI.mittetext.Text.replace(',','.'))
        else:elems[0].get_Parameter(DB.BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION).SetValueString(self.GUI.untentext.Text.replace(',','.'))
        
        t.Commit()
        t.Dispose()
