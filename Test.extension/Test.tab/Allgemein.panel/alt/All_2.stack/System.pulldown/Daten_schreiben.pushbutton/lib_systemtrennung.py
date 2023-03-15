import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI


def Fam_Exemplar_Filter(fam_name = None, typ_name = None, ansicht = False,doc=None):
    param_equality=DB.FilterStringEquals()
    if not fam_name:
        UI.TaskDialog.Show('Fehler','Bitte Geben Sie einen Familienamen ein.')
        return False

    fam_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    fam_prov=DB.ParameterValueProvider(fam_id)
    fam_value_rule=DB.FilterStringRule(fam_prov,param_equality,fam_name,True)
    fam_filter = DB.ElementParameterFilter(fam_value_rule)
    if not ansicht:
        coll = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).WherePasses(fam_filter)
    else:
        coll = DB.FilteredElementCollector(doc,doc.ActiveView.Id).OfClass(DB.FamilyInstance).WherePasses(fam_filter)
    if typ_name:
        typ_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        typ_prov=DB.ParameterValueProvider(typ_id)
        typ_value_rule=DB.FilterStringRule(typ_prov,param_equality,typ_name,True)
        typ_filter = DB.ElementParameterFilter(typ_value_rule)
        coll.WherePasses(typ_filter)

    return coll

class Verbinder(object):
    def __init__(self,elementid,fliessrichtung,doc):
        self.fliessrichtung = fliessrichtung
        self.doc = doc
        self.elem_id = elementid
        self.elem = self.doc.GetElement(self.elem_id)
        self.connid_Leitung = ''
        self.connid_verbinder = ''
        self.elemid_Leitung = ''
        self.Daten_ermitteln()
    
    @property
    def conns(self):
        return self.elem.MEPModel.ConnectorManager.Connectors
    
    @property
    def connid_verbinder(self):
        return self._connid_verbinder
    @connid_verbinder.setter
    def connid_verbinder(self,value):
        self._connid_verbinder = value
    @property
    def connid_Leitung(self):
        return self._connid_Leitung
    @connid_Leitung.setter
    def connid_Leitung(self,value):
        self._connid_Leitung = value
    @property
    def elemid_Leitung(self):
        return self._elemid_Leitung
    @elemid_Leitung.setter
    def elemid_Leitung(self,value):
        self._elemid_Leitung = value
    def Daten_ermitteln(self):
        for conn in self.conns:
            if conn.Direction.ToString() != self.fliessrichtung:
                continue
            self.connid_verbinder = conn.Id
            for ref in conn.AllRefs:
                try:
                    if ref.Owner.Category.Id.ToString() in ['-2008000','-2008044']:
                        self.connid_Leitung = ref.Id
                        self.elemid_Leitung = ref.Owner.UniqueId
                        break
                except:
                    pass
            if self.connid_Leitung:
                break

    def werte_schreiben(self):
        def wert_schreiben(param,value):
            try:
                self.elem.LookupParameter(param).Set(value)
            except:
                print('Fehler beim Werte-Schreiben in Parameter "{}".format(param)')
        wert_schreiben('ConnectorID_Rohre',int(self.connid_Leitung))
        wert_schreiben('ConnectorID_Verbinder',int(self.connid_verbinder))
        wert_schreiben('UniqueId_Rohre',str(self.elemid_Leitung))