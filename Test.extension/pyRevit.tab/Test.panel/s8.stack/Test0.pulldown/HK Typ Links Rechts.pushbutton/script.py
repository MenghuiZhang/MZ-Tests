# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit,DB
from pyrevit import script, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.ComponentModel import INotifyPropertyChanged ,PropertyChangedEventArgs
from math import pi
import os


__title__ = "Heizkörpertyp"
__doc__ = """

Kombirahmen --> MagiCAD DistributionBox
[2022.04.13]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

# try:
#     getlog(__title__)
# except:
#     pass

# try:
#     getloglocal(__title__)
# except:
#     pass

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'HK-Bauteile')

class TemplateItemBase(INotifyPropertyChanged):
    def __init__(self):
        self.propertyChangedHandlers = []

    def RaisePropertyChanged(self, propertyName):
        args = PropertyChangedEventArgs(propertyName)
        for handler in self.propertyChangedHandlers:
            handler(self, args)
            
    def add_PropertyChanged(self, handler):
        self.propertyChangedHandlers.append(handler)
        
    def remove_PropertyChanged(self, handler):
        self.propertyChangedHandlers.remove(handler)

class FamilienSymbol(TemplateItemBase):
    def __init__(self,name,typ,elems):
        TemplateItemBase.__init__(self)
        self._checked = False
        self.familyname = name
        self.typname = typ
        self.elems = elems
        self.Familie = None
        self.Symbol = None
        self.Dict_Symbols = {}
        self.Liste = []
        self.get_Symbol()
        self.get_Symbols()
        self.nichtmirroredfamilie = self.Symbol
        self.mirroredfamilie = self.Symbol

        self._nichtmirroredfamilieindex = -1
        self._mirroredfamilieindex = -1

        self.nichtmirroredfamilieindex = self.Liste.index(self.typname)
        self.mirroredfamilieindex = self.Liste.index(self.typname)
    
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self,value):
        if value != self._checked:
            self._checked = value
            self.RaisePropertyChanged('checked')
    
    @property
    def nichtmirroredfamilieindex(self):
        return self._nichtmirroredfamilieindex
    @nichtmirroredfamilieindex.setter
    def nichtmirroredfamilieindex(self,value):
        if value != self._nichtmirroredfamilieindex:
            self._nichtmirroredfamilieindex = value
            self.RaisePropertyChanged('nichtmirroredfamilieindex')
    
    @property
    def mirroredfamilieindex(self):
        return self._mirroredfamilieindex
    @mirroredfamilieindex.setter
    def mirroredfamilieindex(self,value):
        if value != self._mirroredfamilieindex:
            self._mirroredfamilieindex = value
            self.RaisePropertyChanged('mirroredfamilieindex')
    
    def get_Symbol(self):
        for el in self.elems:
            self.Familie = el.Symbol.Family
            self.Symbol = el.Symbol.Id
            break
    
    def get_Symbols(self):
        for elid in self.Familie.GetFamilySymbolIds():
            elem = doc.GetElement(elid)
            Name = elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            self.Dict_Symbols[Name] = elid
        self.Liste = sorted(self.Dict_Symbols.keys())

HEIZKOERPER_IS = ObservableCollection[FamilienSymbol]()

def get_Heizkoeper_IS():
    Dict = {}
    HLSs = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)\
        .WhereElementIsNotElementType().ToElements()
    for el in HLSs:
        FamilyName = el.Symbol.FamilyName
        typname = el.Name
        if FamilyName not in Dict.keys():
            Dict[FamilyName] = {}
        if typname not in Dict[FamilyName].keys():
            Dict[FamilyName][typname] = []
        Dict[FamilyName][typname].append(el)
    
    for fam in sorted(Dict.keys()):
        for typ in sorted(Dict[fam].keys()):
            HEIZKOERPER_IS.Add(FamilienSymbol(fam,typ,Dict[fam][typ]))
    
get_Heizkoeper_IS()

class Heizkoerper:
    def __init__(self,elemid,familie):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.Mirrored = self.elem.Mirrored
        self.conn1_list = []
        self.conn2_list = []
        self.ponit_alt = None
        self.point_neu = None
        self.pinned = self.elem.Pinned
        try:
            self.nichtmirroredfamilie = familie.nichtmirroredfamilie
            self.mirroredfamilie = familie.mirroredfamilie
        except:
            pass
    
    def changenichtmirrored(self):
        try:
            if self.elem.Symbol.Id.IntegerValue != self.nichtmirroredfamilie.IntegerValue:
                self.elem.ChangeTypeId(self.nichtmirroredfamilie)
        except:pass
    
    def changemirrored(self):
        try:
            if self.elem.Symbol.Id.IntegerValue != self.mirroredfamilie.IntegerValue:
                self.elem.ChangeTypeId(self.mirroredfamilie)
        except:pass
 
        
    def get_verbundene_elements(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for con in conns:
            if con.PipeSystemType.ToString() == 'ReturnHydronic':
                self.ponit_alt = con.Origin
            if con.IsConnected:
                if con.PipeSystemType.ToString() == 'ReturnHydronic':
                    for ref in con.AllRefs:
                        self.conn1_list.append(ref)
                        
                else:
                    for ref in con.AllRefs:
                        self.conn2_list.append(ref)
    
    def Disconnect(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for con in conns:
            if con.PipeSystemType.ToString() == 'ReturnHydronic':
                for conn_temp in self.conn1_list:
                    try:conn_temp.DisconnectFrom(con)
                    except:pass
            else:
                for conn_temp in self.conn2_list:
                    try:conn_temp.DisconnectFrom(con)
                    except:pass
        doc.Regenerate()
    
    def UnPinned(self):
        self.elem.Pinned = False
        doc.Regenerate()
    
    def PinnedRuecksetzen(self):
        self.elem.Pinned = self.pinned

    def entspiegeln(self):
        transform = self.elem.GetTransform()
        plane_XZ = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisX,transform.BasisZ)
        plane_YZ = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisY,transform.BasisZ)
       
        DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([self.elemid]), plane_XZ,False)
        doc.Regenerate()
        if self.elem.Mirrored:
            DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([self.elemid]), plane_XZ,False)
            DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([self.elemid]), plane_YZ,False)
        doc.Regenerate()
    
    def Rotate(self):
        transform = self.elem.GetTransform()
        origin = transform.Origin 
        line = DB.Line.CreateUnbound(origin,transform.BasisZ)
        DB.ElementTransformUtils.RotateElements(doc, List[DB.ElementId]([self.elemid]), line,pi)
        doc.Regenerate()
        line.Dispose()
    
    def Move(self):
        supply = self.elem.LookupParameter('Supply_connector_position')
        return_ = self.elem.LookupParameter('Return_connector_position')
        if supply.AsInteger() == 1:

            self.elem.LookupParameter('Supply_connector_position').Set(4)
            self.elem.LookupParameter('Return_connector_position').Set(5)
        else:

            self.elem.LookupParameter('Supply_connector_position').Set(1)
            self.elem.LookupParameter('Return_connector_position').Set(2)
        doc.Regenerate()
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for con in conns:
            if con.PipeSystemType.ToString() == 'ReturnHydronic':
                self.point_neu = con.Origin
        DB.ElementTransformUtils.MoveElements(doc,List[DB.ElementId]([self.elemid]),self.ponit_alt-self.point_neu)
        doc.Regenerate()
    
    def Connect(self):
        for con in self.elem.MEPModel.ConnectorManager.Connectors:
            if con.PipeSystemType.ToString() == 'ReturnHydronic':
                for ref in self.conn1_list:
                    try:con.ConnectTo(ref)
                    except:pass
            else:
                for ref in self.conn2_list:
                    try:con.ConnectTo(ref)
                    except:pass
    
    def Familiebearbeiten(self):
        if self.Mirrored:
            try:self.Mirrorbearbeiten()
            except Exception as e:
                logger.error(e)
                logger.error('elemid {}'.format(self.elemid.ToString()))
                print(30*'-')
        else:
            try:self.changenichtmirrored()
            except Exception as e:
                logger.error(e)
                logger.error('elemid {}'.format(self.elemid.ToString()))
                print(30*'-')

    
    def Mirrorbearbeiten(self):
        self.get_verbundene_elements()
        self.UnPinned()
        self.Disconnect()
        self.entspiegeln()
        self.Rotate()
        self.Move()
        self.Connect()
        self.PinnedRuecksetzen()
        self.changemirrored()

class GUI(forms.WPFWindow):
    def __init__(self):
        self.HEIZKOERPER_IS = HEIZKOERPER_IS  
        self.read_config() 
        self.temp_coll0 = ObservableCollection[FamilienSymbol]()          
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.lv_HK.ItemsSource = self.HEIZKOERPER_IS
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
    
    def read_config(self):
        try:
            _dict = config.HK_Familie_dict
            for el in self.HEIZKOERPER_IS:
                if el.familyname in _dict.keys():
                    if el.typname in _dict[el.familyname].keys():
                        el.checked = True

                        try:
                            el.nichtmirroredfamilieindex = el.Liste.index(_dict[el.familyname][el.typname][0])
                            el.nichtmirroredfamilie = el.Dict_Symbols[_dict[el.familyname][el.typname][0]]
                        except Exception as e:pass

                        try:
                            el.mirroredfamilieindex = el.Liste.index(_dict[el.familyname][el.typname][1])
                            el.mirroredfamilie = el.Dict_Symbols[_dict[el.familyname][el.typname][1]]
                        except Exception as e:pass
        except Exception as e:pass
    
    def write_config(self):
        try:
        
            _dict = {}
            for el in self.HEIZKOERPER_IS:
                if el.checked:
                    if el.familyname not in _dict.keys():
                        _dict[el.familyname] = {}
                    if el.typname not in _dict[el.familyname].keys():
                        _dict[el.familyname][el.typname] = []
                
                    _dict[el.familyname][el.typname].append(el.Liste[el.nichtmirroredfamilieindex])
                    _dict[el.familyname][el.typname].append(el.Liste[el.mirroredfamilieindex])
            config.HK_Familie_dict = _dict

        except Exception as e:pass
        script.save_config()
    
    def checkedchanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.lv_HK.SelectedIndex != -1:
            if item in self.lv_HK.SelectedItems:
                for el in self.lv_HK.SelectedItems:el.checked = checked
    
    def textchanged0(self,sender,e):
        text = sender.Text
        if not text:
            self.lv_HK.ItemsSource = self.HEIZKOERPER_IS
            return 
        self.temp_coll0.Clear()
        for el in self.HEIZKOERPER_IS:
            if el.familyname.upper().find(text.upper()) != -1 or el.typname.upper().find(text.upper()) != -1:
                self.temp_coll0.Add(el)
        self.lv_HK.ItemsSource = self.temp_coll0

    def visibilitychanged(self,sender,e):
        if sender.IsChecked:
            for el in self.HEIZKOERPER_IS:
                if el.checked:
                    el.visibility = 1
                else:
                    el.visibility = 0
        else:
            for el in self.HEIZKOERPER_IS:
                el.visibility = 1
        self.lv_HK.Items.Refresh()

    def OK(self,sender,args):
        self.write_config()
        self.Close()
        Liste = []
        for el in self.HEIZKOERPER_IS:
            if el.checked:
                Liste.append(el)
                el.nichtmirroredfamilie = el.Dict_Symbols[el.Liste[el.nichtmirroredfamilieindex]]
                el.mirroredfamilie = el.Dict_Symbols[el.Liste[el.mirroredfamilieindex]]
          


        if len(Liste) == 0:
            return
        t = DB.Transaction(doc,'Heizkörpertyp ändern')
        t.Start()
        with forms.ProgressBar(title="{value}/{max_value} Heizkörper",cancellable=True, step=1) as pb:
            for n,familie in enumerate(Liste):
                if pb.cancelled:
                    t.RollBack()
                    script.exit()
                pb.title = str(n+1) + '/' + str(len(Liste)) + ' Familie' + '---' + familie.familyname + ': ' + familie.typname + '---' +"{value}/{max_value} Heizkörper"
                for n1, elem in enumerate(familie.elems):
                    if pb.cancelled:
                        t.RollBack()
                        script.exit()
                    pb.update_progress(n1+1, len(familie.elems))
                    heizkoerper = Heizkoerper(elem.Id,familie)
                    heizkoerper.Familiebearbeiten()
        t.Commit()

    def cancel(self,sender,args):
        self.Close()
    
    # def nichtmirroredchanged(self,sender,e):
    #     item = sender.DataContext
    #     print(item)
    #     print(sender)
    #     # if item.nichtmirroredfamilieindex != -1:
    #     #     item.nichtmirroredfamilie = item.Dict_Symbols[sender.SelectedItem]
    
    # def mirroredchanged(self,sender,e):
    #     item = sender.DataContext
    #     print(item)
        # if item.mirroredfamilieindex != -1:
        #     item.mirroredfamilie = item.Dict_Symbols[sender.SelectedItem]

# gui = GUI()
# gui.ShowDialog()

t = DB.Transaction(doc,'Heizkörpertyp ändern')
t.Start()
for elid in uidoc.Selection.GetElementIds():
    heizkoerper = Heizkoerper(elid,None)
    heizkoerper.get_verbundene_elements()
    heizkoerper.UnPinned()
    heizkoerper.Disconnect()
    # heizkoerper.entspiegeln()
    heizkoerper.Rotate()
    heizkoerper.Move()
    heizkoerper.Connect()
    heizkoerper.PinnedRuecksetzen()
t.Commit()
