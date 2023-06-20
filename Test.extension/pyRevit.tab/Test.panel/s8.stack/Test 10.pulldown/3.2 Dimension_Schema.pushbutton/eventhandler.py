# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
from System.Collections.Generic import List
from pyrevit import script
from System.Windows import Visibility
import System
from pyrevit import script
from System.Windows.Forms import OpenFileDialog,DialogResult
import os
from excel._EPPlus import FileAccess,FileMode,FileStream,ExcelPackage,LicenseContext

logger = script.get_logger()
Visible = Visibility.Visible
Hidden = Visibility.Hidden

doc = __revit__.ActiveUIDocument.Document
projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number
config = script.get_config('Schema-TS-Param-Schema-Rohrzubehoer -' + projectinfo)

_hls = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
_params = []
for el in _hls:
    params = el.Parameters
    for pa in params:
        try:
            name = pa.Definition.Name
            if name not in _params:_params.append(name)
        except:
            pass
    break
_params.sort()


class Excel_Daten:
    def __init__(self,Id,BauteilId,Dimension):
        self.Id = Id
        self.BauteilId = BauteilId
        self.Dimension = Dimension

class FilterNeben(Selection.ISelectionFilter):
    def __init__(self,liste):
        self.Liste = liste

    def AllowElement(self,element):
        if element.Id.ToString() in self.Liste:
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class Filter(Selection.ISelectionFilter):

    def AllowElement(self,element):
        if element.Category.Id.ToString() == '-2008044':
            return True
        else:
            return False
    def AllowReference(self,reference,XYZ):
        return False

class AlleBauteile:
    def __init__(self,Liste,elem):
        self.elem = elem
        self.rohrliste = Liste
        self.liste = []
        self.hls_liste = []
        self.rohr = []
        self.formteil = []
        self.enddeckel = []
        self.zubehoer = []
        self.flexrohe = []
        for el in self.rohrliste:
            self.liste.append(el)
            self.rohr.append(el)
        if self.elem.Category.Id.ToString() == '-2008044':
            self.rohr.append(self.elem.Id.ToString())
        elif self.elem.Category.Id.ToString() == '-2008049':
            self.formteil.append(self.elem.Id.ToString())
        elif self.elem.Category.Id.ToString() == '-2008055':
            self.zubehoer.append(self.elem.Id.ToString())  
        self.get_t_st(self.elem)
        self.alle = []
        self.alle.extend(self.rohr)
        self.alle.extend(self.zubehoer)
        self.alle.extend(self.formteil)
        self.alle.extend(self.enddeckel)
        self.alle.extend(self.flexrohe)

    def get_t_st(self,elem):       
        elemid = elem.Id.ToString()
        self.liste.append(elemid)
        cate = elem.Category.Name

        if not cate in ['Rohr Systeme','Rohrdämmung']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                for conn in conns:
                    refs = conn.AllRefs
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste:  
                            category_temp = owner.Category.Id.ToString()
                            if category_temp == '-2001140':
                                self.hls_liste.append(owner.Id.ToString())
                                return
                            elif category_temp == '-2008044':
                                self.rohr.append(owner.Id.ToString())
                            elif category_temp == '-2008049':
                                conn_temp = owner.MEPModel.ConnectorManager.Connectors
                                if conn_temp.Size == 1:self.enddeckel.append(owner.Id.ToString())
                                else:self.formteil.append(owner.Id.ToString())
                            elif category_temp == '-2008055':
                                self.zubehoer.append(owner.Id.ToString())        
                            elif category_temp == '-2008050':
                                self.flexrohe.append(owner.Id.ToString())             
                            self.get_t_st(owner)

class Rohrzubehoer:
    def __init__(self,elem,param):
        self.elem = elem
        self.param = param
        self.liste = []
        self.nr_liste = []
        self.get_Nr(self.elem)
    @property
    def Dimension(self):
        try:
            conns = self.elem.MEPModel.ConnectorManager.Connectors
            for conn in conns:
                return int(round(conn.Radius*304.8*2))
        except:
            return 0
    
    @property
    def ID(self):
        param = self.elem.LookupParameter(self.param)
        if param:
            if param.AsString():
                return param.AsString()
            else:
                return ''
        else:
            return ''


    def get_Nr(self,elem):
        if len(self.nr_liste) == 2:
            return 
        elemid = elem.Id.ToString()
        self.liste.append(elemid)

        cateid = elem.Category.Id.ToString()
        
        if cateid not in ['-2008043','-2008122']:
            conns = None
            try:conns = elem.MEPModel.ConnectorManager.Connectors
            except:
                try:conns = elem.ConnectorManager.Connectors
                except:pass
            if conns:
                if conns.Size > 2:return
                for conn in conns:
                    refs = conn.AllRefs
                 
                    for ref in refs:
                        owner = ref.Owner
                        if owner.Id.ToString() not in self.liste:  
                            if owner.Category.Id.ToString() == '-2008044': 
                                nr = owner.LookupParameter('IGF_X_RVT_TS_Nr').AsString()
                                if nr:
                                    nr = int(float(nr))
                                    self.nr_liste.append(nr)
                            else:
                                self.get_Nr(owner)
    
    @property
    def Text(self):
        if len(self.nr_liste) == 2:
            self.nr_liste.sort()
            text = str(self.nr_liste[0]) + str(self.nr_liste[1])
            return text
        elif len(self.nr_liste) == 1:
            return str(self.nr_liste[0])+ str(self.nr_liste[0])
        else:
            print(self.elem.Id)
            return ''
    
    def wert_schreiben(self,value):
        try:self.elem.LookupParameter('MC Diameter Instance').SetValueString(str(round(float(value))))
        except Exception as e:print(e)
                                       
class Externalliste(IExternalEventHandler):
    def __init__(self):
        self.GUI = None
        self.ExecuteApp = None
        self.name = ''
        self.bauteillistename = 'IGF_X_Information ans Schema_Rohre'
    
    def Execute(self,app):
        try:
            self.ExecuteApp(app)
        except Exception as e:
            print(e)
    
    def GetName(self):
        return self.name
    
    def select(self,uiapp):
        self.name = 'Strang auswählen'
        if self.GUI.btid.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Bitte BauteilId-Parameter auswählen')
            return
        try:
            self.GUI.write_config()
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document
            el0_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element,Filter(),'Wählt ein Rohr aus')
            el0 = doc.GetElement(el0_ref)
            conns = el0.ConnectorManager.Connectors
            liste = []
            Liste_Rohr = []
            Liste_Rohr.append(el0.Id.ToString())
            for conn in conns:
                if conn.IsConnected:
                    refs = conn.AllRefs
                    for _ref in refs:
                        if _ref.Owner.Category.Name not in ['Rohr Systeme','Rohrdämmung'] and _ref.Owner.Id.ToString() != el0.Id.ToString():
                            liste.append(_ref.Owner.Id.ToString())

            if len(liste) ==  0:
                TaskDialog.Show('Info','kein Bauteil damit verbunden!')
            ref_1 = uidoc.Selection.PickObject(Selection.ObjectType.Element,FilterNeben(liste),'Wählt den nächsten Teil aus')
            self.GUI.allebauteile = AlleBauteile(Liste_Rohr,doc.GetElement(ref_1))
            self.GUI.button_nummer.IsEnabled = True
            
            TaskDialog.Show('Info','Strang ausgewählt!')
        except Exception as e:print(e)

    def _update_pbar(self):
        self.GUI.pb01.Maximum = self.GUI.maxvalue
        self.GUI.pb01.Value = self.GUI.value
        self.GUI.pb_text.Text = self.GUI.PB_text
    
    def exprotieren(self,uiapp):
        self.GUI.write_config()
        self.name = "Daten übernehmen"
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        excel = self.GUI.excel.Text
        if os.path.exists(excel):
            Parameter_Dict = {}
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            fs = FileStream(excel,FileMode.Open,FileAccess.Read)
            book = ExcelPackage(fs)
            Bauteilnummer = ''
            try:
                for sheet in book.Workbook.Worksheets:
                    maxRowNum = sheet.Dimension.End.Row
                    Bauteilnummer = sheet.Cells[1, 1].Value
                    for row in range(2, maxRowNum + 1):
                        Id = sheet.Cells[row, 1].Value
                        bauteilId = sheet.Cells[row, 2].Value
                        dimenison = sheet.Cells[row, 3].Value
                        klass = Excel_Daten(Id,bauteilId,dimenison)
                        if Id not in Parameter_Dict.keys():
                            Parameter_Dict[Id] = []
                        Parameter_Dict[Id].append(klass)
            

            except Exception as e:
                logger.error(e)
                TaskDialog.Show('Fehler','Fehler beim Lesen der Excel-Datei')
                return
            
            book.Save()
            book.Dispose()
            fs.Dispose()
        
        try:
            self.GUI.write_config()
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document
            Ids = self.GUI.allebauteile.zubehoer[:]
            self.GUI.pb01.Visibility = Visible
            self.GUI.pb_text.Visibility = Visible

            t = DB.Transaction(doc,'Rohre')
            t.Start()
            for n,elid in enumerate(Ids):
                if n % 20  == 0:
                    self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                self.GUI.value = n+1
                self.GUI.maxvalue = len(Ids)
                self.GUI.PB_text = str(n+1) + ' / '+ str(len(Ids)) + ' Rohrzubehör'
                elem = Rohrzubehoer(doc.GetElement(DB.ElementId(int(elid))),self.GUI.config.param)
                test = False
                if elem.Text in Parameter_Dict.keys():
                    if len(Parameter_Dict[elem.Text]) == 1:
                        # print(Parameter_Dict[elem.Text][0].Dimension)
                        test = True
                        elem.wert_schreiben(Parameter_Dict[elem.Text][0].Dimension)
                    else:
                        for item in Parameter_Dict[elem.Text]:
                            if item.BauteilId == elem.ID or (item.BauteilId in [None,''] and elem.ID in [None,'']):
                                # print(item.Dimension)
                                elem.wert_schreiben(item.Dimension)
                                test = True
                if not test:
                    print(elem.Text,elid)


            self.GUI.maxvalue = 100
            self.GUI.value = 0
            self.GUI.PB_text = ''
            self.GUI.pb01.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
            
            self.GUI.button_nummer.IsEnabled = False
            t.Commit()
       

        except Exception as e:
            self.GUI.pb01.Visibility = Hidden
            self.GUI.pb_text.Visibility = Hidden
        self.GUI.pb01.Visibility = Hidden
        self.GUI.pb_text.Visibility = Hidden

    def ordnersclect(self,uiapp):
        self.name = 'Speicherort auswählen'
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "Excel"
        dialog.Filter = "Excel Dateien|*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            FileName = dialog.FileName
            self.GUI.excel.Text = FileName
        self.GUI.write_config()
 