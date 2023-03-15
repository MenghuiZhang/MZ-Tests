# coding: utf8
from IGF_log import getlog
from rpw import revit, UI, DB
from pyrevit import script, forms


__title__ = "0.02 Parameter ändern"
__doc__ = """
den Wert eines Parameters auf einen anderen setzen. Funktioniert mit Kategorie von ausgewählten Elementen.

[2021.10.08]
Version: 2.0
"""
__author__ = "Menghui Zhang"
__context__ = 'Selection'

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

try:
    getlog(__title__)
except:
    pass

class CopyParameterWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        self.selectedelemids = uidoc.Selection.GetElementIds()
        self.elem = doc.GetElement(list(self.selectedelemids)[0])
        self.Categoryid = self.elem.Category.Id
        self.Categoryname = self.elem.Category.Name
        self.coll = DB.FilteredElementCollector(doc).OfCategoryId(self.Categoryid).WhereElementIsNotElementType()
        self.elemids = self.coll.ToElementIds()
        self.coll.Dispose()

        self.Parameter_Get = self.get_Parameters()
        self.Parameter_Set = self.set_Parameters()
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.parameterToGet.ItemsSource = sorted(self.Parameter_Get.keys())
        self.parameterToSet.ItemsSource = sorted(self.Parameter_Set.keys())


    def get_value(self,param):
  
        """gibt den gesuchten Wert ohne Einheit zurück"""
        if param.StorageType.ToString() == 'ElementId':
            return param.AsValueString()
        elif param.StorageType.ToString() == 'Integer':
            value = param.AsInteger()
        elif param.StorageType.ToString() == 'Double':
            value = param.AsDouble()
        elif param.StorageType.ToString() == 'String':
            value = param.AsString()
            return value

        try:
            # in Revit 2020
            unit = param.DisplayUnitType
            value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
        except:
            try:
                # in Revit 2021/2022
                unit = param.GetUnitTypeId()
                value = DB.UnitUtils.ConvertFromInternalUnits(value,unit)
            except:
                pass


        return value
    
    def get_Parameters(self):
        param_dict = {}
        for param in self.elem.Parameters:
            param_dict[param.Definition.Name] = param

        return param_dict
    
    def set_Parameters(self):
        param_dict = {}
        for param in self.elem.Parameters:
            if not param.IsReadOnly:
                param_dict[param.Definition.Name] = param
        return param_dict
    
    def setparameterwert(self,param,wert):
        if wert is not None:
            if param.StorageType.ToString() == 'ElementId':
                pass
            elif param.StorageType.ToString() == 'Integer':
                param.Set(int(wert))
            elif param.StorageType.ToString() == 'Double':
                param.SetValueString(str(wert))
            elif param.StorageType.ToString() == 'String':
                param.Set(str(wert))
    
    def run(self, sender, args):
        if self.parameterToGet.SelectedIndex == -1 or self.parameterToSet.SelectedIndex == -1:
            UI.TaskDialog.Show(__title__, 'Parameter nicht ausgewählt')
            return
        paramget = self.parameterToGet.SelectedItem
        paramset = self.parameterToSet.SelectedItem
        self.Close()
        t = DB.Transaction(doc)
        t.Start('Parameter ändern')
        with forms.ProgressBar(title = "{value}/{max_value} "+self.Categoryname,cancellable=True, step=int(len(self.elemids)/1000)+10) as pb:
            for n, elem_id in enumerate(self.elemids):
                if pb.cancelled:
                    t.RollBack()
                    script.exit()

                pb.update_progress(n + 1, len(self.elemids))
                elem = doc.GetElement(elem_id)
                
                # try:
                
                # except:
                #     UI.TaskDialog.Show(__title__, 'Parameter nicht ausgewählt')
                #     return
                # param_get = self.Parameter_Get[paramget]
                # param_set = self.Parameter_Set[paramset]
                # if param_get.StorageType.ToString() != param_set.StorageType.ToString():
                #     UI.TaskDialog.Show(__title__, 'unvereinbare Parametertype')
                #     return
                try:
                    werte = self.get_value(elem.LookupParameter(paramget))
                    self.setparameterwert(elem.LookupParameter(paramset),werte)
                except Exception as e:
                    logger.error(e)
        t.Commit()
Parameter_WPF = CopyParameterWindow('window.xaml')
try:
    Parameter_WPF.ShowDialog()
except Exception as e1:
    logger.error(e1)
    Parameter_WPF.Close()