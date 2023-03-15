# coding: utf8
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from Autodesk.Revit.DB import StorageType, Transaction
import pyapex_parameters as pyap
import pyapex_utils
from pyrevit import revit, UI, DB
from pyrevit.forms import WPFWindow, SelectFromList
from pyrevit import script, forms
import time
import System


start = time.time()
__context__ = 'Selection'

__title__ = "0.30 Parameter ändern"
__doc__ = """den Wert eines Parameters auf einen anderen setzen. Funktioniert mit ausgewählten Elementen. """
__author__ = "Menghui Zhang"

USE_NAMES = __shiftclick__

from pyIGF_logInfo import getlog
getlog(__title__)


logger = script.get_logger()

doc = revit.doc

global parametertoset
global parametertoget

class CheckBoxParameter:
    def __init__(self, parameter, default_state=False):
        self.parameter = parameter
        self.name = parameter.Definition.Name
        self.state = default_state

    def __str__(self):
        return self.name

    def __nonzero__(self):
        return self.state

    def __bool__(self):
        return self.state


class CopyParameterWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        self.elementid = revit.selection.get_selection()[0]
        self.element = doc.GetElement(self.elementid)
        if not self.selection:
            TaskDialog.Show(__title__, "Selection is empty")
            return

        self.parameters_dict = self.Parameters()
        self.parameters_editable = self.Parameter_edit()
        WPFWindow.__init__(self, xaml_file_name)
        self.parameterToGet.ItemsSource = sorted(self.parameters_dict.keys())
        self.parameterToSet.ItemsSource = sorted(self.parameters_editable.keys())
    
    def Parameters(self):
        param_dict = {}
        for param in self.element.Parameters:
            name = param.Definition.Name
            param_dict[name] = param
        return param_dict
    
    def Parameter_edit(self):
        param_dict = {}
        for el in param_dict.keys():
            if not self.Parameters()[el].IsReadOnly:
                if self.Parameters()[el].StorygeType.ToString() != 'ElementId':
                    param_dict[el] = self.Parameters()[el]
        return param_dict

    def run(self, sender, args):
        try:

            count_changed = 0
            # find not empty parameter
            definition_set = self.parameter_to_set.Definition
            definition_get = self.parameter_to_get.Definition

            # collect ones to be updated - parameters_get which aren't empty and aren't equal
            not_empty_list = []
            skip_ids = []
            errors_list_ids = []
            errors_list = []
            errors_text = ""
            for e in self.selection:
                if USE_NAMES:
                    param = e.LookupParameter(definition_set.Name)
                    param_get = e.LookupParameter(definition_get.Name)
                else:
                    param = e.get_Parameter(definition_set)
                    param_get = e.get_Parameter(definition_get)

                if not param or not param_get:
                    logger.debug("One of parameters not found for e.Id:%d" % e.Id.IntegerValue)
                    continue
                if not pyap.is_empty(param) and not pyap.are_equal(param, param_get):
                    value_get, value_set = pyap.convert_value(param_get, param, return_both=True)
                    not_empty_list.append("Target: %s, Source: %s" % (value_set, value_get))
                    skip_ids.append(e.Id)

            if len(not_empty_list) > 0:
                len_limit = 10
                len_not_empty_list = len(not_empty_list)
                if len_not_empty_list > len_limit:
                    not_empty_list = not_empty_list[:len_limit] + [' + %d more...' % (len_not_empty_list - len_limit)]

                text = "%d elements have values already. Replace them?\n" % len_not_empty_list + "\n".join(
                    not_empty_list)
                a = TaskDialog.Show(__title__, text,
                                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No)
                if a == TaskDialogResult.Yes:
                    skip_ids = []

            t = Transaction(doc, __title__)
            t.Start()
            for e in self.selection:
                if USE_NAMES:
                    param = e.LookupParameter(definition_set.Name)
                    param_get = e.LookupParameter(definition_get.Name)
                    if not param or not param_get:
                        logger.debug("One of parameters not found for e.Id:%d" % e.Id.IntegerValue)
                        continue
                else:
                    param = e.get_Parameter(definition_set)
                    param_get = e.get_Parameter(definition_get)
                _definition_set = param.Definition
                _definition_get = param_get.Definition

                if e.Id in skip_ids:
                    continue
                try:
                    if pyap.copy_parameter(e, _definition_get, _definition_set):
                        count_changed += 1
                except Exception as exc:
                    errors_list_ids.append(e.Id)
                    errors_list.append("Id: %d, exception: %s" % (e.Id.IntegerValue, str(exc)))
            if errors_list:
                errors_text = ("\n\nErrors occurred with %d elements :\n" % len(errors_list)) + \
                              "\n".join(errors_list[:5])
                if len(errors_list) > 5:
                    errors_text += "\n..."
                errors_text += "\n\nValues weren't changed, elements with errors selected"
                selection.set_to(errors_list_ids)

            if count_changed or errors_list:
                t.Commit()
                TaskDialog.Show(__title__,
                                "%d of %d elements updated%s" % (count_changed, len(self.selection.ToElementIds()), errors_text))
            else:
                t.RollBack()
                TaskDialog.Show(__title__, "Nothing was changed")
            logger.debug("finished")
        except Exception as exc:
            logger.error(exc)

        # TODO FIX do not write config in lower versions - risk to corrupt config
        if pyRevitNewer4619 or (
                pyapex_utils.is_ascii(self.parameterToGet.Text) and
                pyapex_utils.is_ascii(self.parameterToSet.Text)):
            try:
                self.write_config()
            except:
                logger.warn("Cannot save config")
        self.Close()

    def element_parameter_dict(self, elements, parameter_to_sort):
        result = {}
        for e in elements:
            if type(parameter_to_sort) == str:
                if parameter_to_sort[0] == "<" and parameter_to_sort[2:] == " coordinate>":
                    parameter_loc = parameter_to_sort[1]
                    loc = e.Location
                    v = getattr(loc.Point, parameter_loc)
                else:
                    logger.error("Parameter error")
                    return
            else:
                param = e.get_Parameter(parameter_to_sort.Definition)
                v = pyap.parameter_value_get(param)
            if v:
                result[e] = v

        return result


    def parameter_key(self, parameter):
        if USE_NAMES:
            p_key = parameter.Definition.Name
        else:
            p_key = parameter.Definition.Id
        return p_key


    def get_selection_parameters(self, elements):
        """
        Get parameters which are common for all selected elements

        :param elements: Elements list
        :return: dict - {parameter_name: Parameter}
        """
        result = {}
        all_parameter_set = set()
        all_parameter_dict = dict()
        all_parameter_ids_by_element = {}

        # find all ids
        for e in elements:
            for p in e.Parameters:
                p_key = self.parameter_key(p)
                if p_key not in all_parameter_dict.keys():
                    all_parameter_dict[p_key] = p

                if e.Id not in all_parameter_ids_by_element.keys():
                    all_parameter_ids_by_element[e.Id] = set()
                all_parameter_ids_by_element[e.Id].add(p_key)
            break

        # filter
        if USE_NAMES: # do not filter for names
            for p_key, p in all_parameter_dict.items():
                result[p.Definition.Name] = p
        else:
            for p_key, p in all_parameter_dict.items():
                exists_for_all_elements = True
                for e_id, e_params in all_parameter_ids_by_element.items():
                    if p_key not in e_params:
                        exists_for_all_elements = False
                        break

                if exists_for_all_elements:
                    result[p.Definition.Name] = p

        return result

    def filter_editable(self, parameters):
        """
        Filter parameters which can be modified by users

        :param parameters: list of parameters
        :return: filtered list of parameters
        """
        ignore_types = [StorageType.ElementId, ]
        result = {n: p for n, p in parameters.iteritems()
                  if not p.IsReadOnly
                  and p.StorageType not in ignore_types}
        return result

    def select_parameters(self):
        parameters_dict = self.get_selection_parameters(self.selection)
        parameters_editable = self.filter_editable(parameters_dict)

        options = []
        for param_id, param in parameters_editable.items():
            cb = CheckBoxParameter(param)
            options.append(cb)

        selected = SelectFromList.show(options, title='Parameter to replace', width=300,
                                       button_name='OK')

        return selected

def all_categories():
    outCate = []
    Categories = System.Enum.GetValues(DB.BuiltInCategory.OST_Levels.GetType())
    for item in Categories:
        try:
            name = DB.LabelUtils.GetLabelFor(item)
            outCate.append([name,item])
        except:
            pass
    return outCate
def main():
    # Input
    sel = selection.elements
    element = None
    for i in sel:
        element = i

    allCate = all_categories()
    cateName = element.Category.Name
    builtincate = None
    for item in allCate:
        if item[0] == cateName:
            builtincate = item[1]
            break
    collector = DB.FilteredElementCollector(doc).OfCategory(builtincate)\
        .WhereElementIsNotElementType()

    CopyParameterWindow('window.xaml', collector).ShowDialog()


if __name__ == "__main__":
    main()
