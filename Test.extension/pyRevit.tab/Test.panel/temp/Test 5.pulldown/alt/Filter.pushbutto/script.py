# coding: utf8
from pyrevit import revit, UI, DB, script, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.ObjectModel import ObservableCollection
import clr
clr.AddReference('Fdata.DBD.BIM.Revit.Api')
from pyrevit.forms import WPFWindow

__title__ = "Fliter"
__doc__ = """element filtern"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
Element_config = script.get_config()

uidoc = revit.uidoc
doc = revit.doc

selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]

class ELEMENT(object):
    @property
    def KategoName(self):
        return self._KategoName
    @KategoName.setter
    def KategoName(self, value):
        self._KategoName = value
    @property
    def FamilieName(self):
        return self._FamilieName
    @FamilieName.setter
    def FamilieName(self, value):
        self._FamilieName = value
    @property
    def TypName(self):
        return self._TypName
    @TypName.setter
    def TypName(self, value):
        self._TypName = value
    @property
    def ElementId(self):
        return self._ElementId
    @ElementId.setter
    def ElementId(self, value):
        self._ElementId = value


Liste_Element = ObservableCollection[ELEMENT]()

# public class ParamFilterTest : IExternalCommand
# {
#   public Result Execute(
#     ExternalCommandData commandData,
#     ref string message,
#     ElementSet elements )
#   {
#     UIApplication uiapp = commandData.Application;
#     UIDocument uidoc = uiapp.ActiveUIDocument;
#     Application app = uiapp.Application;
#     Document doc = uidoc.Document;
#
#     Wall wall = uidoc.Selection.PickObject(
#       Autodesk.Revit.UI.Selection.ObjectType.Element )
#       .Element as Wall;
#
#     Parameter parameter = wall.get_Parameter(
#       "Unconnected Height" );
#
#     ParameterValueProvider pvp
#       = new ParameterValueProvider( parameter.Id );
#
#     FilterNumericRuleEvaluator fnrv
#       = new FilterNumericGreater();
#
#     FilterRule fRule
#       = new FilterDoubleRule( pvp, fnrv, 20, 1E-6 );
#
#     ElementParameterFilter filter
#       = new ElementParameterFilter( fRule );
#
#     FilteredElementCollector collector
#       = new FilteredElementCollector( doc );
#
#     // Find walls with unconnected height
#     // less than or equal to 20:
#
#     ElementParameterFilter lessOrEqualFilter
#       = new ElementParameterFilter( fRule, true );
#
#     IList<Element> lessOrEqualFounds
#       = collector.WherePasses( lessOrEqualFilter )
#         .OfCategory( BuiltInCategory.OST_Walls )
#         .OfClass( typeof( Wall ) )
#         .ToElements();
#
#     TaskDialog.Show( "Revit", "Walls found: "
#       + lessOrEqualFounds.Count );
#
#     return Result.Succeeded;
#   }
# }

def DataPruefen(element):
    if not Testclass.HasDBDBIMData(element):
        temp_class = ELEMENT()
        cate = element.Category.Name
        familie = element.Parameter[BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString()
        typ = element.Parameter[BuiltInParameter.ELEM_TYPE_PARAM].AsValueString()
        temp_class.FamilieName = familie
        temp_class.KategoName = cate
        temp_class.TypName = typ
        temp_class.ElementId = element.Id.ToString()
        Liste_Element.Add(temp_class)


for el in selection:
    BIMDataPruefen(el)

class UnbemusterteElem(WPFWindow):
    def __init__(self, xaml_file_name,liste):
        self.Liste = liste
        WPFWindow.__init__(self, xaml_file_name)
        self.dataGrid.ItemsSource = liste

    def show(self,sender,args):
        import rpw
        from Autodesk.Revit.DB import ElementId

        uidoc = rpw.revit.uidoc
        mySelectedElement = self.dataGrid.SelectedItem
        id = ElementId(int(mySelectedElement.ElementId))
        uidoc.ShowElements(id)

        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(id)
        uidoc.Selection.SetElementIds(sel)


FamilienDialog = UnbemusterteElem('Pruefen_BIM_Data.xaml', Liste_Element)
FamilienDialog.Show()
