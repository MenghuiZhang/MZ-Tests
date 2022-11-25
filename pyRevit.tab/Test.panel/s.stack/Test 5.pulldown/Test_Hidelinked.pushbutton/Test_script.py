# coding: utf8
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

__title__ = "Test Visibility"

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

vgCmdId = UI.RevitCommandId.LookupCommandId("ID_VIEW_CATEGORY_VISIBILITY")

uidoc.Application.PostCommand(vgCmdId)



a = '''Jrn.TabCtrl "Modal , Überschreibungen Sichtbarkeit/Grafiken für Grundriss: Abgabe_GA_Ebene E1 1zu200 , 0" _
          , "AFX_IDC_TAB_CONTROL" _
          , "Select" , "Modellkategorien"'+

 'Jrn.TabCtrl "Modal , Überschreibungen Sichtbarkeit/Grafiken für Grundriss: Abgabe_GA_Ebene E1 1zu200 , 0" _
          , "AFX_IDC_TAB_CONTROL" _
          , "Select" , "Revit-Verknüpfungen"

 Jrn.TreeCtrl "0" , "IDC_TREE" _
         ,"ChangeSelection" , ">>XA_2-NA.rvt>>"

 Jrn.TreeCtrl "0" , "IDC_TREE" _
         ,"ChangeSelection" , ">>XA_2-NA.rvt>>"

 Jrn.Grid "ChildControl; Page , Revit-Verknüpfungen , Dialog_Revit_RvtLinkVisibility; ID_TREEGRID_GRID" _
         , "MoveCurrentCell" , "1" , "Anzeigeeinstellungen"

 Jrn.Grid "ChildControl; Page , Revit-Verknüpfungen , Dialog_Revit_RvtLinkVisibility; ID_TREEGRID_GRID" _
         , "Button" , "1" , "Anzeigeeinstellungen"
  ' 0:< ::1213:: Delta VM: Avail -2 -> 134161590 MB, Used 15046 MB; RAM: Avail -11 -> 40990 MB, Used +12 -> 1175 MB 
  ' 0:< GUI Resource Usage GDI: Avail 8566, Used 1434, User: Used 680 

  Jrn.RadioButton "Page , Grundfunktionen , Dialog_Revit_RvtLinkDisplaySettingsBasics" _
         , "Benutzerdefiniert, Control_Revit_RadioOverridden"

 Jrn.TabCtrl "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "AFX_IDC_TAB_CONTROL" _
         , "Select" , "Grundfunktionen"

 Jrn.TabCtrl "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "AFX_IDC_TAB_CONTROL" _
         , "Select" , "Beschriftungskategorien"

 Jrn.ComboBox "Page , Beschriftungskategorien , Dialog_Revit_RvtLinkDisplaySettingsAnnotation" _
         , "Control_Revit_DisplayMethod" _
         , "SelEndOk" , "<Benutzerdefiniert>"

 Jrn.ComboBox "Page , Beschriftungskategorien , Dialog_Revit_RvtLinkDisplaySettingsAnnotation" _
         , "Control_Revit_DisplayMethod" _
         , "Select" , "<Benutzerdefiniert>"

 Jrn.TabCtrl "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "AFX_IDC_TAB_CONTROL" _
         , "Select" , "Beschriftungskategorien"

 Jrn.TabCtrl "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "AFX_IDC_TAB_CONTROL" _
         , "Select" , "Modellkategorien"

 Jrn.ComboBox "Page , Modellkategorien , Dialog_Revit_RvtLinkDisplaySettingsModel" _
         , "Control_Revit_DisplayMethod" _
         , "SelEndOk" , "<Benutzerdefiniert>"

 Jrn.ComboBox "Page , Modellkategorien , Dialog_Revit_RvtLinkDisplaySettingsModel" _
         , "Control_Revit_DisplayMethod" _
         , "Select" , "<Benutzerdefiniert>"

 Jrn.TreeCtrl "0" , "IDC_TREE" _
         ,"ChangeSelection" , ">>Rasterbilder>>"

  Jrn.TreeCtrl "0" , "IDC_TREE" _
          ,"ChangeSelection" , ">>Rasterbilder>>"

   Jrn.TreeCtrl "0" , "IDC_TREE" _
           ,"ToggleCheckBox [ ]" , ">>Rasterbilder>>"

 Jrn.TabCtrl "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "AFX_IDC_TAB_CONTROL" _
         , "Select" , "Modellkategorien"

 Jrn.TabCtrl "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "AFX_IDC_TAB_CONTROL" _
         , "Select" , "Beschriftungskategorien"

 Jrn.TreeCtrl "0" , "IDC_TREE" _
         ,"ChangeSelection" , ">>Raster>>"

 Jrn.TreeCtrl "0" , "IDC_TREE" _
         ,"ChangeSelection" , ">>Raster>>"

 Jrn.TreeCtrl "0" , "IDC_TREE" _
         ,"ToggleCheckBox [ ]" , ">>Raster>>"

 Jrn.PushButton "Modal , Anzeigeeinstellungen für RVT-Verknüpfungen , 0" _
         , "OK, IDOK"'''

# bericht = open(r"C:\Users\Zhang\AppData\Local\Autodesk\Revit\Autodesk Revit 2020\Journals\journal.0627.txt","a")
# bericht.write(a)
# bericht.close()