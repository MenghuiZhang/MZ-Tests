# coding: utf8
from IGF_Klasse import ItemTemplateMitElem,DB,ItemTemplateMitElemUndName,ObservableCollection,List
import Autodesk.Revit.UI as UI
from IGF_Funktionen._Geometrie import get_ClosestPoints
import clr

class Material(ItemTemplateMitElem):
    def __init__(self,elem):
        ItemTemplateMitElem.__init__(self,elem)
    
    def Set_GrafischeFarbe(self,Farbe):
        AppearanceAsset = DB.AppearanceAssetElement.GetAppearanceAssetElementByName(self.doc, self.name)
        if not AppearanceAsset:
            temp = DB.FilteredElementCollector(self.doc).OfClass(clr.GetClrType(DB.AppearanceAssetElement)).ToElements()[0]
            AppearanceAsset = temp.Duplicate(self.name)
        
        appearanceAssetEditScope = DB.Visual.AppearanceAssetEditScope(self.doc)
        editableAsset = appearanceAssetEditScope.Start(AppearanceAsset.Id)
        try:editableAsset.FindByName("generic_is_metal").Value = False
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_reflectivity_at_0deg").Value = 0
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_reflectivity_at_90deg").Value = 0
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_transparency_image_fade").Value = 1
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_diffuse_image_fade").Value = 0
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_transparency").Value = 0
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_glossiness").Value = 0
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_self_illum_filter_map").SetValueAsColor(Farbe)
        except Exception as e:print(e)
        try:editableAsset.FindByName("generic_diffuse").SetValueAsColor(Farbe)
        except Exception as e:print(e)
        appearanceAssetEditScope.Commit(True)
        appearanceAssetEditScope.Dispose()
        self.elem.AppearanceAssetId = AppearanceAsset.Id
        self.elem.UseRenderAppearanceForShading = True
        
    def Set_Thermal_Structural(self,TherId,StruId):
        try:
            if TherId:self.elem.ThermalAssetId = TherId
        except:pass
        try:
            if StruId:self.elem.StructuralAssetId = StruId
        except:pass

    def Set_MaterialFarbe(self,Farbe):
        self.elem.Color = Farbe
        self.elem.SurfaceForegroundPatternColor = Farbe
        self.elem.Transparency = 0
        # self.elem.UseRenderAppearanceForShading = False

