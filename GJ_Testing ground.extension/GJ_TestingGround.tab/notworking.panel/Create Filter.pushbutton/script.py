__doc__ = """
Title: PanelAutoFilter
Author: Goran Jovic
Version: 1.0
Description: Automatsko pravljenje filtera za panele.
"""
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
view = __revit__.ActiveUIDocument.ActiveView

def get_solid_fill_pattern(doc):
    return next((fp for fp in FilteredElementCollector(doc).OfClass(FillPatternElement) if fp.GetFillPattern().IsSolidFill), None)

def get_parameter_id_by_name(doc, param_name):
    for param in FilteredElementCollector(doc).OfClass(ParameterElement):
        if param.Name == param_name:
            return param.Id
    return None

def create_filter_and_set_overrides(doc, view, filter_name, color, solid_fill_pattern, parameter_id):
    existing_filter = [e.Id for e in FilteredElementCollector(doc).OfClass(ParameterFilterElement) if e.Name == filter_name]
    if existing_filter:
        doc.Delete(existing_filter[0])

    categories = List[ElementId]([ElementId(BuiltInCategory.OST_CurtainWallPanels)])
    filterElement = ParameterFilterElement.Create(doc, filter_name, categories)
    rule = FilterStringRule(ParameterValueProvider(parameter_id), FilterStringEquals(), filter_name, False)
    filterElement.SetElementFilter(ElementParameterFilter(rule))
    view.AddFilter(filterElement.Id)

    ovrSettings = OverrideGraphicSettings()
    for method in ['SetSurfaceForegroundPatternId', 'SetSurfaceBackgroundPatternId', 'SetCutForegroundPatternId', 'SetCutBackgroundPatternId']:
        getattr(ovrSettings, method)(solid_fill_pattern.Id)
    for method in ['SetSurfaceForegroundPatternColor', 'SetSurfaceBackgroundPatternColor', 'SetCutForegroundPatternColor', 'SetCutBackgroundPatternColor']:
        getattr(ovrSettings, method)(color)

    view.SetFilterOverrides(filterElement.Id, ovrSettings)

check_status_param_id = get_parameter_id_by_name(doc, "CHECK_STATUS")

if view.ViewTemplateId.IntegerValue == -1 and check_status_param_id is not None:
    t = Transaction(doc, 'Create Filters')
    t.Start()

    solid_fill_pattern = get_solid_fill_pattern(doc)
    if solid_fill_pattern is None:
        print("Solid fill pattern not found. Add logic to create one if needed.")
    else:
        filters_and_colors = [("CHECK HEIGHT", Color(255, 0, 0)), ("CHECK WIDTH", Color(0, 255, 64)), ("CHECK ALL", Color(0, 255, 255))]
        for filter_name, color in filters_and_colors:
            create_filter_and_set_overrides(doc, view, filter_name, color, solid_fill_pattern, check_status_param_id)

    t.Commit()
else:
    print("CREATION OF FILTERS NOT POSSIBLE WHILE VIEW TEMPLATE IS APPLIED. PLEASE CHECK IF YOU ARE ON CORRECT WORKING VIEW.")
