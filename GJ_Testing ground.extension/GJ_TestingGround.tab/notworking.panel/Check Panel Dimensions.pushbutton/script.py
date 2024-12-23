__doc__ = """
Title: Panel Visualisation
Author: Goran Jovic
Version: 1.0
Description: Upisacemo komentar u panel ako dimenzije panela nisu "lepe brojke".
"""
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction

doc = __revit__.ActiveUIDocument.Document
view = __revit__.ActiveUIDocument.ActiveView

# Collect curtain panels in the current view
curtain_panels = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_CurtainWallPanels).ToElements()

# Transaction to modify elements
t = Transaction(doc, 'Update Comments on Curtain Panels')
t.Start()

modified_count = 0
check_status_found = False

for panel in curtain_panels:
    comments_to_add = []

    # Check height parameter
    height_param = panel.LookupParameter('Height')
    if height_param:
        height_value = height_param.AsValueString()
        if not (height_value.endswith('5') or height_value.endswith('0') or height_value.endswith('0.1')):
            comments_to_add.append('CHECK HEIGHT')

    # Check width parameter
    width_param = panel.LookupParameter('Width')
    if width_param:
        width_value = width_param.AsValueString()
        if not (width_value.endswith('5') or width_value.endswith('0')):
            comments_to_add.append('CHECK WIDTH')

    # Update CHECK_STATUS parameter
    comments_param = panel.LookupParameter('CHECK_STATUS')
    if comments_param:
        check_status_found = True
        new_comment = 'CHECK ALL' if len(comments_to_add) == 2 else ', '.join(comments_to_add)
        new_comment = new_comment if comments_to_add else ''
        old_comment = comments_param.AsString()
        if old_comment != new_comment:
            comments_param.Set(new_comment)
            modified_count += 1

t.Commit()

# Print the number of modified elements
if check_status_found:
    print('Operation complete. {} elements were modified.'.format(modified_count))
else:
    print("Dodajte shared parameter CHECK_STATUS za Curtain Panels da bi skripta pravilno radila")
