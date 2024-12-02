# -*- coding: utf-8 -*-
"""Create Worksets and Corresponding Views"""

__title__ = 'Create Worksets and Views'
__author__ = 'Goran Jovic'

# pyRevit modules
from pyrevit import forms

# Import Revit API modules
from Autodesk.Revit.DB import (
    Workset, FilteredWorksetCollector, WorksetKind, Transaction,
    ViewFamilyType, ViewFamily, FilteredElementCollector, View, WorksetId, WorksetVisibility,
    View3D, ElementId, BuiltInParameter
)

from Autodesk.Revit.UI import (
    UIDocument, UIApplication, TaskDialog
)

# Import common language runtime
import clr
import os

# Get the current Revit document and application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = __revit__.Application

# Function to create a view with only one workset visible
def create_view_for_workset(workset_name):
    # Start a transaction to create the view
    t_view = Transaction(doc, 'Create View for Workset: ' + workset_name)
    t_view.Start()
    
    try:
        # Get a ViewFamilyType for 3D views
        view_family_types = FilteredElementCollector(doc) \
            .OfClass(ViewFamilyType) \
            .ToElements()
        
        view_family_type_3d = None
        for vft in view_family_types:
            if vft.ViewFamily == ViewFamily.ThreeDimensional:
                view_family_type_3d = vft
                break
        
        if view_family_type_3d is None:
            raise Exception("No 3D ViewFamilyType found.")

        # Create a new 3D view
        new_view = View3D.CreateIsometric(doc, view_family_type_3d.Id)
        new_view.Name = 'RVT_HYG_' + workset_name

        # Collect all user worksets
        all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
        workset_dict = {ws.Name: ws.Id for ws in all_worksets}

        # Set visibility of all worksets to Hidden
        for ws_name, ws_id in workset_dict.items():
            new_view.SetWorksetVisibility(ws_id, WorksetVisibility.Hidden)

        # Set visibility of the target workset to Visible
        target_ws_id = workset_dict.get(workset_name)
        if target_ws_id:
            new_view.SetWorksetVisibility(target_ws_id, WorksetVisibility.Visible)
        else:
            print("Workset '{0}' not found.".format(workset_name))

        t_view.Commit()
    except Exception as e:
        print("Failed to create view for workset '{0}': {1}".format(workset_name, e))
        t_view.RollBack()

# Main execution
# Prompt the user to select the text file
txt_file_path = forms.pick_file(file_ext='txt', init_dir=os.path.expanduser('~'), multi_file=False)

if txt_file_path:
    # Read workset names from the text file
    with open(txt_file_path, 'r') as file:
        workset_names = [line.strip() for line in file if line.strip()]

    # Get existing user worksets
    existing_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    existing_workset_names = [ws.Name for ws in existing_worksets]

    # Start a transaction to create worksets
    t_create = Transaction(doc, 'Create Worksets from Text File')
    t_create.Start()

    new_worksets = []
    for name in workset_names:
        if name not in existing_workset_names:
            try:
                Workset.Create(doc, name)
                new_worksets.append(name)
            except Exception as e:
                print("Failed to create workset '{0}': {1}".format(name, e))

    t_create.Commit()

    # After creating worksets, create views for each
    for workset_name in new_worksets:
        create_view_for_workset(workset_name)

    # Show a message with the result
    if new_worksets:
        message = 'The following worksets have been created:\n\n' + '\n'.join(new_worksets)
    else:
        message = 'No new worksets were created. All worksets already exist or an error occurred.'

    TaskDialog.Show('Worksets Creation Result', message)
else:
    TaskDialog.Show('Operation Cancelled', 'No text file was selected.')
