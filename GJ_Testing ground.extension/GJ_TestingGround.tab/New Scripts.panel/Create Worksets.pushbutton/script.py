# -*- coding: utf-8 -*-
"""Create Worksets from Text File"""

__title__ = 'Create Worksets'
__author__ = 'Goran Jovic'

# pyRevit modules
from pyrevit import forms

# Import Revit API modules
from Autodesk.Revit.DB import Workset, WorksetTable, FilteredWorksetCollector, WorksetKind, Transaction
from Autodesk.Revit.UI import TaskDialog

# Import common language runtime
import clr
import os

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

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
    t = Transaction(doc, 'Create Worksets from Text File')
    t.Start()

    new_worksets = []
    for name in workset_names:
        if name not in existing_workset_names:
            try:
                Workset.Create(doc, name)
                new_worksets.append(name)
            except Exception as e:
                print("Failed to create workset '{0}': {1}".format(name, e))

    t.Commit()

    # Show a message with the result
    if new_worksets:
        message = 'The following worksets have been created:\n\n' + '\n'.join(new_worksets)
    else:
        message = 'No new worksets were created. All worksets already exist or an error occurred.'

    TaskDialog.Show('Worksets Creation Result', message)
else:
    TaskDialog.Show('Operation Cancelled', 'No text file was selected.')
