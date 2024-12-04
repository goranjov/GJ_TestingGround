# -*- coding: utf-8 -*-
__title__ = 'User View Creator'
__author__ = 'Goran Jovic'
__doc__ = 'Creates sets of views for selected users from usernames.txt'

import sys
import os
import clr

from pyrevit import revit, DB, forms

# Get the active document and application
doc = revit.doc
uidoc = revit.uidoc
app = uidoc.Application.Application  # Corrected application access

# Get the directory of the script
script_directory = os.path.dirname(__file__)

# Path to the TXT file containing usernames
txt_file_path = os.path.join(script_directory, 'usernames.txt')

# Read the usernames from the TXT file
try:
    with open(txt_file_path, 'r') as file:
        all_usernames = [line.strip() for line in file if line.strip()]
except Exception as e:
    forms.alert('Failed to read usernames.txt file.\n{}'.format(e), exitscript=True)

if not all_usernames:
    forms.alert('The usernames.txt file is empty.', exitscript=True)

# Present a selection window to the user to choose which users to process
selected_usernames = forms.SelectFromList.show(
    sorted(all_usernames),
    title='Select Users to Create Views For',
    multiselect=True,
    button_name='Create Views'
)

if not selected_usernames:
    forms.alert('No users selected. Exiting script.', exitscript=True)

# Collect all views in the document
try:
    collector = DB.FilteredElementCollector(doc)
    views = collector.OfClass(DB.View).ToElements()
    example_views = []
    for view in views:
        # Exclude view templates
        if not view.IsTemplate and view.LookupParameter('View Category'):
            # Check if 'View Category' equals '00_Example'
            if view.LookupParameter('View Category').AsString() == '00_Example':
                example_views.append(view)
except Exception as e:
    forms.alert('Failed to collect example views.\n{}'.format(e), exitscript=True)

if not example_views:
    forms.alert('No example views with View Category "00_Example" found.', exitscript=True)

# Check if required parameters exist in the views
param_names = ['Under-Discipline', 'View Category']
missing_params = []
for param_name in param_names:
    param = example_views[0].LookupParameter(param_name)
    if not param:
        missing_params.append(param_name)

if missing_params:
    forms.alert('Missing parameters in the model: {}'.format(', '.join(missing_params)), exitscript=True)

# Start a transaction to make changes to the document
t = DB.Transaction(doc, 'Duplicate Views for Users')
t.Start()

try:
    # Keep track of existing 'View Category' values to avoid duplicates
    existing_view_categories = set()
    for view in views:
        param = view.LookupParameter('View Category')
        if param:
            existing_value = param.AsString()
            if existing_value:
                existing_view_categories.add(existing_value)

    skipped_users = []

    for username in selected_usernames:
        # Get initials from the username
        name_parts = username.split('_')
        if len(name_parts) >= 2:
            full_name = name_parts[1]
            name_tokens = full_name.split()
            # Take the first letters of up to three name tokens
            initials = ''.join([token[0] for token in name_tokens][:3]).upper()
        else:
            initials = 'XXX'  # Default initials if format is unexpected

        # Check if views for this user already exist
        if username in existing_view_categories:
            skipped_users.append(username)
            continue  # Skip to the next user

        for view in example_views:
            # Handle views with '{}' in their names
            if '{}' in view.Name:
                view.Name = view.Name.replace('{}', '*')

            # Duplicate the view with detailing
            try:
                new_view_id = view.Duplicate(DB.ViewDuplicateOption.WithDetailing)
                new_view = doc.GetElement(new_view_id)

                # Replace 'Copy 1' in the view name with user's initials
                new_view_name = new_view.Name.replace('Copy 1', initials)
                new_view.Name = new_view_name

                # Set the required parameters
                new_view.LookupParameter('Under-Discipline').Set('01_Work In Progress')
                new_view.LookupParameter('View Category').Set(username)
            except Exception as e:
                forms.alert('Failed to duplicate view "{}".\n{}'.format(view.Name, e))

    t.Commit()

except Exception as e:
    t.RollBack()
    forms.alert('An error occurred: {}'.format(e), exitscript=True)

# Inform the user about any skipped users
if skipped_users:
    message = 'Views for the following users already exist and were skipped:\n' + '\n'.join(skipped_users)
    forms.alert(message)
else:
    forms.alert('Views created successfully.')
