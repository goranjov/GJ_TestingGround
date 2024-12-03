# -*- coding: utf-8 -*-
__title__ = 'Copy Parameter Values'
__author__ = 'Assistant'
__doc__ = 'Copies parameter values from one parameter to another for selected elements with enhanced UI and error handling.'

import clr
import sys

# Add references to required .NET assemblies
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')

# Import Revit and .NET namespaces
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form, ListBox, Button, DialogResult, CheckBox, Label, DockStyle, SelectionMode, Panel
)

# Get the Revit application and document
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def get_all_parameters(elements, include_read_only=True):
    """Get a set of all parameter names from the selected elements."""
    param_info = {}  # Dictionary to store parameter names and their read-only status
    for elem in elements:
        # Get instance parameters
        for param in elem.Parameters:
            param_def = param.Definition
            param_name = param_def.Name
            is_read_only = param.IsReadOnly
            if param_name not in param_info or not is_read_only:
                param_info[param_name] = is_read_only
        # Get type parameters
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            for param in elem_type.Parameters:
                param_def = param.Definition
                param_name = param_def.Name
                is_read_only = param.IsReadOnly
                if param_name not in param_info or not is_read_only:
                    param_info[param_name] = is_read_only
    # Filter parameters based on read-only status
    if not include_read_only:
        param_info = {k: v for k, v in param_info.items() if not v}
    # Return sorted list of parameter names
    return sorted(param_info.keys()), param_info

def prompt_for_parameter(param_list, param_info, title):
    """Prompt the user to select a parameter from the list."""
    class ParameterForm(Form):
        def __init__(self, param_list, param_info, title):
            self.Text = title
            # Adjust the form size as needed
            self.Width = 500  # <-- Adjust form width here
            self.Height = 600  # <-- Adjust form height here

            self.param_info = param_info  # Store parameter info for filtering
            self.all_params = param_list  # Original list of parameters

            self.label = Label()
            self.label.Text = title
            self.label.Dock = DockStyle.Top
            self.Controls.Add(self.label)

            # Panel to contain the ListBox and enable scrolling
            self.panel = Panel()
            self.panel.Dock = DockStyle.Fill
            self.panel.AutoScroll = True

            self.listbox = ListBox()
            self.listbox.SelectionMode = SelectionMode.One  # Use the enumeration value
            # Adjust the ListBox size as needed
            self.listbox.Width = 460  # <-- Adjust ListBox width here
            self.listbox.Height = 400  # <-- Adjust ListBox height here
            self.listbox.Top = 30  # <-- Adjust ListBox top position if needed

            # Populate ListBox with parameter names
            self.populate_listbox(param_list)

            self.panel.Controls.Add(self.listbox)
            self.Controls.Add(self.panel)

            self.show_read_only_check = CheckBox()
            self.show_read_only_check.Text = "Show Read-Only Parameters"
            self.show_read_only_check.Checked = True  # Default to show all parameters
            self.show_read_only_check.Dock = DockStyle.Bottom
            self.show_read_only_check.CheckedChanged += self.toggle_read_only
            self.Controls.Add(self.show_read_only_check)

            self.convert_check = CheckBox()
            self.convert_check.Text = "Convert numerical values to text"
            self.convert_check.Dock = DockStyle.Bottom
            self.Controls.Add(self.convert_check)

            self.ok_button = Button()
            self.ok_button.Text = "OK"
            self.ok_button.DialogResult = DialogResult.OK
            self.ok_button.Dock = DockStyle.Bottom
            self.Controls.Add(self.ok_button)

            self.cancel_button = Button()
            self.cancel_button.Text = "Cancel"
            self.cancel_button.DialogResult = DialogResult.Cancel
            self.cancel_button.Dock = DockStyle.Bottom
            self.Controls.Add(self.cancel_button)

            self.AcceptButton = self.ok_button
            self.CancelButton = self.cancel_button

        def populate_listbox(self, param_list):
            self.listbox.Items.Clear()
            for param_name in param_list:
                if self.param_info[param_name]:
                    display_name = param_name + " (Read-Only)"
                else:
                    display_name = param_name
                self.listbox.Items.Add(display_name)

        def toggle_read_only(self, sender, event):
            include_read_only = self.show_read_only_check.Checked
            if include_read_only:
                filtered_params = self.all_params
            else:
                filtered_params = [p for p in self.all_params if not self.param_info[p]]
            self.populate_listbox(filtered_params)

        @property
        def SelectedParameter(self):
            selected_item = self.listbox.SelectedItem
            if selected_item:
                # Remove "(Read-Only)" from parameter name if present
                return selected_item.replace(" (Read-Only)", "")
            else:
                return None

        @property
        def ConvertValues(self):
            return self.convert_check.Checked

    form = ParameterForm(param_list, param_info, title)
    result = form.ShowDialog()

    if result == DialogResult.OK and form.SelectedParameter:
        return form.SelectedParameter, form.ConvertValues
    else:
        return None, False

def main():
    # Get selected elements
    selection_ids = uidoc.Selection.GetElementIds()
    if not selection_ids:
        TaskDialog.Show("Copy Parameter Values", "Please select at least one element.")
        return

    elements = [doc.GetElement(id) for id in selection_ids]

    # Get all parameters from selected elements
    all_params, param_info = get_all_parameters(elements, include_read_only=True)

    if not all_params:
        TaskDialog.Show("Copy Parameter Values", "No parameters found in selected elements.")
        return

    # Prompt for source parameter
    src_param_name, convert_values = prompt_for_parameter(all_params, param_info, "Select Source Parameter")
    if not src_param_name:
        return

    # Prompt for destination parameter
    dest_param_name, _ = prompt_for_parameter(all_params, param_info, "Select Destination Parameter")
    if not dest_param_name:
        return

    # Start a transaction
    t = Transaction(doc, "Copy Parameter Values")
    t.Start()

    error_log = []
    elements_processed = 0  # Counter for elements successfully processed

    for elem in elements:
        elem_name = elem.Name if hasattr(elem, 'Name') else "Unnamed Element"
        elem_id = elem.Id.IntegerValue

        # Initialize parameters
        src_param = elem.LookupParameter(src_param_name)
        dest_param = elem.LookupParameter(dest_param_name)

        # Check in type parameters if instance parameters are not found
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            if not src_param:
                src_param = elem_type.LookupParameter(src_param_name)
            if not dest_param:
                dest_param = elem_type.LookupParameter(dest_param_name)

        if not src_param:
            error_log.append("Element ID {} ('{}'): Source parameter '{}' not found.".format(elem_id, elem_name, src_param_name))
            continue
        if not dest_param:
            error_log.append("Element ID {} ('{}'): Destination parameter '{}' not found.".format(elem_id, elem_name, dest_param_name))
            continue

        if src_param.IsReadOnly:
            error_log.append("Element ID {} ('{}'): Source parameter '{}' is read-only.".format(elem_id, elem_name, src_param_name))
            continue
        if dest_param.IsReadOnly:
            error_log.append("Element ID {} ('{}'): Destination parameter '{}' is read-only.".format(elem_id, elem_name, dest_param_name))
            continue

        try:
            # Get the source value based on its storage type
            if src_param.StorageType == StorageType.String:
                src_value = src_param.AsString()
            elif src_param.StorageType == StorageType.Integer:
                src_value = src_param.AsInteger()
            elif src_param.StorageType == StorageType.Double:
                src_value = src_param.AsDouble()
            elif src_param.StorageType == StorageType.ElementId:
                src_value = src_param.AsElementId()
            else:
                src_value = None

            if src_value is None:
                error_log.append("Element ID {} ('{}'): Source parameter '{}' has no value.".format(elem_id, elem_name, src_param_name))
                continue

            # Set the destination value based on its storage type
            if dest_param.StorageType == src_param.StorageType:
                dest_param.Set(src_value)
                elements_processed += 1
            elif dest_param.StorageType == StorageType.String and convert_values:
                dest_param.Set(str(src_value))
                elements_processed += 1
            else:
                error_log.append("Element ID {} ('{}'): Type mismatch between source parameter '{}' and destination parameter '{}'.".format(
                    elem_id, elem_name, src_param_name, dest_param_name))
        except Exception as e:
            error_log.append("Element ID {} ('{}'): Error - {}".format(elem_id, elem_name, str(e)))

    t.Commit()

    # Provide feedback to the user
    if elements_processed > 0:
        message = "Parameter values copied successfully for {} elements.".format(elements_processed)
        if error_log:
            message += "\nHowever, the following issues were encountered:\n\n" + "\n".join(error_log)
            TaskDialog.Show("Copy Parameter Values - Partial Success", message)
        else:
            TaskDialog.Show("Copy Parameter Values", message)
    else:
        message = "No elements were processed. The following issues were encountered:\n\n" + "\n".join(error_log)
        TaskDialog.Show("Copy Parameter Values - Errors", message)

if __name__ == "__main__":
    main()
