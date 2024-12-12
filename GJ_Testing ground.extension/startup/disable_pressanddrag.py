# Save this script in an extension’s "startup" folder, for example:
# [YourExtensionName].extension\startup\disable_pressanddrag.py
# Ensure the extension is properly recognized by pyRevit (avoid spaces in folder names).
# After loading Revit, open the pyRevit output panel to confirm the startup message prints.

import clr
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
clr.AddReference('RevitUI')

from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.ApplicationServices import Application

uiapp = __revit__
app = uiapp.Application

app.PressAndDragEnabled = False
print("AllowPressAndDrag has been disabled at startup.")

# Display a dialog at startup to confirm script loading
TaskDialog.Show("Startup", "PressAndDrag has been disabled.")

def on_idling(sender, args):
    if app.PressAndDragEnabled:
        TaskDialog.Show("Warning", "Using PressAndDrag is not recommended.")
        app.PressAndDragEnabled = False

uiapp.Idling += on_idling
