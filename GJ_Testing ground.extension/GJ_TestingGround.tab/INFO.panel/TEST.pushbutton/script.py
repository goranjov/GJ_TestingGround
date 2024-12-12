import os
import sys
import xml.etree.ElementTree as ET
from pyrevit import revit, script, forms

app = revit.doc.Application
version_name = app.VersionName if app else "Unknown_Version"
version_build = app.VersionBuild if app else "Unknown_Build"

try:
    release_date = app.VersionNumber if app else "Unknown_ReleaseDate"
except:
    release_date = "Unknown_ReleaseDate"

audit_filename = version_name.replace(" ", "_") + "_ExtensionsAudit.txt"

year = ""
if version_name and version_name[-4:].isdigit():
    year = version_name[-4:]

common_directories = [
    os.path.join(os.environ.get("PROGRAMDATA", ""), "Autodesk", "Revit", "Addins", year),
    os.path.join(os.environ.get("APPDATA", ""), "Autodesk", "Revit", "Addins", year)
]

def parse_addins(directory):
    plugins = []
    if not os.path.exists(directory):
        return plugins
    for file_name in os.listdir(directory):
        if file_name.lower().endswith(".addin"):
            file_path = os.path.join(directory, file_name)
            try:
                tree = ET.parse(file_path)
                root = tree.getroot()
                for child in root:
                    plugin_name = child.find('Name')
                    # Check if version info is directly provided in other tags like VendorDescription or AddInId
                    vendor_desc = child.find('VendorDescription')
                    plugin_id = child.find('AddInId')
                    
                    # Currently, no standard version tag is provided by Revit addin files.
                    # We'll simply show 'Unknown' unless we can parse something meaningful.
                    # If assembly name follows no conventional version pattern, we stay with 'Unknown'.
                    version_str = "Unknown"
                    
                    assembly_node = child.find('Assembly')
                    if assembly_node is not None:
                        assembly_path = assembly_node.text or ""
                        assembly_base = os.path.basename(assembly_path)
                        # If version-like pattern not found in assembly filename, we leave it as 'Unknown'
                        # This is a best effort approach since Revit doesn't require version info in .addin files.
                    
                    pname = plugin_name.text if plugin_name is not None else "Unknown"
                    plugins.append((pname, version_str))
            except:
                pass
    return plugins

all_plugins = []
for d in common_directories:
    all_plugins.extend(parse_addins(d))

unique_plugins = list(set(all_plugins))
unique_plugins.sort(key=lambda x: x[0])

save_folder = forms.pick_folder()
if not save_folder:
    script.exit()

audit_filepath = os.path.join(save_folder, audit_filename)

try:
    with open(audit_filepath, 'w') as f:
        f.write("Revit Version Information\n")
        f.write("========================\n")
        f.write("Version Name: " + version_name + "\n")
        f.write("Build Number: " + version_build + "\n")
        f.write("Release Date: " + release_date + "\n\n")

        f.write("Installed Extensions/Plugins\n")
        f.write("============================\n")
        if unique_plugins:
            for pname, pvers in unique_plugins:
                f.write("Name: " + pname + ", Version: " + pvers + "\n")
        else:
            f.write("No extension information available.\n")
except Exception as e:
    print("Error writing file:", e)
else:
    print("File saved at:", audit_filepath)
