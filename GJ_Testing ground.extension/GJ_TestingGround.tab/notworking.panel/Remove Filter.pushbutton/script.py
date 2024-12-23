__doc__ = """
Title: FilterRemoval
Author: Goran Jovic
Version: 1.0
Description: Uklanjanje filtera za panele
"""
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document

# Start a new transaction
t = Transaction(doc, 'Delete Filters')
t.Start()

# List of filter names to be deleted
filter_names = ["CHECK HEIGHT", "CHECK WIDTH", "CHECK ALL"]

# Loop to find and delete each filter
for filter_name in filter_names:
    existing_filter = [e.Id for e in FilteredElementCollector(doc).OfClass(ParameterFilterElement) if e.Name == filter_name]
    if existing_filter:
        doc.Delete(existing_filter[0])

# Commit the transaction
t.Commit()
