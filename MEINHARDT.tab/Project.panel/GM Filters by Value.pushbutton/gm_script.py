# -*- coding: utf-8 -*-
# GMToolbox version of Filterbyvalue (renamed)
# ...original script.py logic from pyChilizer...

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List
import csv
import re
import sys
from rpw.ui.forms import FlexForm, Label, ComboBox, Separator, Button, TextBox, CheckBox

doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView

# --- Linked File Helpers ---
def get_linked_docs():
    links = [el for el in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)]
    link_docs = []
    for link in links:
        try:
            ldoc = link.GetLinkDocument()
            if ldoc:
                link_docs.append((link, ldoc))
        except:
            pass
    return link_docs

def get_linked_doc_by_name(name):
    for link, ldoc in get_linked_docs():
        if ldoc.Title == name:
            return ldoc
    return None

# --- UI: Linked File Selection ---
link_docs = get_linked_docs()
link_names = [ldoc.Title for link, ldoc in link_docs]
link_choice = None
if link_names:
    link_choice = forms.SelectFromList.show(link_names, title="Select Linked File (optional)")
    if link_choice:
        selected_link_doc = get_linked_doc_by_name(link_choice)
    else:
        selected_link_doc = None
else:
    selected_link_doc = None

# --- Collect MEP system types from linked file if selected ---
def collect_system_types_from_doc(target_doc, bic):
    return (
        DB.FilteredElementCollector(target_doc)
        .OfCategory(bic)
        .WhereElementIsElementType()
        .ToElements()
    )

if selected_link_doc:
    duct_systems_link = collect_system_types_from_doc(selected_link_doc, DB.BuiltInCategory.OST_DuctSystem)
    pipe_systems_link = collect_system_types_from_doc(selected_link_doc, DB.BuiltInCategory.OST_PipingSystem)
else:
    duct_systems_link = []
    pipe_systems_link = []

# --- UI: Per-category selection for linked file ---
linked_cat_options = []
if selected_link_doc:
    # List all MEP categories present in the linked file
    mep_bics = [
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_DuctFitting,
        DB.BuiltInCategory.OST_DuctAccessory,
        DB.BuiltInCategory.OST_DuctInsulations,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory,
        DB.BuiltInCategory.OST_PipeInsulations,
        # Add more MEP categories as needed
    ]
    for bic in mep_bics:
        elems = DB.FilteredElementCollector(selected_link_doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
        if elems:
            linked_cat_options.append((bic, bic.ToString().replace('Autodesk.Revit.DB.BuiltInCategory, ', '')))

if linked_cat_options:
    cat_names = [name for bic, name in linked_cat_options]
    selected_linked_cats = forms.SelectFromList.show(cat_names, multiselect=True, title="Select Categories from Linked File")
    if selected_linked_cats:
        selected_linked_bics = [bic for bic, name in linked_cat_options if name in selected_linked_cats]
    else:
        selected_linked_bics = []
else:
    selected_linked_bics = []

# ...existing code...
# Now you can use selected_linked_bics to filter or process only the selected categories from the linked file.
