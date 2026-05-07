# -*- coding: utf-8 -*-
__title__ = "Copy Room Tags"

__doc__ = """Version = 1.0
Date: 2026-03-30
Author: GM
Description:
Copy room tag families from linked models to make them available in the current project.
How-to:
1. Run the tool in a view with visible architectural links.
2. Select which room tag families to copy from the linked models.
3. Families will be copied to the current project.
"""

from pyrevit import revit, DB, forms
from System.Collections.Generic import List

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

view = revit.active_view

if not doc:
    forms.alert("No active document", exitscript=True)

if not view:
    forms.alert("No active view", exitscript=True)


def _safe_link_type_name(link_type):
    """Safely get link type name with fallback"""
    try:
        return link_type.Name or "<Unnamed Link>"
    except Exception:
        return "<Unnamed Link>"


def _find_link_doc_for_type(host_doc, link_type_id):
    """Find linked document for a given link type by searching all instances"""
    try:
        instances = DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance).ToElements()
        for instance in instances:
            try:
                if instance.GetTypeId() == link_type_id:
                    link_doc = instance.GetLinkDocument()
                    if link_doc:
                        return link_doc
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_room_tag_families_from_link(link_doc, link_name):
    """Get all room tag family symbols from a linked document"""
    families = []
    try:
        # Get room tag types from linked document
        tag_types = DB.FilteredElementCollector(link_doc)\
                    .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
                    .WhereElementIsElementType()\
                    .ToElements()

        for tag_type in tag_types:
            try:
                name = getattr(tag_type, 'Name', '<Unnamed>') or '<Unnamed>'
                family = getattr(tag_type.Family, 'Name', '<Unnamed Family>') if hasattr(tag_type, 'Family') else '<Unnamed Family>'
                display = "{} - {} [{}]".format(family, name, link_name)
                families.append((tag_type, display))
            except Exception:
                continue
    except Exception:
        pass

    return families


def get_existing_room_tag_names(host_doc):
    """Get names of existing room tag families in host document"""
    existing = set()
    try:
        tag_types = DB.FilteredElementCollector(host_doc)\
                    .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
                    .WhereElementIsElementType()\
                    .ToElements()

        for tag_type in tag_types:
            try:
                name = getattr(tag_type, 'Name', '') or ''
                if name:
                    existing.add(name.lower())
            except Exception:
                continue
    except Exception:
        pass

    return existing


# Get visible links in current view
collector = DB.FilteredElementCollector(doc, view.Id)\
    .OfClass(DB.RevitLinkInstance)

link_types = {}
for inst in collector:
    try:
        link_type = doc.GetElement(inst.GetTypeId())
        if link_type is not None:
            link_types[link_type.Id.IntegerValue] = link_type
    except Exception:
        continue

link_types = list(link_types.values())

if not link_types:
    forms.alert("No visible links found in the current view.", exitscript=True)

# Collect all room tag families from all links
all_tag_families = []
for lt in link_types:
    link_name = _safe_link_type_name(lt)
    link_doc = _find_link_doc_for_type(doc, lt.Id)
    if link_doc:
        families = get_room_tag_families_from_link(link_doc, link_name)
        all_tag_families.extend(families)

if not all_tag_families:
    forms.alert("No room tag families found in the linked models.", exitscript=True)

# Get existing family names to avoid duplicates
existing_names = get_existing_room_tag_names(doc)

# Filter out families that already exist
available_families = []
for tag_type, display in all_tag_families:
    try:
        name = getattr(tag_type, 'Name', '') or ''
        if name.lower() not in existing_names:
            available_families.append((tag_type, display))
    except Exception:
        continue

if not available_families:
    forms.alert("All room tag families from links are already available in the current project.", exitscript=True)

# Let user select which families to copy
family_labels = [display for _, display in available_families]
selected_labels = forms.SelectFromList.show(
    family_labels,
    title="Select Room Tag Families to Copy",
    multiselect=True,
    button_name="Copy Selected"
)

if not selected_labels:
    forms.alert("No families selected", exitscript=True)

if isinstance(selected_labels, (str, unicode if 'unicode' in globals() else str)):
    selected_labels = [selected_labels]

# Get the selected family elements
selected_families = []
for i, (_, display) in enumerate(available_families):
    if display in selected_labels:
        selected_families.append(available_families[i][0])

if not selected_families:
    forms.alert("No valid families found to copy", exitscript=True)

# Copy the families
copied_count = 0
with revit.Transaction("Copy Room Tag Families"):
    try:
        # Copy elements from link document to host document
        # We need to get the source document for each family
        for tag_type in selected_families:
            try:
                # Find which link this family belongs to
                source_doc = None
                for lt in link_types:
                    link_doc = _find_link_doc_for_type(doc, lt.Id)
                    if link_doc:
                        # Check if this family exists in this link
                        link_tag_types = DB.FilteredElementCollector(link_doc)\
                                        .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
                                        .WhereElementIsElementType()\
                                        .ToElements()
                        for link_tag_type in link_tag_types:
                            if link_tag_type.Id == tag_type.Id:
                                source_doc = link_doc
                                break
                    if source_doc:
                        break

                if source_doc:
                    # Copy the family
                    element_ids = List[DB.ElementId]()
                    element_ids.Add(tag_type.Id)
                    DB.ElementTransformUtils.CopyElements(source_doc, element_ids, doc, None, None)
                    copied_count += 1

            except Exception as e:
                # Continue with other families if one fails
                continue

    except Exception as e:
        forms.alert("Error copying families: {}".format(str(e)), exitscript=True)

# Report results
if copied_count > 0:
    forms.alert("Successfully copied {} room tag families to the current project.".format(copied_count), warn_icon=False)
else:
    forms.alert("No families were copied. They may already exist or there was an error.", exitscript=True)