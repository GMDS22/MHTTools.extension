# -*- coding: utf-8 -*-
__title__ = "Copy From\nLink"

__doc__ = """Version = 1.0
Date: 2026-07-02
Author: GM
Description:
Copy selected linked elements or selected-family instances from a linked model into the current project.
The copied elements keep the same world-space location as the link source.
How-to:
1. Select the source loaded Revit link.
2. Choose copy mode: pick linked elements or choose family names.
3. Confirm the copy and review the result summary.
"""

from pyrevit import revit, DB, forms, UI
from System.Collections.Generic import List
from Autodesk.Revit.Exceptions import OperationCanceledException

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

if not doc:
    forms.alert("No active document.", exitscript=True)


class _LinkedElementSelectionFilter(UI.Selection.ISelectionFilter):
    def __init__(self, allowed_link_id):
        self._allowed_link_id = allowed_link_id

    def AllowElement(self, element):
        try:
            return isinstance(element, DB.RevitLinkInstance) and element.Id == self._allowed_link_id
        except Exception:
            return False


class _UseDestinationTypesHandler(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes

    def AllowReference(self, reference, point):
        try:
            return (
                reference is not None
                and reference.ElementId == self._allowed_link_id
                and reference.LinkedElementId != DB.ElementId.InvalidElementId
            )
        except Exception:
            return False


def _loaded_links(host_doc):
    links = []
    for link_inst in DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance):
        try:
            link_doc = link_inst.GetLinkDocument()
            if link_doc is not None:
                links.append((link_inst, link_doc))
        except Exception:
            continue
    return links


def _pick_link(links):
    link_labels = []
    link_map = {}
    for link_inst, link_doc in links:
        try:
            link_name = link_inst.Name or "<Unnamed Link>"
        except Exception:
            link_name = "<Unnamed Link>"
        label = "{}  |  {}".format(link_name, link_doc.Title)
        if label in link_map:
            label = "{} ({})".format(label, link_inst.Id.IntegerValue)
        link_labels.append(label)
        link_map[label] = (link_inst, link_doc)

    selected = forms.SelectFromList.show(
        sorted(link_labels),
        title="Copy From Link | Select Source Link",
        button_name="Select Link",
        multiselect=False,
    )
    if not selected:
        return None, None
    return link_map.get(selected, (None, None))


def _pick_mode():
    return forms.CommandSwitchWindow.show(
        [
            "Pick linked elements in canvas",
            "Copy by family name (all instances)",
        ],
        message="Choose what to copy from the selected link",
    )


def _collect_ids_from_canvas(link_inst):
    pick_filter = _LinkedElementSelectionFilter(link_inst.Id)
    try:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            pick_filter,
            "Pick linked elements to copy, then click Finish",
        )
    except OperationCanceledException:
        return []
    except Exception as ex:
        forms.alert("Linked selection failed: {}".format(str(ex)), exitscript=True)
        return []

    ids = []
    seen = set()
    for reference in refs:
        try:
            linked_id = reference.LinkedElementId
            key = linked_id.IntegerValue
            if key not in seen:
                ids.append(linked_id)
                seen.add(key)
        except Exception:
            continue
    return ids


def _collect_ids_from_family_selection(link_doc):
    family_instances = list(DB.FilteredElementCollector(link_doc).OfClass(DB.FamilyInstance))
    if not family_instances:
        forms.alert("No family instances found in the selected linked model.", exitscript=True)

    family_to_ids = {}
    for inst in family_instances:
        try:
            symbol = getattr(inst, "Symbol", None)
            family = getattr(symbol, "Family", None)
            family_name = getattr(family, "Name", None) or "<Unnamed Family>"
            family_to_ids.setdefault(family_name, []).append(inst.Id)
        except Exception:
            continue

    if not family_to_ids:
        forms.alert("No copyable family instances found in the selected linked model.", exitscript=True)

    selected_families = forms.SelectFromList.show(
        sorted(family_to_ids.keys()),
        title="Copy From Link | Select Families",
        button_name="Copy Families",
        multiselect=True,
    )

    if not selected_families:
        return []

    if isinstance(selected_families, str):
        selected_families = [selected_families]

    ids = []
    seen = set()
    for family_name in selected_families:
        for elem_id in family_to_ids.get(family_name, []):
            key = elem_id.IntegerValue
            if key not in seen:
                ids.append(elem_id)
                seen.add(key)
    return ids


def _copy_from_link(link_inst, link_doc, source_ids):
    if not source_ids:
        forms.alert("No elements selected for copy.", exitscript=True)

    id_list = List[DB.ElementId]()
    for elem_id in source_ids:
        id_list.Add(elem_id)

    # Use the link total transform so copied elements land at matching world coordinates.
    transform = link_inst.GetTotalTransform()
    options = DB.CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(_UseDestinationTypesHandler())

    copied_ids = []
    with revit.Transaction("Copy From Link"):
        try:
            copied_ids = DB.ElementTransformUtils.CopyElements(
                link_doc,
                id_list,
                doc,
                transform,
                options,
            )
        except Exception as ex:
            forms.alert("Copy failed: {}".format(str(ex)), exitscript=True)

    return copied_ids


def main():
    links = _loaded_links(doc)
    if not links:
        forms.alert("No loaded Revit links found in this project.", exitscript=True)

    link_inst, link_doc = _pick_link(links)
    if link_inst is None or link_doc is None:
        forms.alert("No source link selected.", exitscript=True)

    mode = _pick_mode()
    if not mode:
        forms.alert("Copy canceled.", exitscript=True)

    if mode == "Pick linked elements in canvas":
        source_ids = _collect_ids_from_canvas(link_inst)
    else:
        source_ids = _collect_ids_from_family_selection(link_doc)

    copied_ids = _copy_from_link(link_inst, link_doc, source_ids)
    copied_count = len(list(copied_ids)) if copied_ids is not None else 0

    forms.alert(
        "Copied {} element(s) from link '{}' into the current model.".format(copied_count, link_doc.Title),
        title="Copy From Link",
        warn_icon=False,
    )


if __name__ == "__main__":
    main()
