# -*- coding: utf-8 -*-
__title__ = "Hide Grids (Safe)"

from pyrevit import revit, DB, forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

view = revit.active_view

if not view:
    forms.alert("No active view", exitscript=True)

collector = DB.FilteredElementCollector(doc, view.Id)\
    .OfClass(DB.RevitLinkInstance)

link_types = {}
for inst in collector:
    link_types[inst.GetTypeId().IntegerValue] = doc.GetElement(inst.GetTypeId())

link_types = link_types.values()

if not link_types:
    forms.alert("No visible links", exitscript=True)


def hide_grids(view, link_type):
    try:
        settings = view.GetLinkOverrides(link_type.Id)
        if not settings:
            return

        grid_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Grids)

        # Only modify grid visibility
        try:
            settings.SetCategoryHidden(grid_cat.Id, True)
        except:
            return

        view.SetLinkOverrides(link_type.Id, settings)

    except:
        pass


with revit.Transaction("Hide Linked Grids"):
    for lt in link_types:
        hide_grids(view, lt)