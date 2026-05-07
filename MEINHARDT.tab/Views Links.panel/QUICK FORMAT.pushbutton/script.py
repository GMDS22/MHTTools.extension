# -*- coding: utf-8 -*-
__title__ = "QUICK FORMAT"

from pyrevit import revit, DB, forms
from System.Collections.Generic import List

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

view = revit.active_view

if not view:
    forms.alert("No active view", exitscript=True)


def _get_view_template(view):
    """Get view template if assigned"""
    try:
        template_id = view.ViewTemplateId
        if template_id and template_id != DB.ElementId.InvalidElementId:
            return doc.GetElement(template_id)
    except Exception:
        pass
    return None


def _template_controls_revit_links(view):
    """Check if view template controls RVT Links visibility"""
    template = _get_view_template(view)
    if template is None:
        return False
    try:
        non_controlled = template.GetNonControlledTemplateParameterIds()
        if non_controlled is None:
            return False
        links_param = DB.ElementId(int(DB.BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))
        return links_param not in non_controlled
    except Exception:
        return False


def _release_revit_links_template_control(view):
    """Release RVT Links from template control"""
    template = _get_view_template(view)
    if template is None:
        return False, None

    links_param = DB.ElementId(int(DB.BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))

    try:
        non_controlled = list(template.GetNonControlledTemplateParameterIds() or [])
        if any(item == links_param for item in non_controlled):
            return False, None

        updated = List[DB.ElementId]()
        for item in non_controlled:
            updated.Add(item)
        updated.Add(links_param)
        template.SetNonControlledTemplateParameterIds(updated)
        return True, getattr(template, 'Name', '<Unnamed Template>') or "<Unnamed Template>"
    except Exception as exc:
        return False, str(exc)


def _find_best_room_view_in_link(link_doc, host_level):
    """Find the best plan view in linked document that shows rooms at matching level"""
    try:
        if not link_doc or not host_level:
            return None
        
        # Get all plan views from linked document
        plan_views = []
        try:
            collector = DB.FilteredElementCollector(link_doc)\
                        .OfClass(DB.View)\
                        .WhereElementIsNotElementType()
            
            for v in collector:
                try:
                    # Check if it's a plan view (ViewFamily.Plan)
                    if hasattr(v, 'ViewType') and v.ViewType == DB.ViewType.FloorPlan:
                        # Check if view level matches host level (by elevation)
                        view_level = get_view_level(v)
                        if view_level and abs(view_level.Elevation - host_level.Elevation) < 0.1:  # Within 1mm
                            plan_views.append(v)
                except Exception:
                    continue
        except Exception:
            pass
        
        if not plan_views:
            return None
        
        # Find views that have rooms visible
        room_views = []
        for v in plan_views:
            try:
                # Check if Rooms category is visible in this view
                rooms_cat = link_doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Rooms)
                if rooms_cat:
                    vis_settings = v.GetCategoryHidden(rooms_cat.Id)
                    if not vis_settings:  # Not hidden = visible
                        room_views.append(v)
            except Exception:
                continue
        
        if not room_views:
            return None
        
        # Select the best view - prefer views with names containing common plan view keywords
        best_view = None
        priority_keywords = ['plan', 'floor', 'level', 'archi', 'room']
        
        for v in room_views:
            try:
                name = (v.Name or "").lower()
                if any(kw in name for kw in priority_keywords):
                    best_view = v
                    break
            except Exception:
                continue
        
        # If no priority match, take the first room view
        if not best_view and room_views:
            best_view = room_views[0]
        
        return best_view
    
    except Exception:
        return None


def _safe_link_type_name(link_type):
    """Safely get link type name with fallback"""
    try:
        return link_type.Name or "<Unnamed Link>"
    except Exception:
        return "<Unnamed Link>"


def _get_view_level(view):
    """Get the level of a view"""
    try:
        return view.GenLevel
    except Exception:
        return None


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


def _configure_link_for_room_names(view, link_type):
    """Configure link to show room names: find best view, set Custom mode with By Host filters"""
    try:
        settings = view.GetLinkOverrides(link_type.Id)
        if not settings:
            return False, "No link settings"
        
        # Get the linked document
        link_doc = _find_link_doc_for_type(doc, link_type.Id)
        if not link_doc:
            return False, "Cannot access linked document"
        
        # Get current view's level
        host_level = _get_view_level(view)
        if not host_level:
            return False, "Cannot determine view level"
        
        # Find the best room view in the linked document
        best_room_view = _find_best_room_view_in_link(link_doc, host_level)
        if not best_room_view:
            return False, "No suitable room view found in linked document"
        
        # Set the specific linked view to show
        try:
            settings.LinkedViewId = best_room_view.Id
        except Exception as exc:
            return False, "Failed to set LinkedViewId: {}".format(exc)
        
        # Set visibility to Custom mode
        custom_mode = _get_link_visibility_type('Custom')
        if custom_mode is not None:
            try:
                settings.LinkVisibilityType = custom_mode
            except Exception:
                pass
        
        # Set Underlay to "By Host View" for clean display
        by_host_view = _get_link_visibility_type('ByHostView')
        if by_host_view is not None:
            try:
                settings.Underlay = by_host_view
            except Exception:
                pass
        
        # Enable halftone for better visibility
        if hasattr(settings, 'Halftone'):
            try:
                settings.Halftone = True
            except Exception:
                pass
        
        # Hide grids
        grid_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Grids)
        if grid_cat is not None:
            try:
                settings.SetCategoryHidden(grid_cat.Id, True)
            except Exception:
                pass
        
        # Hide dimensions
        dim_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Dimensions)
        if dim_cat is not None:
            try:
                settings.SetCategoryHidden(dim_cat.Id, True)
            except Exception:
                pass
        
        try:
            view.SetLinkOverrides(link_type.Id, settings)
            return True, "Configured with view: {}".format(getattr(best_room_view, 'Name', '<Unnamed>') or "<Unnamed>")
        except Exception as exc:
            return False, "SetLinkOverrides failed: {}".format(exc)
    
    except Exception as exc:
        return False, str(exc)


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
    forms.alert("No visible links in the current view.", exitscript=True)

with revit.Transaction("QUICK FORMAT"):
    template_released = False
    template_note = None
    
    # Check and release template control on RVT Links if needed
    if _template_controls_revit_links(view):
        released, detail = _release_revit_links_template_control(view)
        if released:
            template_released = True
            template_note = "Released RVT Links from template: {}".format(detail or "<Unnamed>")
        
        # Verify template control was released
        if _template_controls_revit_links(view):
            forms.alert(
                "RVT Links are still controlled by the view template.\n"
                "Cannot proceed. Please manually release the template control.",
                exitscript=True,
            )
    
    # Configure each link for room names display
    configured = 0
    config_details = []
    for lt in link_types:
        success, msg = _configure_link_for_room_names(view, lt)
        if success:
            configured += 1
            config_details.append("{}: {}".format(_safe_link_type_name(lt), msg))
        else:
            config_details.append("{}: FAILED - {}".format(_safe_link_type_name(lt), msg))

# Report results
summary = [
    "View: {}".format(getattr(view, 'Name', '<Unnamed>') or "<Unnamed>"),
    "Visible links: {}".format(len(link_types)),
    "Links configured: {}/{}".format(configured, len(link_types)),
]
if template_released:
    summary.append("Template control released: Yes")
if template_note:
    summary.append(template_note)

if config_details:
    summary.append("")
    summary.append("Configuration Details:")
    summary.extend(config_details)

forms.alert("\n".join(summary), title="QUICK FORMAT", warn_icon=False)
