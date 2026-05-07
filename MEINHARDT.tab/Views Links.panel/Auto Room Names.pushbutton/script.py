# -*- coding: utf-8 -*-
__title__ = "Tag Rooms in Links"
__doc__ = """Version = 2.0
Multi-sheet room tagging from architectural links.
1. Select sheets to process.
2. For each sheet, finds placed views.
3. For each view, detects visible architectural links.
4. Choose room tag family.
5. Tags all rooms in linked models.
"""

from pyrevit import revit, DB, forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

if not doc:
    forms.alert("No active document", exitscript=True)


def _safe_link_type_name(link_type):
    try:
        return link_type.Name or "<Unnamed Link>"
    except Exception:
        return "<Unnamed Link>"


def _is_architectural_link(link_name):
    """Detect if link name suggests architectural model"""
    arch_keywords = ["arch", "architectural", "ar_", "a_", "-a-", "[a]"]
    name_lower = (link_name or "").lower()
    return any(kw in name_lower for kw in arch_keywords)


def get_all_sheets():
    """Collect all sheets in document"""
    sheets = []
    try:
        all_sheets = DB.FilteredElementCollector(doc)\
                     .OfClass(DB.ViewSheet)\
                     .ToElements()
        for sheet in all_sheets:
            try:
                sheet_num = getattr(sheet, 'SheetNumber', '') or ''
                sheet_name = getattr(sheet, 'Name', '') or ''
                display = "{} | {}".format(sheet_num, sheet_name)
                sheets.append((sheet, display, sheet_num))
            except Exception:
                continue
    except Exception:
        pass
    
    # Sort by sheet number
    sheets.sort(key=lambda x: x[2].lower())
    return sheets


def get_placed_views_on_sheet(sheet):
    """Get all views placed on a given sheet"""
    views = []
    try:
        viewport_ids = sheet.GetAllViewports()
        for vp_id in viewport_ids:
            vp = doc.GetElement(vp_id)
            if vp and hasattr(vp, 'ViewId'):
                view = doc.GetElement(vp.ViewId)
                if view:
                    views.append(view)
    except Exception:
        pass
    return views


def get_visible_arch_links_in_view(view):
    """Find all visible architectural links in a view"""
    arch_links = []
    try:
        instances = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.RevitLinkInstance).ToElements()
        for inst in instances:
            try:
                if inst.IsHidden(view):
                    continue
                
                lt = doc.GetElement(inst.GetTypeId())
                if lt is None:
                    continue
                if hasattr(lt, 'IsLoaded') and not lt.IsLoaded:
                    continue
                
                link_name = _safe_link_type_name(lt)
                if _is_architectural_link(link_name):
                    arch_links.append(inst)
            except Exception:
                continue
    except Exception:
        pass
    
    return arch_links


def get_view_level(view):
    try:
        return view.GenLevel
    except Exception:
        return None


def get_room_tag_types():
    """Collect all room tag family symbols"""
    types = []
    try:
        tag_types = DB.FilteredElementCollector(doc)\
                    .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
                    .WhereElementIsElementType()\
                    .ToElements()
        for t in tag_types:
            try:
                name = getattr(t, 'Name', '<Unnamed>') or '<Unnamed>'
                family = getattr(t.Family, 'Name', '<Unnamed Family>') if hasattr(t, 'Family') else '<Unnamed Family>'
                display = "{} - {}".format(family, name)
                types.append((t, display))
            except Exception:
                continue
    except Exception:
        pass
    
    types.sort(key=lambda x: x[1].lower())
    return types


def get_existing_room_tags_in_view(view):
    """Track already-tagged room IDs"""
    existing = set()
    try:
        tags = DB.FilteredElementCollector(doc, view.Id)\
               .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
               .WhereElementIsNotElementType()\
               .ToElements()
        
        for tag in tags:
            try:
                if hasattr(tag, 'TaggedLocalElementId'):
                    linked_id = tag.TaggedLocalElementId
                    if linked_id and linked_id != DB.ElementId.InvalidElementId:
                        existing.add(linked_id.IntegerValue)
                        continue
            except Exception:
                pass
            
            try:
                if hasattr(tag, 'Room') and tag.Room is not None:
                    existing.add(tag.Room.Id.IntegerValue)
            except Exception:
                pass
    except Exception:
        pass
    
    return existing


def create_room_tag(room, link_instance, tag_type, view):
    """Create a room tag at room location, transformed to view coords"""
    try:
        if room.Location is None:
            return None
        
        point = room.Location.Point
        if point is None:
            return None
        
        # Transform room point to host coordinates
        world_point = point
        try:
            transform = link_instance.GetTransform()
            if transform:
                world_point = transform.OfPoint(point)
        except Exception:
            pass
        
        uv = DB.UV(world_point.X, world_point.Y)
        tag = None
        
        # Try NewRoomTag
        try:
            if hasattr(doc.Create, 'NewRoomTag'):
                linked_room_id = DB.LinkElementId(room.Id)
                tag = doc.Create.NewRoomTag(linked_room_id, uv, view.Id)
        except Exception:
            tag = None
        
        # Fallback to IndependentTag.Create
        if tag is None:
            try:
                if hasattr(DB.IndependentTag, 'Create'):
                    tag = DB.IndependentTag.Create(
                        doc,
                        view.Id,
                        tag_type.Id if tag_type else DB.ElementId.InvalidElementId,
                        link_instance.Id,
                        DB.TagMode.TM_ADDBY_CATEGORY,
                        world_point
                    )
            except Exception:
                tag = None
        
        # Apply tag type
        if tag and tag_type:
            try:
                tag.ChangeTypeId(tag_type.Id)
            except Exception:
                try:
                    if hasattr(tag, 'RoomTagType'):
                        tag.RoomTagType = tag_type
                except Exception:
                    pass
        
        return tag
    except Exception:
        return None


def tag_rooms_in_link(view, link_inst, existing_ids, tag_type):
    """Tag rooms from linked model in given view"""
    count = 0
    try:
        link_doc = link_inst.GetLinkDocument()
        if not link_doc:
            return 0
        
        view_level = get_view_level(view)
        
        # Get rooms from link
        rooms = list(DB.FilteredElementCollector(link_doc)\
                     .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                     .WhereElementIsNotElementType()\
                     .ToElements())
        
        # Filter by level for performance
        if view_level:
            rooms = [r for r in rooms if hasattr(r, 'LevelId') and r.LevelId == view_level.Id]
        
        for room in rooms:
            try:
                if room.Id.IntegerValue in existing_ids:
                    continue
                
                tag = create_room_tag(room, link_inst, tag_type, view)
                if tag:
                    count += 1
                    existing_ids.add(room.Id.IntegerValue)
            except Exception:
                continue
    except Exception:
        pass
    
    return count


# ====== MAIN UI FLOW ======

# 1) Collect sheets
all_sheets = get_all_sheets()
if not all_sheets:
    forms.alert("No sheets found in document", exitscript=True)

# 2) Select sheets
sheet_labels = [display for _, display, _ in all_sheets]
selected_labels = forms.SelectFromList.show(
    sheet_labels,
    title="Select sheets to tag rooms in",
    multiselect=True,
    button_name="Select"
)

if not selected_labels:
    forms.alert("No sheets selected", exitscript=True)

if isinstance(selected_labels, basestring if 'basestring' in globals() else str):
    selected_labels = [selected_labels]

selected_sheets = [all_sheets[i][0] for i, label in enumerate(sheet_labels) if label in selected_labels]

# 3) Get room tag families
tag_types = get_room_tag_types()
if not tag_types:
    forms.alert("No Room Tag families found in project", exitscript=True)

tag_labels = [display for _, display in tag_types]
chosen_tag_label = forms.SelectFromList.show(
    tag_labels,
    title="Choose Room Tag family",
    multiselect=False,
    button_name="Select"
)

if not chosen_tag_label:
    forms.alert("No tag family selected", exitscript=True)

if isinstance(chosen_tag_label, basestring if 'basestring' in globals() else str):
    chosen_tag_label = [chosen_tag_label]

chosen_tag_type = tag_types[tag_labels.index(chosen_tag_label[0])][0]

# 4) Process sheets and placed views
total_tagged = 0
view_count = 0
skipped_views = 0

with revit.Transaction("Tag Rooms in Links"):
    for sheet in selected_sheets:
        placed_views = get_placed_views_on_sheet(sheet)
        
        for view in placed_views:
            try:
                if view.IsTemplate:
                    continue
                
                arch_links = get_visible_arch_links_in_view(view)
                if not arch_links:
                    skipped_views += 1
                    continue
                
                view_count += 1
                existing_ids = get_existing_room_tags_in_view(view)
                
                for link_inst in arch_links:
                    total_tagged += tag_rooms_in_link(view, link_inst, existing_ids, chosen_tag_type)
            except Exception:
                continue

# 5) Report
summary = (
    "Sheets processed: {}\n"
    "Views with arch links: {}\n"
    "Views skipped (no arch link): {}\n"
    "Room tags created: {}"
).format(len(selected_sheets), view_count, skipped_views, total_tagged)

forms.alert(summary, title=__title__)
