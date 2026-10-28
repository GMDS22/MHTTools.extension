# -*- coding: utf-8 -*-
"""Dependency-free geometry and reconciliation helpers for the detector."""

import csv
import hashlib
import math
import re


MM_TO_FEET = 0.0032808398950131233
FEET_TO_MM = 304.8


def ft_to_mm(value):
    return float(value) * FEET_TO_MM


def mm_to_ft(value):
    return float(value) * MM_TO_FEET


def distance(a, b):
    return math.sqrt(sum((float(a[i]) - float(b[i])) ** 2 for i in range(3)))


def boxes_overlap(a, b, tolerance=0.0):
    if not a or not b:
        return False
    for i in range(3):
        if min(a[1][i], b[1][i]) < max(a[0][i], b[0][i]) - tolerance:
            return False
    return True


def segment_box_intersection(start, end, bounds_min, bounds_max, tolerance=0.0):
    """Return (distance range, intersection point) or None for an axis-aligned solid."""
    lo = 0.0
    hi = 1.0
    p0 = [float(start[i]) for i in range(3)]
    delta = [float(end[i]) - p0[i] for i in range(3)]
    for i in range(3):
        if abs(delta[i]) <= tolerance:
            if p0[i] < float(bounds_min[i]) - tolerance or p0[i] > float(bounds_max[i]) + tolerance:
                return None
            continue
        a = (float(bounds_min[i]) - tolerance - p0[i]) / delta[i]
        b = (float(bounds_max[i]) + tolerance - p0[i]) / delta[i]
        axis_lo, axis_hi = min(a, b), max(a, b)
        lo = max(lo, axis_lo)
        hi = min(hi, axis_hi)
        if lo > hi:
            return None
    if hi < 0.0 or lo > 1.0:
        return None
    lo = max(0.0, lo)
    hi = min(1.0, hi)
    point = [p0[i] + (lo * delta[i]) for i in range(3)]
    return lo, hi, point


def make_finding_id(host_document_guid, link_instance_uid, wall_uid, duct_uid, geometry_signature):
    parts = [host_document_guid or "UNKNOWN-HOST", link_instance_uid or "HOST-WALL", wall_uid or "", duct_uid or "", geometry_signature or ""]
    raw = "|".join(parts)
    return "MHT-FS-" + hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16].upper()


def reconcile(existing_rows, current_rows, coordinate_tolerance=0.5):
    """Reuse finding IDs for source identities and flag ambiguous moved findings."""
    current_by_identity = {}
    for row in current_rows:
        current_by_identity.setdefault(row["source_identity"], []).append(row)

    existing_by_identity = {}
    for row in existing_rows:
        existing_by_identity.setdefault(row.get("source_identity", ""), []).append(row)

    new_rows = []
    reused = []
    obsolete = []
    ambiguous = []
    for identity, rows in current_by_identity.items():
        matches = existing_by_identity.get(identity, [])
        if len(rows) == 1 and len(matches) == 1:
            row = dict(rows[0])
            row["finding_id"] = matches[0].get("finding_id") or row["finding_id"]
            row["review_status"] = matches[0].get("review_status", "Not Reviewed")
            row["comments"] = matches[0].get("comments", "")
            reused.append(row)
        elif len(rows) > 1 or len(matches) > 1:
            ambiguous.extend(rows)
        else:
            new_rows.extend(rows)

    matched_existing_ids = set(r.get("finding_id") for r in reused)
    for rows in existing_by_identity.values():
        for row in rows:
            if row.get("finding_id") not in matched_existing_ids:
                obsolete.append(row)
    return new_rows, reused, obsolete, ambiguous


def wall_classification(fire_text="", smoke_text="", fire_both_text=""):
    """Normalize common explicit project values without inferring defaults."""
    def tokens(value):
        return set(str(value or "").strip().lower().replace("_", " ").replace("-", " ").split())
    fire = tokens(fire_text)
    smoke = tokens(smoke_text)
    fire_both = tokens(fire_both_text)
    truth = {"yes", "true", "1", "fire", "rated", "smoke"}
    has_fire = bool(fire & truth)
    has_smoke = bool(smoke & truth) or bool(fire_both & {"smoke"})
    values = [str(fire_both_text or ""), str(fire_text or ""), str(smoke_text or "")]
    if not any(str(v).strip() for v in values):
        return "UNKNOWN"
    return "BOTH" if has_fire and has_smoke else "FIRE" if has_fire else "SMOKE" if has_smoke else "NOT_RATED"


def classify_wall_material(value):
    """Classify explicit material/rating text from NWB_Material without inference."""
    text = str(value or "").strip().lower()
    if not text:
        return "UNKNOWN"
    normalized = re.sub(r"[^a-z0-9]+", " ", text).strip()
    if normalized in ("0", "false", "no", "none", "unrated", "non rated", "nr"):
        return "NOT_RATED"
    has_fire = bool(re.search(r"(?:^|\s)(?:fr|fire rated|fireresistant|fire resistance)(?:\s|$)", normalized))
    has_smoke = bool(re.search(r"(?:^|\s)(?:smoke rated|smoke resistant|smoke)(?:\s|$)", normalized))
    if has_fire and has_smoke:
        return "BOTH"
    if has_fire:
        return "FIRE"
    if has_smoke:
        return "SMOKE"
    return "UNKNOWN"


def split_parameter_aliases(value):
    """Parse comma, semicolon, or newline-separated parameter aliases."""
    aliases = []
    for alias in re.split(r"[,;\r\n]+", value or ""):
        alias = alias.strip()
        if alias and alias.lower() not in [item.lower() for item in aliases]:
            aliases.append(alias)
    return aliases


def write_csv(path, rows):
    fields = [
        "finding_id", "status", "review_status", "comments", "host_document_guid",
        "link_instance_uid", "source_document", "wall_uid", "duct_uid", "wall_classification",
        "point_x", "point_y", "point_z", "wall_thickness_mm", "source_identity",
        "warnings", "detected_at"
    ]
    import io
    with io.open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow(dict((field, row.get(field, "")) for field in fields))
    return len(rows)
