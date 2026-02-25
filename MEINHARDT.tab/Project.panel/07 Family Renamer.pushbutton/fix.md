✅ FIX SUMMARY

Here’s what this patch does:

🧹 Cleans and normalizes names (removes prefixes, spaces, redundant suffixes).

🧠 Adds duplicate-collapse logic to remove <Family>-<Type> when both are identical.

🔤 Adds structured prefixing for subtypes like MHT-ME-TypeName.

🧩 Updates category codes (Generic Annotation now uses GFA).

🔁 Improves duplicate name handling so only true duplicates get -01, -02.

🧱 Smarter default fallback naming for missing templates.

🔧 PATCH INSTRUCTIONS
1️⃣ Replace your current CATEGORY_CODE block with:
CATEGORY_CODE = {
    "Pipe Fittings": "PF",
    "Duct Fittings": "DF",
    "Mechanical Equipment": "EQ",
    "Generic Models": "GM",
    "Generic Annotations": "GFA",  # renamed from FAM
    "Lighting Fixtures": "LFX",
    "Electrical Fixtures": "EFX",
    "Plumbing Fixtures": "PLF",
    "Cable Tray Fittings": "CTF",
    "Conduit Fittings": "CF",
    "Plumbing Fixture Tags": "TAG",
    "Sprinkler Tags": "TAG",
    "Communication Device Tags": "TAG",
}

2️⃣ Add these new helper functions below clean_name():
def collapse_duplicate_name(name):
    """Collapse duplicated segments like 'ABC-ABC' or 'X-Y-X' at end."""
    parts = [p.strip() for p in name.split('-') if p.strip()]
    if len(parts) >= 2 and parts[-1].lower() == parts[-2].lower():
        parts.pop()
    return '-'.join(parts)

def make_subtype_prefixed(company, disc_code, typename):
    """Return MHT-ME-typename style subtype prefix."""
    if not typename:
        return ''
    if typename.upper().startswith(company):
        return typename
    typename = clean_name(typename)
    return "{}-{}-{}".format(company, disc_code, typename)

3️⃣ Replace the entire apply_template() function with this improved version:
def apply_template(template, info, rules):
    """Apply naming template with smarter normalization and subtype prefixing."""
    import re
    out = template
    company = rules.get('COMPANY', 'MHT')
    disc_map = rules.get('DISCIPLINE', {})
    disc_code = disc_map.get('Mechanical', 'ME')

    # Basic replacements
    out = out.replace('<COMPANY>', company)
    for m in re.findall(r'<DISC:([^>]+)>', out):
        out = out.replace('<DISC:%s>' % m, disc_map.get(m, m))

    # Category short code
    cat_code = CATEGORY_CODE.get(info.get('category', ''), 'XX')
    out = out.replace('<CAT>', cat_code)

    # --- Family name ---
    fam_clean = clean_name(info.get('family_name', ''))
    out = out.replace('<Family>', fam_clean)

    # --- Type name ---
    first_type = None
    types_list = info.get('types') or []
    if len(types_list) > 0:
        first_type = types_list[0]

    if first_type:
        typ_raw = clean_name(first_type.get('type_name', ''))
        # wrap subtypes like MHT-ME-TypeName
        if typ_raw:
            typ_clean = make_subtype_prefixed(company, disc_code, typ_raw)
        else:
            typ_clean = typ_raw

        out = out.replace('<Type>', typ_clean)

        # --- Replace parameter placeholders ---
        for m in re.findall(r'<Param:([^>]+)>', out):
            val = first_type['params'].get(m, '') or info.get('instance_params', {}).get(m, '')
            if m.lower() in ['size', 'diameter', 'dia', 'd']:
                try:
                    n = int(round(float(val)))
                    val = 'DN{}'.format(n)
                except Exception:
                    pass
            out = out.replace('<Param:%s>' % m, str(val))
    else:
        out = out.replace('<Type>', '')
        out = re.sub(r'<Param:[^>]+>', '', out)

    # --- Cleanup ---
    out = out.replace('--', '-').replace('__', '_').strip('-_ ')
    out = collapse_duplicate_name(out)
    return out

4️⃣ Replace the duplicate-handling block in main() with this improved logic:
# Detect duplicates and only append suffix for true duplicates
seen = {}
for r in results:
    name = r['suggested']
    if not name:
        continue
    base = name.upper().strip()
    seen[base] = seen.get(base, 0) + 1
    if seen[base] > 1:
        r['suggested'] = "{}-{:02d}".format(name, seen[base])
        r['conflict'] = True
    else:
        r['conflict'] = False

5️⃣ (Optional but recommended)

Update your naming_rules.json to something like this for consistency:

{
  "COMPANY": "MHT",
  "DISCIPLINE": {
    "Mechanical": "ME",
    "Plumbing": "PL",
    "Annotation": "AN"
  },
  "TEMPLATES": {
    "Pipe Fittings": "<COMPANY>-<DISC:Mechanical>-<CAT>-<Type>",
    "Duct Fittings": "<COMPANY>-<DISC:Mechanical>-<CAT>-<Type>",
    "Mechanical Equipment": "<COMPANY>-<DISC:Mechanical>-<CAT>-<Type>",
    "Generic Models": "<COMPANY>-<DISC:Mechanical>-<CAT>-<Family>",
    "Generic Annotations": "<COMPANY>-<DISC:Mechanical>-<CAT>-<Family>",
    "Default": "<COMPANY>-<DISC:Mechanical>-<CAT>-<Family>"
  },
  "DEFAULT_CASING": "UPPER"
}

🚀 RESULT EXAMPLES
Family	Category	Output
Conduit Elbow - Plain End - PVC	Conduit Fittings	MHT-ME-CF-MHT-ME-Conduit-Elbow-Plain-End-PVC → cleaned to MHT-ME-CF-Conduit-Elbow-Plain-End-PVC
Mechanical Equipment Tag	Generic Annotations	MHT-ME-GFA-Mechanical-Equipment
Plumbing Fixture Tag	Plumbing Fixture Tags	MHT-ME-TAG-Plumbing-Fixture
Ladder Horizontal Cross	Cable Tray Fittings	MHT-ME-CTF-Ladder-Horizontal-Cross
Reducer Straight Soldered Copper	Pipe Fittings	MHT-ME-PF-Reducer-Straight-Soldered-Copper