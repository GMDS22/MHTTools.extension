# NWB Autofill Speed Optimization Implementation Plan

## Overview

This document describes the implementation plan for creating fast, optimized versions of:
- **NWB_PARAMETERS AutoFill (Fast).pushbutton**
- **NWB Dim AutoFill (Fast).pushbutton**

Both tools are copied from the original versions and will receive targeted performance optimizations without changing functional accuracy.

**Estimated speedup:** Autofill 60-80%, Dim 50-70%

---

## Optimization Sequence (In Priority Order)

### Phase 1: High-Impact, Low-Risk Optimizations (Implement First)

#### 1.1 Defer doc.Regenerate() to Chunk Commit
**Impact:** Very High (40-60% speedup alone)  
**Risk:** Low  
**Affected files:**
- `NWB_PARAMETERS AutoFill (Fast).pushbutton/script.py` — lines ~1481, ~4838
- `NWB Dim AutoFill (Fast).pushbutton/script.py` — lines ~1778, ~2716

**Changes:**
1. In `_write_stringish()`, remove `doc.Regenerate()` from `_verify_after_set()` function
   - Keep the verification logic (checking if value was set)
   - Remove the regenerate call
   - Return immediately after Set() succeeds
2. Verify chunk-level regenerate remains at commit (already present)

**Safety checks:**
- Post-commit `_verify_pending_writes()` uses identical matching logic
- Audit trail and logging unchanged
- Failed verifies still trigger per-element retry

**Testing:**
- Run on sample 50-100 element model, check audit results
- Compare written values between original and Fast versions

---

#### 1.2 Skip Per-Write Immediate Verification
**Impact:** High (15-25% additional speedup)  
**Risk:** Medium (validation happens post-commit instead)  
**Affected files:**
- Both tools, same location as 1.1

**Changes:**
1. In `_write_stringish()`, remove verification block after `_verify_after_set()` succeeds
   - Remove call to `_parameter_value_matches_incoming()` 
   - Remove check and error logging for verification mismatch
   - Return success immediately if Set() succeeds
2. Keep all verification in post-commit phase

**Safety checks:**
- Post-commit `_verify_pending_writes()` verifies ALL writes uniformly
- If a write slipped through, retry logic catches it
- Audit still records outcome (OK, FAIL, etc.)

**Testing:**
- Test 100-element model, verify post-commit verification catches issues
- Disable audit output, run again to confirm speeds match expectations

---

### Phase 2: Medium-Impact, Low-Risk Caching (Implement Second)

#### 2.1 Add Parameter Resolution Cache
**Impact:** Medium-High (20-30% speedup)  
**Risk:** Low  
**Affected files:**
- `NWB_PARAMETERS AutoFill (Fast).pushbutton/script.py` — lines ~1260-1280, ~4705+
- `NWB Dim AutoFill (Fast).pushbutton/script.py` — lines ~1573-1593, ~2545+

**Changes:**
1. Add global cache at module top (near `_PROJECT_CONTEXT_CACHE`):
   ```python
   _PARAM_RESOLUTION_CACHE = {}  # Format: {(owner_id, param_name, target_mode): [candidates]}
   ```

2. In `run()`, reset cache at start:
   ```python
   global _PARAM_RESOLUTION_CACHE
   _PARAM_RESOLUTION_CACHE = {}
   ```

3. Modify `_find_param_target()` to check cache:
   ```python
   def _find_param_target(element, name, target_mode, logger=None):
       # Build cache key from element (or type) owner ID
       owner_id = int(element.Id.IntegerValue) if element else None
       cache_key = (owner_id, name, target_mode)
       
       if cache_key in _PARAM_RESOLUTION_CACHE:
           candidates = _PARAM_RESOLUTION_CACHE[cache_key]
       else:
           candidates = _collect_param_candidates(element, name, logger)
           if owner_id:
               _PARAM_RESOLUTION_CACHE[cache_key] = candidates
       # ... rest of logic
   ```

**Safety checks:**
- Cache scoped to single run (cleared at start)
- Same parameter on same owner always returns same candidates (Revit API guarantee)
- Fallback to full scan if cache miss

**Testing:**
- Run 1000-element model, measure element-level timing
- Verify parameter resolution time drops per element

---

#### 2.2 Add Asset Row Matching Cache (Autofill Only)
**Impact:** Medium (10-20% on Autofill)  
**Risk:** Low  
**Affected files:**
- `NWB_PARAMETERS AutoFill (Fast).pushbutton/script.py` — lines ~3546, ~3806, ~4265-4270, ~4750+

**Changes:**
1. Add global cache at module top:
   ```python
   _ASSET_MATCH_CACHE = {}  # Format: {facts_key: (asset_row, asset_score)}
   ```

2. In `run()`, reset cache at start:
   ```python
   global _ASSET_MATCH_CACHE
   _ASSET_MATCH_CACHE = {}
   ```

3. Create cache key function (from MT12 export pattern):
   ```python
   def _facts_cache_key(facts):
       return (
           facts.get("category_norm", ""),
           facts.get("family_norm", ""),
           facts.get("type_norm", ""),
           facts.get("name_norm", ""),
           facts.get("system_norm", ""),
       )
   ```

4. Modify `_build_target_values()`:
   ```python
   def _build_target_values(element, wb_data):
       facts = _derive_element_facts(element)
       
       # Check cache first
       cache_key = _facts_cache_key(facts)
       if cache_key in _ASSET_MATCH_CACHE:
           asset_row, asset_score = _ASSET_MATCH_CACHE[cache_key]
       else:
           asset_row, asset_score = _select_asset_row(wb_data.asset_rows, facts)
           _ASSET_MATCH_CACHE[cache_key] = (asset_row, asset_score)
       
       # ... rest of logic
   ```

**Safety checks:**
- Cache based on normalized facts (facts uniqueness preserved)
- Same facts always produce same match
- Fallback to full scan if cache miss

**Testing:**
- Run model with many similar elements (same family/type/system)
- Measure speedup on asset matching phase

---

### Phase 3: Optimization (If Needed After Testing)

#### 3.1 Linked Room Search Optimization (Autofill Only, Optional)
**Impact:** High for room-heavy models (20-40% speedup)  
**Risk:** Medium (changes search strategy)  
**Affected files:**
- `NWB_PARAMETERS AutoFill (Fast).pushbutton/script.py` — lines ~2304-2470

**Changes:**
1. Add spatial bucketing for rooms by level:
   ```python
   def _build_linked_room_index_with_level_buckets(link_inst):
       # Existing index build
       index = _build_linked_room_index(link_inst)
       # Group rooms by level, return both index and level_buckets
       return index, level_buckets
   ```

2. In element loop, pre-filter rooms by element's level:
   ```python
   element_level = _get_level_name(element)
   rooms_at_level = level_buckets.get(_normalize_text(element_level), [])
   # Check only rooms at element's level first
   ```

3. Optional: Skip BFS fallback for mechanical/external elements that rarely have linked rooms

**Safety checks:**
- All probing logic unchanged, just filtered
- Audit logging shows room_source unchanged
- Fallback still works if level-based search fails

**Testing:**
- Test element with linked room at same level — verify found quickly
- Test element with no level match — verify fallback works
- Check audit records for room_source consistency

---

## Implementation Dependencies

```
Phase 1.1 (Regenerate defer)
    ↓
Phase 1.2 (Skip verify)
    ↓
Phase 2.1 (Param cache) — Independent
Phase 2.2 (Asset cache) — Autofill only, Independent
    ↓
Phase 3.1 (Room optimization) — Only if Phase 1-2 show room search is still bottleneck
```

**Note:** Phases 1.1 and 1.2 must be completed first; they enable accurate measurement of remaining bottlenecks.

---

## Testing Strategy

### Unit Tests (Per Phase)
1. **Phase 1.1:** Run 100 elements, measure regenerate count
2. **Phase 1.2:** Verify post-commit verify catches all issues
3. **Phase 2.1:** Cache hit rate >80% on 500+ element model
4. **Phase 2.2:** Asset cache hit rate >50% on diverse element model

### Integration Tests
1. Compare audit results (Fast vs. Original) on same model — should be identical
2. Compare parameter values written — must be identical
3. Compare department/subdepartment resolution — must be identical (user requirement)

### Performance Acceptance Tests
- **Autofill:** 50+ element model completes in <50% of original time
- **Dim:** 100+ element model completes in <50% of original time
- Both: No degradation in accuracy or audit quality

### Regression Tests
- Linked room detection still works (Autofill)
- Department resolution preserves accuracy (Autofill)
- Worksharing mode still respected (Both)
- WBS derivation still correct (Autofill)

---

## Rollback Plan

If any phase shows accuracy degradation:
1. Revert that phase's changes
2. Document which optimization caused issue
3. Investigate alternative approach or mark as incompatible
4. Keep all prior phases that passed testing

Each tool maintains a separate `.pushbutton` folder, so rollback is isolated.

---

## Success Criteria

### Minimum (Phase 1 Only)
- ✓ Both Fast tools run without errors
- ✓ Audit results identical to original
- ✓ Parameter values written are identical
- ✓ Speedup ≥ 40% on Autofill, ≥ 30% on Dim

### Ideal (Phases 1-2)
- ✓ All minimum criteria
- ✓ Speedup ≥ 60% on Autofill, ≥ 50% on Dim
- ✓ Cache hit rates >80% on typical models
- ✓ Department/subdepartment resolution unchanged

### Stretch (Phases 1-3)
- ✓ All ideal criteria
- ✓ Speedup ≥ 70% on Autofill, ≥ 60% on Dim
- ✓ Room search optimization validated

---

## Modification Log

| Phase | File | Function | Change | Status |
|-------|------|----------|--------|--------|
| 1.1 | script.py (both) | `_write_stringish()` | Remove per-write regenerate | Pending |
| 1.2 | script.py (both) | `_write_stringish()` | Skip per-write verify | Pending |
| 2.1 | script.py (both) | `_find_param_target()` | Add param cache | Pending |
| 2.2 | script.py (Autofill) | `_build_target_values()` | Add asset cache | Pending |
| 3.1 | script.py (Autofill) | `_find_linked_room_for_element()` | Room optimization | Pending |

---

## References

- Validation findings: `/memories/session/autofill-speed-validation.md`
- Original tools: `NWB_PARAMETERS AutoFill.pushbutton`, `NWB Dim AutoFill.pushbutton`
- Fast versions: `NWB_PARAMETERS AutoFill (Fast).pushbutton`, `NWB Dim AutoFill (Fast).pushbutton`
- Pattern reference (asset cache): `NWB MT12 Schedule Export.source/script.py` lines 599-620

