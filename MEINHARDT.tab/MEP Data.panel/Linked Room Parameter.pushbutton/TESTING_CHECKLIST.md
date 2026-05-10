# Linked Room Parameter Transfer - Testing Checklist

## Preconditions
- Host model open in Revit.
- At least one architectural link loaded with Rooms.
- At least one target MEP element category present in host model.

## Smoke Tests
1. Launch tool from MEINHARDT.tab > MEP Data.panel.
2. Confirm linked model dropdown populates with loaded links.
3. Confirm no crash when closing with Cancel.

## Linked Room Selection
1. Select a link in dropdown.
2. Click Select Linked Room.
3. Pick a linked Room from selected link:
   - Expect room info panel updates (name/number/level/link).
   - Expect room parameter list populates.
4. Pick a linked element that is not a Room:
   - Expect validation message.
5. Pick a Room from a different link than dropdown:
   - Expect validation message.

## Category and Element Selection
1. Use category search to find a category.
2. Toggle Select All / Deselect All.
3. Click Use Current Selection with valid preselected elements:
   - Expect element summary by category.
4. Click Select Elements From Model and pick mixed categories:
   - Expect only selected categories are allowed.

## Parameter Discovery and Mapping
1. Verify writable target parameters list updates.
2. Toggle Show Non-Common Parameters:
   - OFF: only common writable params.
   - ON: union of writable params.
3. Verify auto mappings appear.
4. Use Room Param + Target combos and Add/Update:
   - Expect mapping line updates.
5. Remove a mapping:
   - Expect mapping marked unmapped.
6. Auto Match and Clear Mappings buttons:
   - Expect mapping list updates accordingly.

## Transfer Behavior
1. Set duplicate mode Overwrite and transfer.
2. Set duplicate mode Skip and transfer.
3. Set duplicate mode Append and transfer on text parameters.
4. Toggle Skip Empty Values and verify behavior.
5. After transfer, verify summary includes:
   - elements updated
   - writes
   - failures
   - skipped
   - log file path

## Data Validation
1. Spot-check updated elements in Properties palette.
2. Verify values map to expected target parameters.
3. Verify non-existing target parameters are skipped safely.
4. Verify read-only parameters are not written.

## Error Handling
1. Unload all links and run tool:
   - Expect graceful error and exit.
2. Attempt transfer without selecting room:
   - Expect warning.
3. Attempt transfer without selecting elements:
   - Expect warning.
4. Attempt transfer with no mappings:
   - Expect warning.

## Performance Sanity
1. Run transfer on small set (10-20 elements).
2. Run transfer on medium set (200+ elements).
3. Confirm UI remains responsive between actions and no Revit freeze persists.

## Acceptance Criteria
- Tool completes end-to-end workflow without unhandled exceptions.
- Transfer results are consistent with selected mappings.
- Validation/warning messages are clear and actionable.
- Log file is created after each transfer.
