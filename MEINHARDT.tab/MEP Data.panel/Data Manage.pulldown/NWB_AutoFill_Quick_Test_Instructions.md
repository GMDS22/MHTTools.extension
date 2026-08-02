# NWB AutoFill Quick Test

1. Refresh pyRevit.
2. Run `NWB Data Validator` on a small mixed-discipline selection before auto-fill.
3. For `NWB Data Validator` and `NWB_PARAMETERS AutoFill`, choose the correct `Arch room link` if more than one architectural link is loaded.
4. Open the validator CSV in `NWB Data Validator.pushbutton/logs` and review all rows that are not `PASS`.
5. Resolve `MISSING_PARAMETER` and `READ_ONLY_MISMATCH` findings in the project or family bindings before expecting auto-fill to correct them.
6. Review `NO_EXPECTED_VALUE` findings as missing workbook, material, or room-context source data; they are not safe values for auto-fill to invent.
7. Test both `NWB_PARAMETERS AutoFill` and `NWB Dim AutoFill` on the same selection.
8. Re-run `NWB Data Validator` and confirm expected rows now report `PASS`.
9. If something is wrong, send the latest validator CSV, the relevant autofill log, and one screenshot of the affected element.

## Asset ID Acceptance Check

For Asset Master rows whose `NWB_AssetID` value is the workbook instruction `AssetTypeCode-Building-Level-ElementID`, Parameter AutoFill generates:

`NWB_AssetTypeCode-NWB_WBS01-NWB_WBS02-RevitElementId`

The Revit element ID is padded to at least four digits, so an element with asset type `DPIT`, building `M`, level `GR`, and element ID `1` receives `DPIT-M-GR-0001`. Longer element IDs are not truncated. A fixed Asset Master value, such as `jhu`, is written unchanged. When a matched Asset Master row has a blank `NWB_AssetID`, Parameter AutoFill applies the same generated format; only an unmatched Asset Master row or missing Asset Type Code, building, level, or Revit element ID is skipped.

Logs already capture:

- active view name and view ID
- selected run options
- selected/eligible element IDs
- per-element category, family, type, level, and source details
- linked-room department/building/site context for parameter autofill
- parsed size, width, height, and diameter for dim autofill

The validator CSV additionally captures:

- workbook-derived expected value and stored Revit value for every NWB target parameter
- validation status, precise reason, and suggested corrective action
- selected parameter candidates, scope, shared/non-shared state, and read-only state
- explicit asset-match status/reason, score, selected asset identity, and WBS derivation context
- generated or fixed Asset ID, Asset Master Asset ID rule, and any missing component reason
