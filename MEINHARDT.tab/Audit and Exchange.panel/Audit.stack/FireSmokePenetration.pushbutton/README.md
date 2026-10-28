# MHT Fire & Smoke Penetration Detector

## Status

Audit-gated implementation: geometry and CSV/report logic can be run in a Revit model, but marker placement is deliberately blocked until an approved MHT marker family and parameter schema are supplied.

## Architecture

Native pyRevit was selected because this repository contains no Dynamo graphs and has established Revit API, WPF, transaction, link, logging, and CSV patterns. The implementation separates read-only analysis from model modification.

## Required configuration

1. Enter the exact names of project parameters that explicitly classify each wall as fire-rated, smoke-rated, both, or not rated.
2. Missing classification data remains `UNKNOWN`; it is never treated as a negative result.
3. Supply an approved marker family and family symbol in the tool UI. The family must not be substituted automatically.
4. Approve the marker family category, shared parameters, GUIDs, schedule fields, visibility, and reconciliation rules before marker placement is enabled.

No supported shared parameter or approved marker family was found in the repository audit, so parameter names and marker definitions cannot be invented.

## Detection

- Host walls and host ducts are collected from the active model or active view.
- Revit link instances are collected independently; unloaded links are reported.
- Link transforms are applied to linked wall solids before solid intersection checks.
- Nested links are not followed.
- Confirmed intersections require non-empty host and linked wall solids and a non-disjoint `Solid.Intersect` result.
- The tool does not infer engineering damper requirements or automatically place dampers.

## Known limitations

- Precise entry/exit point sampling and wall thickness require additional validated solid-boundary implementation in Revit.
- Marker reconciliation/placement is not enabled until the approved family and parameter contract exists.
- Live Revit integration tests must be run in a compatible supported Revit environment before production use.
- Phases and design options are currently collected in the UI but require release-specific validation.
- Opening/sleeve classification is not implemented in this audit-gated release.

## Usage

1. Open an eligible Revit project.
2. Enter exact wall classification parameter names approved for the project.
3. Run `Analyze`.
4. Review unknown/unclassified and verified findings in the diagnostics area.
5. Export findings as UTF-8 CSV.
6. Do not place markers until the approved family, symbol, shared parameters, and placement workflow have been validated.

## Tests

Run the dependency-free core tests with CPython:

`python tests_core.py`

The tests cover wall-classification normalization, unit conversion, bounding boxes, segment intersections, stable IDs, and CSV export. Live Revit tests remain required.
