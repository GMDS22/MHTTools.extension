# Ductulator - How To Use

## Inputs
- **Airflow (L/s)**: design airflow in liters per second.
- **Target Velocity (m/s)**: target duct velocity.
- **Known Duct Width (mm)**: width used to suggest rectangular height.

## Buttons
- **Calculate**: updates all result cards and both velocity tables.
- **Reset**: restores defaults (`240 L/s`, `4 m/s`, `250 mm`).
- **How To Use**: opens quick instructions in the tool.

## Results
- **Flowrate (m3/h)**: airflow converted from L/s.
- **Round Equivalent Diameter**: diameter needed to meet target velocity.
- **Suggested Rect. Size**: width x nearest-height (rounded to nearest 50 mm).

## Tables
- **Speed Through Round Duct**: velocity at the entered airflow for standard diameters.
- **Speed Through Square Duct**: velocity matrix for heights (rows) and widths (columns).

## Notes
- Tool replicates the ENVAR Ductulator logic and arrangement using computed values.
- Values are reference figures for design checks and should be validated against project standards.
