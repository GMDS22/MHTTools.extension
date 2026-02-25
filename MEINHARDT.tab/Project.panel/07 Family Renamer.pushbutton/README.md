# MHT Family Namer (pyRevit)

This pyRevit extension scans loaded Revit families and suggests standardized names based on templates and smart classification rules defined in `naming_rules.json`.

Installation
- Copy the `pyrevit_family_namer` folder into your pyRevit extensions folder (e.g. `%appdata%\pyRevit\Extensions` or `C:\Users\<user>\pyRevit\extensions`).
- Restart Revit and you'll see a new button labeled "MHT Family Namer".

Usage
- Click the tool to run the script. The current version prints suggestions to the Revit Python Console and the pyRevit logger.
- Edit `naming_rules.json` to customize templates. Use placeholders: `<Family>`, `<Type>`, and `<Param:ParameterName>` for type parameters.

Notes
- This is an initial scaffold. Next steps: add a review UI, apply rename operations inside a Revit Transaction, handle unit conversions and conflicts.
