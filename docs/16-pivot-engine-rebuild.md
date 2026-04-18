# Rebuild Step 2: Pivot Engine

After building the source data table, the next step is to create the pivot-based logic layer.

This layer is responsible for transforming raw records into summarized metrics, grouped values, and dashboard-ready outputs.

## Recommended support sheets

Create two support sheets:

- `Main_Dash_pivots`
- `Projects_Tab_Pivots`

These sheets will act as the backend analytical engine of the workbook.

## Why use separate pivot sheets

Keeping PivotTables away from the main dashboard helps:

- keep the dashboard clean
- separate logic from presentation
- simplify maintenance
- make formulas easier to manage
- improve organization of calculations

## Main role of `Main_Dash_pivots`

This sheet should contain PivotTables that support the main dashboard.

Typical outputs include:

- counts of projects
- counts of contracts
- total invoice values
- total executed works values
- grouped progress-related data
- chart source ranges for dashboard visuals

## Main role of `Projects_Tab_Pivots`

This sheet should contain PivotTables that support drill-down and detail views.

Typical outputs include:

- project-specific labels
- company-specific breakdowns
- contract lists
- detailed KPI outputs
- chart source ranges used in UserForms or detailed sections

## Recommended pivot design principles

When building the pivot layer:

- use the same source table for all pivots
- keep PivotTable names clear when possible
- reserve helper cells for KPI extraction
- keep chart source ranges close to related pivots
- avoid mixing raw manual data with pivot outputs

## Connection to formulas

The pivot layer can feed the dashboard through:

- direct cell references
- helper formulas
- `GETPIVOTDATA`

This allows dashboard cards and visuals to stay dynamic without pulling directly from raw records.

## Why this step matters

The pivot engine is what turns the workbook from a simple data file into an analytical dashboard.

Without this layer, the dashboard would depend too heavily on front-end formulas and manual logic.

## Next step

After the pivot engine is ready, the next step is to build the dashboard interface and visual layout.
