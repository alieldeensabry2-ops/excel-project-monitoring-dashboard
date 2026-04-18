# Pivot Architecture

The dashboard is built on a pivot-driven analytical structure.

Instead of placing all calculations directly on the dashboard sheet, the workbook uses separate pivot sheets to handle aggregation, filtering, and metric preparation.

This approach makes the solution more organized, maintainable, and scalable.

## Main pivot layers

The workbook uses dedicated sheets for pivot logic:

- `Main_Dash_pivots`
- `Projects_Tab_Pivots`

These sheets act as the calculation and aggregation engine behind the visible dashboard.

## Why PivotTables were used

PivotTables were used to make the dashboard more dynamic and efficient.

They help in:
- aggregating project data
- summarizing contract values
- calculating totals and counts
- feeding KPIs and charts
- supporting dashboard filtering

This reduces the need for heavy formulas directly on the front-end dashboard.

## Workbook structure logic

The workbook follows a layered structure:

### 1. Data layer
The raw source data is stored in a structured Excel table in the `Contracts` sheet.

### 2. Logic layer
PivotTables process and summarize the source data in hidden or support sheets.

### 3. Presentation layer
The `Dash` sheet displays the final user-facing dashboard elements such as KPIs, charts, search, and details.

## Centralized calculation flow

The pivot layer works as an intermediate engine between source data and the dashboard.

The flow is:

1. Source data is stored in structured table format
2. PivotTables summarize and group the data
3. Helper formulas such as `GETPIVOTDATA` read values from pivots
4. Dashboard elements display the final outputs

This makes the dashboard easier to update and more stable than relying on direct cell logic everywhere.

## Benefits of this design

Using a pivot-based architecture provides several advantages:

- better separation between data and presentation
- easier maintenance
- more structured filtering logic
- reusable calculations
- stronger support for dashboard visuals
- improved performance compared to crowded front-end formulas

## Dashboard dependency on pivots

Several dashboard metrics and visuals depend on pivot outputs rather than manual formulas.

This includes:
- KPI cards
- project-level summaries
- progress values
- chart source ranges
- detail views

## Why this matters

A lot of Excel dashboards place everything in one sheet, which makes them harder to maintain and scale.

In this project, the pivot architecture helps turn the workbook into a more system-like solution rather than a single reporting page.

## Technical note

From the workbook structure, multiple PivotTables are distributed across support sheets and connected to a central data source.

This architecture is one of the main reasons the dashboard is able to support filtering, search-driven updates, and drill-down views effectively.
