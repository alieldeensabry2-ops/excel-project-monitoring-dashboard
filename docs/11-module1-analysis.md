# Module1 Analysis

`Module1` contains a major part of the interactive logic behind the dashboard.

It is not a simple helper module.  
Instead, it acts as a controller for project drill-down behavior, dashboard interaction, and `Area_Dash` data loading.

## Main responsibilities

The module handles:

- opening specialized views
- loading the `Area_Dash` form
- filtering PivotTables by selected project
- updating dynamic labels and images
- populating KPI boxes
- exporting and loading chart-based visuals
- handling company-level detail filtering
- resetting and hiding UI elements

## Main procedures

### `Lisy2_Show`
Toggles visibility of dashboard UI elements related to `ListBox2`.

### `Industrial_Area_Dash`
Opens the `IND_Area` UserForm.

### `Greenhouses_A3`
Loads a project/area view linked to cell `A3` and prepares the `Area_Dash` form.

### `Wholesale_market_A2`
Loads a project/area view linked to cell `A2` and prepares the `Area_Dash` form.

### `hide_buttons`
Resets visibility and appearance of image and label controls inside `Area_Dash`.

### `FilterPivotByAreaAndProjectAndUpdateLabels`
This is one of the core procedures in the project.

It:
- reads a selected project from a worksheet cell
- filters multiple PivotTables
- loads project labels into the form
- updates KPI values
- exports a chart as an image
- loads that image into the form
- initializes visible multipage sections
- updates contract-related details
- finally shows `Area_Dash`

### `hidemultipages`
Hides all MultiPage controls in `Area_Dash`.

### `FilterByLabelClick`
Applies additional filtering based on a selected label, likely representing a company, then updates detailed contract-related labels inside `Area_Dash`.

### `show_inDash`
Loads all related items into the form without applying a company-specific filter.

### `PlayVideo`
Triggers an embedded OLE object from the workbook.

## Technical characteristics

The module is strongly tied to:

- named PivotTables
- specific worksheet ranges
- UserForm control naming patterns
- chart export logic
- event-driven dashboard interaction

## Why this module matters

This module represents a central control layer between:
- pivot-based backend logic
- form-based UI
- user-triggered interaction

It is one of the main reasons the workbook behaves like an application rather than a static report.
