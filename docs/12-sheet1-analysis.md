# Sheet1 (Dash) Analysis

`Sheet1` contains the event-driven interaction logic behind the main dashboard sheet.

It is one of the most important code components in the project because it controls the search behavior, suggestion handling, selection routing, and contract-level detail launching.

## Main responsibilities

The sheet code handles:

- dynamic search box behavior
- suggestion list population
- search mode selection
- result selection handling
- routing selected values to the correct workflow
- opening contract details through `Contracts_Dash`
- resetting search UI elements

## Main procedures

### `Search_Box_Change`
Builds dynamic search suggestions based on user input and adjusts the suggestion list visibility and size.

### `ListBox1_Click`
Handles the user selection from the search suggestions and stores the selected value for further processing.

### `RemoveSpaceBetweenNumberAndWord`
Cleans selected text to ensure consistent formatting.

### `ListBox2_Click`
Handles selection of the search mode and stores it in a control cell.

### `Search_Selection`
Acts as a routing procedure that decides what action to take based on the selected search mode and selected value.

### `run_Contract`
Loads contract-level data from the `Contracts_Info` table into `Contracts_Dash` and displays a chart-based visual.

### `cancel_Search`
Clears and hides search-related UI elements.

### `Search_Box_GotFocus`
Improves search UX by showing a clear/cancel icon when appropriate.

## Technical characteristics

This code is strongly tied to:

- ActiveX controls on the dashboard
- worksheet helper cells such as `A26` and `A29`
- workbook shapes used as UI elements
- contract data stored in the `Contracts` sheet
- chart export and image loading logic
- downstream calls to `Area_Dash`, `Contracts_Dash`, and `IND_Area`

## Why this file matters

Without `Sheet1`, the dashboard would lose most of its interactive behavior.

This file is a core part of the application-like experience of the project, especially around search and navigation.
