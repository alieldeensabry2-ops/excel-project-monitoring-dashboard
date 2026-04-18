# VBA Code Map

This folder contains the VBA components behind the Excel dashboard.

## Confirmed workbook components

### Excel Objects
- ThisWorkbook
- Sheet1
- Sheet2
- Sheet3
- Sheet4

### Workbook Sheets
- Dash
- Contracts
- Main_Dash_pivots
- Projects_Tab_Pivots

### UserForms
- Area_Dash
- IND_Area
- Contracts_Dash

### Standard Modules
- Module1

## Confirmed detected logic and events

The workbook includes VBA procedures related to:

- search box interaction
- suggestion list handling
- label click events
- label hover behavior
- scrollbar behavior
- UserForm initialization
- pivot filtering
- project detail loading
- dynamic content visibility
- chart loading inside forms

## Confirmed procedure names detected from the VBA project

- Search_Box_Change
- Search_Box_GotFocus
- ListBox1_Click
- ListBox2_Click
- FilterPivotByAreaAndProjectAndUpdateLabels
- FilterByLabelClick
- ScrollBar1_Change
- ScrollBar2_Change
- UserForm_Initialize
- show_inDash
- hidemultipages

## Note

This file is based on confirmed component and procedure names extracted from the workbook structure.

Some relationships between workbook sheets and VBA sheet code modules may require direct export from the VBA Editor for full one-to-one verification.
