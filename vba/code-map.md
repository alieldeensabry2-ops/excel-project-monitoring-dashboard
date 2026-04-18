# VBA Code Map

This file maps the exported VBA components currently included in the repository.

It is intended to help readers understand the structure of the Excel VBA project and navigate the source code more easily.

## Confirmed exported components

### Standard Module
- `Module1.bas`

### Excel Objects
- `ThisWorkbook.cls`
- `Sheet1.cls` *(dashboard sheet code export)*

### UserForms
- `Area_Dash.frm`
- `Area_Dash.frx`
- `Contracts_Dash.frm`
- `Contracts_Dash.frx`
- `IND_Area.frm`
- `IND_Area.frx`

## Confirmed workbook structure detected from the file

### Workbook sheets
- `Dash`
- `Contracts`
- `Main_Dash_pivots`
- `Projects_Tab_Pivots`

### VBA project components detected
- `ThisWorkbook`
- `Sheet1`
- `Sheet2`
- `Sheet3`
- `Sheet4`
- `Area_Dash`
- `Contracts_Dash`
- `IND_Area`
- `Module1`

## Confirmed detected procedure and event names

The workbook contains confirmed VBA procedure names related to interaction and filtering, including:

- `Search_Box_Change`
- `Search_Box_GotFocus`
- `ListBox1_Click`
- `ListBox2_Click`
- `FilterPivotByAreaAndProjectAndUpdateLabels`
- `FilterByLabelClick`
- `ScrollBar1_Change`
- `ScrollBar2_Change`
- `UserForm_Initialize`
- `show_inDash`
- `hidemultipages`

## Component roles

### Module1.bas
Contains shared VBA procedures and helper logic used by the workbook.

### ThisWorkbook.cls
Contains workbook-level behavior such as startup logic or global workbook events.

### Sheet1.cls
Contains dashboard interaction logic and event-driven behavior connected to controls on the main dashboard.

### Area_Dash.frm / Area_Dash.frx
Represent the main project drill-down form used to display dynamic project details and interactive content.

### Contracts_Dash.frm / Contracts_Dash.frx
Represent a dedicated contract-related form used for deeper contract interaction or detail presentation.

### IND_Area.frm / IND_Area.frx
Represent an additional form used for a focused area-related or specialized interaction view.

## Notes

This file only includes components and names that are either:

- directly exported from the VBA project
- or directly detected from the workbook structure

Some deeper mappings between procedures and exact components may require line-by-line reading of the exported code files.
