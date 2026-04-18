# Rebuild Step 3: Dashboard Interface

After building the data and pivot layers, the next step is to design the dashboard interface.

This is where the user interacts with the system.

## Main worksheet

Create a sheet called:

- `Dash`

This will act as the main front-end interface.

## Design approach

The dashboard should not look like a traditional Excel sheet.

Instead, it should behave like a structured interface.

## Main components

### KPI cards
Use shapes and text to display:

- total contracts
- total values
- progress percentages
- key aggregated metrics

### Search area
Add:

- a search text box (`Search_Box`)
- a suggestion list (`ListBox1`)
- a filter selector (`ListBox2`)
- background shapes for dropdowns

### Visual sections
Include:

- charts (linked to pivots)
- progress bars
- grouped information blocks

### Layout structure

Divide the dashboard into:

- header section (title + search)
- KPI section
- visual section
- navigation / interaction elements

## Use of shapes

Use shapes instead of plain cells to:

- create containers
- simulate cards
- build buttons
- improve visual hierarchy

## Interaction elements

Add clickable elements such as:

- labels
- icons
- shapes

These will later be connected to VBA.

## Why this matters

A strong UI design transforms the workbook from a report into an interactive tool.

## Next step

After building the interface, the next step is implementing the search system.
