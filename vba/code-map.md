# VBA Code Map

This folder contains the VBA components behind the Excel dashboard.

## Detected workbook components

### Excel Objects
- ThisWorkbook
- Sheet1 (Dash)
- Sheet2 (Contracts)
- Sheet3 (Main_Dash_pivots)
- Sheet4 (Projects_Tab_Pivots)

### UserForms
- Area_Dash
- IND_Area
- Contracts_Dash

### Standard Modules
- Module1

## Detected logic and events

The workbook includes event-driven VBA procedures related to:

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

## Example detected procedure names

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

## Purpose

This file acts as a map for the VBA project structure before exporting all source code files into this repository.
