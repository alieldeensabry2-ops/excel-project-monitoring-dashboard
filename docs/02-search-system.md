# Search System

One of the main interactive features in this project is the search system built inside the Excel dashboard.

Instead of relying only on manual filtering or standard Excel slicers, the dashboard includes a custom search experience designed to make navigation easier and faster.

## Purpose

The purpose of the search system is to help the user quickly locate a project and immediately update the dashboard based on the selected value.

This makes the dashboard more practical for real use, especially when dealing with multiple projects and repeated lookups.

## Main behavior

The search system is designed to support the following flow:

1. The user starts typing in the search box
2. Matching values are searched dynamically
3. A suggestion list is displayed
4. The user selects a result from the list
5. The selected value is applied to the dashboard filters
6. The related KPIs and details are updated

## Main components

The search feature is built using interactive controls inside Excel, including:

- Search text box
- Suggestion list box
- VBA event procedures
- Filter update logic connected to PivotTables

## Interaction logic

The feature is based on event-driven VBA behavior.

When the user types in the search box, VBA captures the input and checks for matching values.  
The matching items are then pushed into a suggestion list.

The suggestion list is not static. It changes based on the entered text and can be shown or hidden depending on whether matching results are found.

## Suggestion handling

The search system includes dynamic suggestion behavior such as:

- clearing previous suggestions
- adding only matching items
- limiting visible items when needed
- resizing the list depending on the result count
- hiding the list when no results are found

This gives the feature a more application-like feeling inside Excel.

## Why it matters

This feature improves usability significantly because it reduces the need for manual browsing, repetitive filtering, and direct interaction with raw Excel sheets.

It also adds a level of user experience that is uncommon in typical Excel dashboards.

## Technical note

From the workbook structure, the search feature is connected to dashboard controls and filter procedures through VBA events and interactive objects.

This makes it a key part of the dashboard’s usability and overall design.
