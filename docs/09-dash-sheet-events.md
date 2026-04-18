# Dash Sheet Events

The dashboard sheet is one of the core interaction layers in this project.

It is responsible for handling direct user actions from the visible dashboard interface.

## Purpose

The purpose of the dashboard sheet code is to connect front-end controls with the backend logic of the workbook.

This includes reacting to user input and triggering updates across the dashboard.

## Main interaction types

Based on the workbook structure and detected VBA procedure names, the dashboard sheet includes event-driven behavior related to:

- search box input
- suggestion list selection
- label click actions
- hover-based interaction
- dashboard-triggered filtering
- opening project detail views

## Search-related events

The dashboard includes events connected to the search system, such as:

- `Search_Box_Change`
- `Search_Box_GotFocus`
- `ListBox1_Click`
- `ListBox2_Click`

These events help create a more dynamic and user-friendly search experience inside Excel.

## Label and UI events

The dashboard also includes interactive label-based events such as:

- click events
- mouse move events

These are used to make the dashboard more interactive and reduce the need for raw sheet navigation.

## Dashboard role in the architecture

The dashboard sheet acts as a bridge between:

- visible user actions
- PivotTable filtering logic
- UserForm opening and updates
- dynamic UI state changes

## Why this matters

Without this event layer, the dashboard would behave like a static visual report.

With sheet-level VBA events, the dashboard becomes interactive and much closer to a lightweight application.

## Export note

The repository includes the exported dashboard sheet code as a class module file.

This file is important for understanding how the front-end user actions are connected to the workbook logic.
