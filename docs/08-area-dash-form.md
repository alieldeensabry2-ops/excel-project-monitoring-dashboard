# Area_Dash UserForm

`Area_Dash` is one of the core interactive components in this project.

It acts as a drill-down interface that displays project-specific details after the user selects a project from the dashboard.

## Purpose

The purpose of this form is to provide a focused details view without overcrowding the main dashboard.

It allows the user to move from a high-level dashboard summary to a deeper project-level view.

## Main responsibilities

Based on the workbook structure and exported form component, `Area_Dash` is responsible for:

- displaying selected project details
- showing related labels and values
- presenting grouped detail sections
- managing scrollable content
- displaying chart-based visuals inside the form
- handling multi-section content through form controls

## Interaction pattern

The form is triggered after a project selection or dashboard interaction.

A filtering process updates the pivot-based backend, then the form is populated dynamically with the relevant values.

## Dynamic controls

The form includes multiple dynamic controls such as:

- labels
- images
- scrollbars
- multipage controls
- picture-based visual areas

The form logic shows or hides controls depending on the available data.

## Why this form is important

This form is one of the main reasons the workbook behaves more like an application than a static Excel dashboard.

It separates overview and detail interaction in a cleaner and more scalable way.

## Export note

The repository includes:

- `Area_Dash.frm`
- `Area_Dash.frx`

These files represent the exported UserForm and its associated resources.
