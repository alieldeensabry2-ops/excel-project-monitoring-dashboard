# How to Rebuild the Search Feature

This document explains how to rebuild the dynamic search system used in this project.

This is one of the most advanced and important features in the dashboard.

## Overview

The search system allows the user to:

- type freely in a search box
- see dynamic suggestions
- select a result
- trigger different actions based on the selection

## Main components

### Controls

- `Search_Box` (TextBox)
- `ListBox1` (suggestions)
- `ListBox2` (search mode selector)

### Helper cells

- `A26` → search mode
- `A29` → selected value

## Step 1: Create the Search Box

Insert an ActiveX TextBox and name it:

- `Search_Box`

This will capture user input.

## Step 2: Create Suggestion List

Insert an ActiveX ListBox and name it:

- `ListBox1`

Set it to:
- initially hidden
- positioned under the search box

## Step 3: Create Search Mode Selector

Insert another ListBox:

- `ListBox2`

This will allow switching between:
- contract search
- project search
- company search

## Step 4: Write Search Logic

Use the `Search_Box_Change` event:

```vba
Private Sub Search_Box_Change()
