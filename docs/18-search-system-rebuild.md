# Rebuild Step 4: Search System

The search system is one of the most advanced features in the project.

## Main components

- TextBox: `Search_Box`
- ListBox: `ListBox1`
- ListBox: `ListBox2`
- Helper cells: `A26`, `A29`

## Search logic

### A26 → Search Mode
Controls which column to search in.

### A29 → Selected Value
Stores the selected result.

## Behavior flow

1. user types in `Search_Box`
2. VBA searches in `Contracts` sheet
3. results appear in `ListBox1`
4. user selects item
5. value stored in `A29`
6. routing logic is triggered

## Dynamic behavior

- suggestion list resizes automatically
- list hides when no results
- UI updates based on interaction

## Why this matters

This feature replaces manual filtering and improves usability significantly.

## Next step

Connect search results to dashboard actions using VBA routing logic.
