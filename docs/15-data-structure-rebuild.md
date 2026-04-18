# Rebuild Step 1: Source Data Structure

The first step in rebuilding this project is creating a clean and structured source data table.

This table acts as the backbone of the entire solution.  
PivotTables, formulas, search logic, and forms all depend on it.

## Main worksheet

Create a worksheet called:

- `Contracts`

## Main table

Inside this sheet, create an Excel Table named:

- `Contracts_Info`

This should be a structured table, not just a normal range.

## Recommended columns

The original project structure is built around fields like:

- `Project`
- `Area`
- `Company Name`
- `Contract Name`
- `Contract Value`
- `Contract Code`
- `Contract Date`
- `Contract Duration`
- `Contract End Date`
- `Invoice Value`
- `Executed Works Value`
- `Expected Final Contract Value`
- `Progress Percentage`
- `Company type`

## Why use an Excel Table

Using a structured Excel Table makes the solution more scalable and easier to maintain.

It helps with:
- automatic range expansion
- cleaner formulas
- easier PivotTable source management
- more stable VBA interaction

## Calculated fields

Some values can be calculated directly in the source sheet before being used in PivotTables.

Examples include:

### Executed Works Value
Can be calculated from:
- expected final contract value
- progress percentage

### Ratio or percentage fields
Can be calculated from:
- invoice value
- expected final contract value

## Data quality recommendations

To make the workbook reliable:

- keep one row per contract record
- avoid merged cells
- keep column names consistent
- avoid blanks in key fields
- use consistent text formatting
- store percentages as actual numeric percentages

## Why this step matters

If the source table is poorly structured, the rest of the dashboard becomes harder to build and maintain.

A strong source table makes the PivotTables, search logic, and UserForms much easier to implement.

## Next step

After the source data table is ready, the next step is to build the PivotTable logic layer in separate support sheets.
