# Progress Bars & Visual Design

One of the distinctive visual elements in this project is the use of custom progress bars inside the Excel dashboard.

Instead of relying on standard Excel visuals, the dashboard uses a combination of charts and shapes to create more advanced and visually appealing progress indicators.

## Design goal

The goal was to represent progress in a clear and intuitive way while maintaining a clean and modern look.

The progress visuals were designed to:
- show execution progress clearly
- differentiate between completed and remaining values
- integrate smoothly with the dashboard layout
- enhance the overall user experience

## Implementation approach

The progress bars are not simple cell-based conditional formatting.

They are built using:

- 100% stacked bar charts
- calculated values (progress vs remaining)
- shape containers for styling
- layout alignment inside the dashboard

## Data logic

Each progress bar is based on two main values:

- Progress value (e.g. execution percentage)
- Remaining value = 1 - progress

These values are usually sourced from PivotTables and calculated using helper formulas such as:

- GETPIVOTDATA
- direct cell references to pivot outputs

## Visual structure

The final appearance of the progress bars is achieved by combining:

- chart elements (for dynamic behavior)
- shapes (for rounded edges and styling)
- positioning (to integrate inside cards or sections)

This creates a more polished look compared to standard Excel bars.

## Multiple states and colors

Some progress visuals include multiple segments or color variations to represent different states.

This allows the dashboard to communicate more than just completion percentage, such as:
- performance levels
- thresholds
- different categories

## Why this matters

Default Excel visuals can be limiting when building modern dashboards.

By customizing the progress bars, the dashboard achieves:
- better visual clarity
- stronger user engagement
- a more professional appearance
- improved readability of key metrics

## Technical note

From the workbook structure, these visuals are powered by PivotTables and helper calculations, then rendered through stacked charts and enhanced using shapes.

This combination allows both flexibility and control over the final design.
