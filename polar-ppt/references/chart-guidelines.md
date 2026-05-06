# Chart & Data Visualization Guidelines

## Chart Type Selection

| Data Purpose | Chart Type | `chart_type` value |
|-------------|-----------|-------------------|
| Compare categories | Vertical bar | `column` |
| Rank / long labels | Horizontal bar | `bar` |
| Composition | Stacked bar | `stacked_bar` |
| Share comparison | 100% stacked bar | `stacked_bar_100` |
| Trend over time | Line | `line` |
| Bridge analysis | Waterfall | `waterfall` |
| Correlation | Scatter | `scatter` |
| Simple split (≤3 segments) | Pie | `pie` |

**Avoid**: 3D effects, pie charts with >4 segments, radar charts (unless specifically requested), dual-axis charts.

## Formatting Rules

### Axes
- X-axis labels: 9-10pt, `text_mid` color
- Y-axis labels: 9-10pt, `text_mid` color
- Axis lines: 0.75pt, `separator` color; remove top and right borders
- Gridlines: horizontal only, 0.5pt, `gridline` color; never vertical

### Data Labels
- Font: 9pt Arial, `text_dark` color
- Position: outside end for bar/column, above for line
- Bold the key data point referenced in the action title
- Include units (%, $M, etc.)

### Legend
- Prefer direct labeling (no legend) when ≤3 series
- If needed: horizontal, below chart, 9pt, no border
- Never place legend inside chart area

### Annotations / Callout Boxes
- Thin border box (1pt, `primary` color) or `bg_light` fill
- Text: 10pt Arial
- Connector arrow (thin gray line) to data point
- Content: brief insight text with key figure bolded
- Example: "Revenue from Asia grew **3×** in 2 years"

## Color Usage

### Single Series
Use `primary` color only.

### Two Series
`primary` + `accent2` (or lighter variant of primary).

### Three+ Series
Follow the chart series order defined in design-system.md.

### Highlight Technique ("Gray Out")
- Key bar/segment: `primary` (full color)
- All other bars: `#D9D9D9` (grayed out)
- This immediately draws attention to the insight

### Positive / Negative
- Positive: `positive` (green)
- Negative: `negative` (red)
- Neutral/baseline: `text_mid` (gray)

## Waterfall Chart Specifics

Since python-pptx has no native waterfall, build with rectangles:

- **Start bar**: `primary` color, anchored at baseline
- **Increase bars**: `positive` color, floating from previous cumulative
- **Decrease bars**: `negative` color, floating downward from previous cumulative
- **Total bar**: `primary` color, anchored at baseline
- **Connector lines**: thin gray dashed lines between bar tops
- **Data labels**: above each bar, showing the delta value with +/- sign
- **Bar width**: 70% of available space per category
- **Gap between bars**: 30%

## Harvey Balls & RAG Status

### Harvey Balls (qualitative assessment)
```
0 = ○ (empty)
1 = ◔ (quarter)
2 = ◑ (half)
3 = ◕ (three-quarter)
4 = ● (full)
```
Render as circles with partial fills using MSO_SHAPE.OVAL + arc shapes, or use Unicode characters in text.

### RAG Status (traffic light)
```
R (Red)   = negative color, ● or filled square
A (Amber) = accent3/gold color, ● or filled square
G (Green) = positive color, ● or filled square
```

### Japanese Evaluation Symbols
```
◎ = Excellent (primary color)
○ = Good (positive color)
△ = Fair (accent3/gold color)
× = Poor (negative color)
```

## Content Density Rules

- **One chart per slide** (rarely two; never three)
- Chart occupies **60-70%** of content area
- Maximum **75-100 words** per slide (excluding chart labels)
- **Glance test**: viewer should understand the key message within 3 seconds
- For data-heavy appendix slides, density can increase but still maintain clarity
