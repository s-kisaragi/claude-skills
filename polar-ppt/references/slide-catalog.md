# Slide Catalog

## Table of Contents
1. [Slide Types Overview](#slide-types-overview)
2. [Title Slide](#title-slide)
3. [Agenda Slide](#agenda-slide)
4. [Section Divider](#section-divider)
5. [Content Slide](#content-slide)
6. [Chart Slide](#chart-slide)
7. [Two-Column Slide](#two-column-slide)
8. [Three-Column Slide](#three-column-slide)
9. [Table Slide](#table-slide)
10. [Process Flow Slide](#process-flow-slide)
11. [Comparison Matrix Slide](#comparison-matrix-slide)
12. [KPI Dashboard Slide](#kpi-dashboard-slide)
13. [Waterfall Chart Slide](#waterfall-chart-slide)
14. [Key Takeaways Slide](#key-takeaways-slide)
15. [Thank You / End Slide](#thank-you-slide)
16. [Appendix Divider](#appendix-divider)

## Slide Types Overview

Each slide type maps to a `slide_type` string in the JSON input to `generate_pptx.py`.

| Type | `slide_type` | When to Use |
|------|-------------|-------------|
| Title | `title` | First slide: report title, date, author |
| Agenda | `agenda` | Deck outline, recurring with active section highlighted |
| Section Divider | `section_divider` | Visual break between major sections |
| Content | `content` | Bullet points with action title |
| Chart | `chart` | Data visualization with one chart |
| Two-Column | `two_column` | Side-by-side comparison or text+visual |
| Three-Column | `three_column` | Three parallel points or comparison |
| Table | `table` | Structured data in rows/columns |
| Process Flow | `process_flow` | Sequential steps (3-6 boxes with arrows) |
| Comparison Matrix | `matrix` | 2×2 or evaluation matrix with Harvey balls |
| KPI Dashboard | `kpi_dashboard` | Key metrics with cards |
| Waterfall | `waterfall` | Bridge analysis (from → to) |
| Key Takeaways | `takeaways` | Conclusion, next steps |
| Thank You | `thank_you` | End slide |
| Appendix Divider | `appendix_divider` | Marks start of appendix |

## Title Slide

```
slide_type: "title"
fields:
  title: str          # Report title
  subtitle: str       # Client name, project name
  date: str           # Date string
  author: str         # Optional author/team
  confidential: bool  # Show "CONFIDENTIAL" label
```

**Layout (McKinsey)**: Full deep blue background. White serif title centered-left. Subtitle and date below.

**Layout (Accenture)**: Purple gradient or hero image with overlay. White sans-serif title. ">" decorative element.

## Agenda Slide

```
slide_type: "agenda"
fields:
  title: str              # e.g. "Agenda" or action title
  items: list[str]         # 3-5 agenda items
  active_index: int|null   # Which item is current (0-based), null for overview
```

**Layout**: Header bar + numbered vertical list. Active item in bold/color, others in gray. Numbers: large circle or bold numeral.

## Section Divider

```
slide_type: "section_divider"
fields:
  section_number: int|null  # e.g. 01, 02
  section_title: str        # Section name
  subtitle: str|null        # Optional description
```

**Layout (McKinsey)**: Full blue background. Large white text. Section number in oversized font on left.

**Layout (Accenture)**: Full purple background or white with large purple accent. Bold section title.

## Content Slide

```
slide_type: "content"
fields:
  title: str               # Action title (conclusion sentence)
  bullets: list[str|dict]  # Bullet points; dict for nested: {"text": str, "sub": list[str]}
  source: str|null         # Source citation
  exhibit_number: int|null # Exhibit label
```

**Layout**: Header bar with action title. Bullets in content area. 1-2 levels of nesting. Source at bottom.

## Chart Slide

```
slide_type: "chart"
fields:
  title: str               # Action title
  chart_type: str          # "bar", "column", "line", "pie", "stacked_bar", "stacked_bar_100", "scatter"
  categories: list[str]    # X-axis labels
  series: dict[str, list]  # {series_name: [values]}
  highlight_index: int|null # Which category to highlight
  annotation: str|null     # Callout text
  source: str|null
  exhibit_number: int|null
  number_format: str       # "#,##0", "0.0%", "$#,##0", etc.
  y_axis_label: str|null
```

**Layout**: Header bar with action title. Full-width chart below. Exhibit label top-left. Source bottom-left. Callout box near highlight point.

## Two-Column Slide

```
slide_type: "two_column"
fields:
  title: str
  left: dict               # {"heading": str, "bullets": list[str]} or {"chart_type": ..., ...}
  right: dict              # Same structure
  split: str               # "equal" (50/50) or "wide_left" (60/40) or "wide_right" (40/60)
  source: str|null
  exhibit_number: int|null
```

## Three-Column Slide

```
slide_type: "three_column"
fields:
  title: str
  columns: list[dict]      # 3 items, each {"heading": str, "bullets": list[str], "icon": str|null}
  source: str|null
```

## Table Slide

```
slide_type: "table"
fields:
  title: str
  headers: list[str]       # Column headers
  rows: list[list[str]]    # Data rows
  col_widths: list[float]|null  # Relative widths (must sum to 1.0)
  highlight_rows: list[int]|null  # Row indices to highlight
  source: str|null
  exhibit_number: int|null
```

**Layout**: Header bar. Table with blue header row (white text), alternating light/white rows. Clean borders.

## Process Flow Slide

```
slide_type: "process_flow"
fields:
  title: str
  steps: list[dict]        # [{"label": str, "description": str}], 3-6 steps
  source: str|null
```

**Layout**: Header bar. Horizontal boxes connected by arrows. Box fill: light blue or white with blue border. Arrow: gray chevron.

## Comparison Matrix Slide

```
slide_type: "matrix"
fields:
  title: str
  row_headers: list[str]
  col_headers: list[str]
  cells: list[list[str|int]]  # String for text, int 0-4 for Harvey balls (○◔◑◕●)
  cell_type: str             # "text", "harvey_ball", "rag" (red/amber/green)
  source: str|null
```

## KPI Dashboard Slide

```
slide_type: "kpi_dashboard"
fields:
  title: str
  kpis: list[dict]         # [{"label": str, "value": str, "delta": str|null, "trend": "up"|"down"|"flat"}]
  source: str|null
```

**Layout**: Header bar. 3-4 KPI cards in a row. Each card: large value, small label, delta indicator with color.

## Waterfall Chart Slide

```
slide_type: "waterfall"
fields:
  title: str
  items: list[dict]        # [{"label": str, "value": float, "type": "start"|"increase"|"decrease"|"total"}]
  number_format: str
  source: str|null
  exhibit_number: int|null
```

**Layout**: Built with rectangles (python-pptx has no native waterfall). Connector lines between bars. Colors: blue for start/total, green for increase, red for decrease.

## Key Takeaways Slide

```
slide_type: "takeaways"
fields:
  title: str               # e.g. "Key Takeaways" or action title conclusion
  points: list[str]        # 3-5 key takeaway bullet points
  next_steps: list[str]|null  # Optional next steps section
```

## Thank You Slide

```
slide_type: "thank_you"
fields:
  title: str               # "Thank you" or custom
  contact: str|null        # Contact info
```

**Layout**: Full colored background. Large centered white text. Minimal.

## Appendix Divider

```
slide_type: "appendix_divider"
fields:
  title: str               # "Appendix"
```

Same visual as section_divider but labeled "Appendix".
