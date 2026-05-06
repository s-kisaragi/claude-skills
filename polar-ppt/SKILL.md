---
name: polar-ppt
description: Generate consulting-grade PowerPoint presentations (PPTX) in McKinsey or Accenture style using python-pptx. Supports strategy proposals, market analyses, management reports, and more with action titles, data visualizations, and professional layouts. Bilingual (Japanese/English). Use when asked to create a PowerPoint, PPT, PPTX, presentation, slide deck, consulting report, 資料作成, プレゼン作成, スライド作成, パワポ, or any request involving professional business presentations.
---

# Consulting-Grade PPTX Generator

Generate McKinsey/Accenture-level PowerPoint presentations with python-pptx.

## Prerequisites

Ensure python-pptx is installed in the project:

```bash
uv add python-pptx  # or: pip install python-pptx
```

## Workflow

### Phase 1: Scope & Storyline

Clarify with the user:

1. **Topic & purpose** — What the presentation covers
2. **Style** — `mckinsey` (default) or `accenture`
3. **Language** — Japanese, English, or bilingual
4. **Deck type** — Strategy proposal, market analysis, management report, or custom
5. **Audience** — Executive, operational, external

Then design the **storyline** using the Pyramid Principle:
- Draft all **action titles** (conclusion-driven sentences, NOT topic labels)
- Verify the titles tell a coherent story when read sequentially
- Present the storyline outline to the user for approval

See [references/storytelling.md](references/storytelling.md) for frameworks: Pyramid Principle, SCQA, MECE, deck structure templates.

### Phase 2: Slide Design

For each slide, determine:
- **Slide type** from the catalog
- **Content** (text, data, chart specs)
- **Layout** (full width, 2-col, 3-col)

Present the slide-by-slide plan to the user. Each slide entry should include the action title and content summary.

See [references/slide-catalog.md](references/slide-catalog.md) for all available slide types and their JSON field specifications.

### Phase 3: PPTX Generation

Build the JSON structure and run the generation script:

```bash
uv run python ~/.claude/skills/polar-ppt/scripts/generate_pptx.py input.json output.pptx --style mckinsey
```

The script reads a JSON file with this structure:

```json
{
  "style": "mckinsey",
  "slides": [
    {
      "slide_type": "title",
      "title": "Digital Transformation Strategy",
      "subtitle": "Client Name",
      "date": "February 2026",
      "confidential": true
    },
    {
      "slide_type": "content",
      "title": "Three factors drive the need for digital transformation",
      "bullets": [
        "Customer expectations have shifted to digital-first interactions",
        {"text": "Legacy systems create operational bottlenecks", "sub": ["ERP modernization overdue", "Integration gaps across channels"]},
        "Competitors have invested 2-3x more in digital capabilities"
      ],
      "source": "Industry interviews; McKinsey analysis",
      "exhibit_number": 1
    },
    {
      "slide_type": "chart",
      "title": "Revenue grew 23% YoY driven by digital channel expansion",
      "chart_type": "column",
      "categories": ["2022", "2023", "2024", "2025"],
      "series": {"Revenue ($M)": [120, 135, 152, 187]},
      "highlight_index": 3,
      "annotation": "Digital channels contributed 60% of 2025 growth",
      "number_format": "$#,##0",
      "source": "Company financials",
      "exhibit_number": 2
    }
  ]
}
```

### Available Slide Types

| Type | `slide_type` | Key Fields |
|------|-------------|------------|
| Title/Cover | `title` | title, subtitle, date, author, confidential |
| Agenda | `agenda` | title, items[], active_index |
| Section Divider | `section_divider` | section_number, section_title, subtitle |
| Content (Bullets) | `content` | title, bullets[] |
| Chart | `chart` | title, chart_type, categories, series, highlight_index, annotation |
| Two-Column | `two_column` | title, left{}, right{}, split |
| Three-Column | `three_column` | title, columns[] |
| Table | `table` | title, headers[], rows[][], col_widths[], highlight_rows[] |
| Process Flow | `process_flow` | title, steps[{label, description}] |
| Matrix / Harvey Ball | `matrix` | title, row_headers, col_headers, cells[][], cell_type |
| KPI Dashboard | `kpi_dashboard` | title, kpis[{label, value, delta, trend}] |
| Waterfall | `waterfall` | title, items[{label, value, type}] |
| Key Takeaways | `takeaways` | title, points[], next_steps[] |
| Thank You | `thank_you` | title, contact |
| Appendix Divider | `appendix_divider` | title |

All content slides support: `source`, `exhibit_number` (optional).

### Chart Types

`column`, `bar`, `line`, `line_markers`, `pie`, `stacked_bar`, `stacked_bar_100`, `scatter`, `area`

See [references/chart-guidelines.md](references/chart-guidelines.md) for chart selection and formatting rules.

## Design Reference

For detailed design specifications (colors, typography, spacing, layout grid):

- [references/design-system.md](references/design-system.md) — Color palettes, typography, layout grid, common elements
- [references/slide-catalog.md](references/slide-catalog.md) — Slide type specifications and JSON schemas
- [references/storytelling.md](references/storytelling.md) — Pyramid Principle, action titles, MECE, deck structures
- [references/chart-guidelines.md](references/chart-guidelines.md) — Chart selection, formatting, annotation patterns

## Quality Checklist

Before delivering the PPTX, verify:

1. Every slide has an **action title** (conclusion sentence, not topic label)
2. Reading all titles sequentially tells a **coherent story**
3. Content is **MECE** (no overlaps, no gaps)
4. Charts support their action titles (the data proves the claim)
5. Source citations present on all data slides
6. Consistent style throughout (no mixed fonts/colors)
7. Maximum ~100 words per slide (excluding chart labels)
