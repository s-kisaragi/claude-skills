# Design System Reference

## Table of Contents
1. [Style Presets](#style-presets)
2. [Color Palettes](#color-palettes)
3. [Typography](#typography)
4. [Layout Grid](#layout-grid)
5. [Common Elements](#common-elements)

## Style Presets

Two built-in styles: `mckinsey` and `accenture`. User may also specify `custom`.

## Color Palettes

### McKinsey Style

| Token | Hex | Role |
|-------|-----|------|
| `primary` | `#051C2C` | Header bar, primary emphasis |
| `primary_light` | `#003A70` | Accent elements |
| `accent1` | `#2251FF` | Chart highlight, links |
| `accent2` | `#009AA6` | Teal, secondary chart series |
| `accent3` | `#D4A843` | Gold, tertiary series |
| `text_dark` | `#333333` | Body text |
| `text_mid` | `#666666` | Secondary text |
| `text_light` | `#999999` | Footnotes, source |
| `bg_light` | `#F2F2F2` | Table striping, fills |
| `bg_blue` | `#D6E8F5` | Light blue fills |
| `positive` | `#4A8C2A` | Growth, success |
| `negative` | `#D04A3E` | Decline, alert |
| `gridline` | `#D9D9D9` | Chart gridlines |
| `separator` | `#B0B0B0` | Divider lines |
| `white` | `#FFFFFF` | Backgrounds, text on dark |

**Chart series order**: `primary` → `accent1` → `accent2` → `accent3` → `text_mid` → `bg_light`

**Highlight ("gray out") technique**: Key bar in `primary`, all others in `#D9D9D9`.

### Accenture Style

| Token | Hex | Role |
|-------|-----|------|
| `primary` | `#A100FF` | Purple, primary brand |
| `primary_dark` | `#7500C0` | Dark purple |
| `accent1` | `#FF00C8` | Hot pink |
| `accent2` | `#00BAFF` | Electric blue |
| `accent3` | `#00D7A0` | Green teal |
| `text_dark` | `#000000` | Body text |
| `text_mid` | `#505050` | Secondary text |
| `text_light` | `#808080` | Footnotes |
| `bg_light` | `#F5F5F5` | Light fills |
| `positive` | `#00D7A0` | Growth |
| `negative` | `#FF4444` | Decline |
| `gridline` | `#E0E0E0` | Chart gridlines |
| `white` | `#FFFFFF` | Backgrounds |

**Chart series order**: `primary` → `accent2` → `accent1` → `accent3` → `text_mid`

**Gradient patterns**: Purple→Pink gradient (`#A100FF` → `#FF00C8`) for hero elements.

## Typography

### McKinsey

| Element | Font | Size (pt) | Weight | Color |
|---------|------|-----------|--------|-------|
| Action title | Georgia | 22-24 | Bold | white (on bar) or `text_dark` |
| Subtitle | Arial | 14-16 | Regular/Italic | `text_dark` |
| Body | Arial | 11-12 | Regular | `text_dark` |
| Bullet L1 | Arial | 11 | Regular | `text_dark` |
| Bullet L2 | Arial | 10 | Regular | `text_mid` |
| Chart label | Arial | 9-10 | Regular | `text_mid` |
| Exhibit label | Arial | 8 | Bold | `primary` |
| Source/footnote | Arial | 7 | Regular | `text_light` |

Line spacing: 1.0–1.15 for body. Bullet markers: `—` (L1), `–` (L2), `·` (L3).

### Accenture

| Element | Font | Size (pt) | Weight | Color |
|---------|------|-----------|--------|-------|
| Action title | Arial | 24-28 | Bold | `text_dark` or white |
| Subtitle | Arial | 16-18 | Regular | `text_mid` |
| Body | Arial | 11-12 | Regular | `text_dark` |
| Bullet L1 | Arial | 11 | Regular | `text_dark` |
| Chart label | Arial | 9-10 | Regular | `text_mid` |
| Source/footnote | Arial | 7-8 | Regular | `text_light` |

Bullet marker: `>` (Accenture's brand symbol) or standard bullet. Line spacing: 1.15-1.2.

### Japanese Font Fallback

For Japanese text, append Japanese font to font spec:
- McKinsey JP: Georgia (title) + Yu Gothic or Meiryo (body)
- Accenture JP: Arial (title) + Yu Gothic or Meiryo (body)

Set both `font.name` (Latin) and East Asian font via XML when Japanese text detected.

## Layout Grid

### Slide Dimensions

- **16:9 Widescreen**: 13.333" × 7.5" (default)
- **4:3 Standard**: 10" × 7.5" (legacy)

### Margins (16:9, inches)

| Zone | Value |
|------|-------|
| Left margin | 0.5" |
| Right margin | 0.5" |
| Header bar height | 0.85"–1.0" |
| Content top | 1.15" (below header) |
| Content bottom | 6.7" (above footer) |
| Footer zone | 6.7"–7.3" |
| Usable content width | 12.33" |
| Usable content height | ~5.5" |

### Column Systems

| Layout | Col width | Gutter | Positions (left edges) |
|--------|-----------|--------|----------------------|
| Full width | 12.33" | — | 0.5" |
| 2-col equal | 5.9" | 0.5" | 0.5", 6.9" |
| 2-col 60/40 | 7.2" / 4.6" | 0.5" | 0.5", 8.2" |
| 3-col equal | 3.78" | 0.4" | 0.5", 4.68", 8.86" |
| Left sidebar | 3.5" / 8.3" | 0.5" | 0.5", 4.5" |

## Common Elements

### Header Bar (McKinsey)

```
Position: (0, 0)
Size: (slide_width, 0.85")
Fill: primary (#051C2C) solid, no border
Title text: inside bar, left-aligned at (0.5", 0.15"), Georgia 22pt bold white
```

### Header Bar (Accenture)

```
Position: (0, 0)
Size: (slide_width, 0.9")
Fill: white or gradient purple
Title text: below bar at (0.5", 1.0"), Arial 24pt bold black
Accent line: thin purple line below header (0.5", 0.9") width=12.33" color=primary
```

### Footer / Source Line

```
Source text: (0.5", 6.8") font=7pt Arial color=text_light
Page number: (12.5", 7.05") font=7pt Arial right-aligned color=text_light
Separator line: optional 0.5pt line at y=6.7" color=separator
```

### Exhibit Label (McKinsey only)

```
Text: "Exhibit {N}" or "EXHIBIT {N}"
Position: (0.5", 1.05")
Font: Arial 8pt bold color=primary
```

### Callout Box

```
Border: 1pt primary color
Fill: white or bg_light
Text: 10pt Arial
Position: adjacent to chart data point with connector arrow
```
