# Backlog Krew house style theme

`BacklogKrewHouseStyle.json` is the Power BI report theme for the approved page designs in
`docs/design/<report>/*.png`. Every colour and size in it was read from those images. The images
are the source of truth: if the theme and an image disagree, the image wins and the theme is
wrong.

Applying the theme to a report is a design-pass step. This folder only holds the file.

## Colours

The values were sampled from all 22 approved pages (2560 x 1440 PNG, a 1280 x 720 canvas at 2x).
Every value below appears on every page unless the "Seen on" column says otherwise.

| Role | Hex | Where it is on the pages | Seen on |
| -- | -- | -- | -- |
| Navy | `#16324A` | Header band, footer band, table header row, navy bars | all |
| Teal | `#138A8A` | Rule under the header band, Premium box border, positive numbers and bars | all |
| Red | `#C62828` | ONE FIX box border and label, "Fix now" values, risk bars | all |
| Amber | `#D98A00` | Warning values, report-only and medium-risk series | 11 of 13 reports |
| Page background | `#EEF1F4` | Canvas behind the panels and tiles | all |
| Visual background | `#FFFFFF` | Panels, tiles, table rows, ONE FIX box | all |
| Visual border | `#D5DCE2` | 1 px border around panels and the tile row, square corners | all |
| Main text | `#1F2A33` | Headline sentence, chart titles, tile numbers, table text | all |
| Axis and body text | `#3D4C58` | Axis labels, legend text, ONE FIX body, Premium body | all |
| Secondary text | `#5B6B78` | Tile labels, subtitles under titles, notes | all |
| Table row divider | `#EEF1F4` | 1 px line between table rows | all |
| Premium box fill | `#F2FAFA` | Fill inside the dashed teal Premium box | 11 of 13 reports |
| Premium heading | `#0E6B6B` | Heading text in the Premium box | 11 of 13 reports |
| Grey series | `#A9B7C2` | "Blocked" series (identity posture, legacy sign-ins) | 1 report |
| Header subtitle | `#B8C7D3` | Subtitle and run line on the navy header band | all |
| Light teal | `#9ED3D3` | Heat map fill (Entra activity, odd hours) | 1 report |
| Green | `#2E7D32` | Positive figures (license utilization and one other report) | 2 reports |

Measured from `teams-activity/page-1-where-conversations-live.png`: the ONE FIX border is 2 px
`#C62828` on white, and the Premium box is a 1 px dashed `#138A8A` border on `#F2FAFA`.

## Sizes

Measured on the 2x images and converted to points (1 canvas px = 0.75 pt). They are
approximate, within about 1 pt.

| Text | Size | Weight |
| -- | -- | -- |
| Tile number (callout) | 17 pt | Semibold |
| Header band report name | 14 pt | Semibold |
| Chart and panel title, headline sentence | 10 pt | Semibold |
| Tile label, table body, legend | 9 pt | Regular |
| Axis labels, table header | 8 pt | Semibold for the header |

The pages are drawn in a sans-serif that does not ship with Power BI. The theme uses Segoe UI
and Segoe UI Semibold, which do ship with it.

## What the theme sets

- `dataColors`: teal, navy, red, amber, grey, dark teal, green, light teal. Teal and navy come
  first, as on the pages. The last two are light, so use them for fills, not thin lines.
- `good` teal, `neutral` amber, `bad` red. Gradient minimum `#F2FAFA`, centre `#9ED3D3`,
  maximum `#138A8A`, as in the heat map.
- Structural colours (`foreground`, `background`, `secondaryBackground`, `tableAccent`, the
  four `*LevelElements`) from the table above.
- `textClasses`: `callout`, `largeTitle`, `title`, `header` and `label` at the sizes above.
- `visualStyles`:
  - `page`: background and outspace `#EEF1F4`.
  - every visual: white background, 1 px `#D5DCE2` border with radius 0, no shadow, title
    text, axis and legend text, no gridlines (the pages show none).
  - `card`: title on by default, in `#5B6B78` at 9 pt, above a 17 pt value. Category labels
    off. BRO-481 found the legacy card shows no title unless it is turned on, and that category
    labels did not render reliably in Desktop.
  - `cardVisual`: label on, in `#5B6B78` at 9 pt, value at 17 pt.
  - `tableEx`, `pivotTable`: navy header with white 8 pt text, white rows with no banding,
    `#EEF1F4` horizontal dividers, no vertical lines.
  - `slicer`: header text.
  - `textbox`: no background or border by default, so a headline sits on its panel. Two style
    presets: `One fix` (white, 2 px red border) and `Premium` (`#F2FAFA`, 1 px teal border).

## What the theme cannot set

- **The Premium box's dashed border.** The report theme schema has dash settings only for
  chart lines and gridlines, not for a visual's border. The `Premium` preset gives the fill and
  a solid teal border. The design pass has to draw the dashes on the page.
- **The header band, footer band, teal rule and tile icon circles.** These are shapes and
  images on the page, not theme defaults.
- **The pink risk rows and the red, amber and teal numbers in tables.** These are conditional
  formatting, which a theme cannot hold ([Learn](https://learn.microsoft.com/power-bi/create-reports/report-themes-create-custom#considerations-and-limitations)).

## Schema

`$schema` points at `reportThemeSchema-2.157.json`, the newest file in
[microsoft/powerbi-desktop-samples](https://github.com/microsoft/powerbi-desktop-samples/tree/main/Report%20Theme%20JSON%20Schema)
on 2026-10-13. The theme validates against it with 0 errors. The schema file is about 1.2 MB, so
it is not committed.

## Relation to the skill's `base.json`

`.claude/skills/powerbi-report-cli/references/design/assets/base.json` is the skill's generic
starting theme, not the house style. Where it disagrees with the approved pages (grey table
headers, 8 px rounded `#E8E8E8` borders, its colour palette), this theme follows the pages. Two of
its safeguards are kept because the pages do not contradict them: a textbox has no background
or border, and table columns grow to fit.

## How a report applies it

- **Power BI Desktop:** on the **Design** ribbon, select **Import theme** and choose
  `BacklogKrewHouseStyle.json`. Importing a theme removes existing customizations from the
  report. See [Use report themes in Power BI](https://learn.microsoft.com/power-bi/create-reports/desktop-report-themes#apply-a-custom-report-theme-json-file).
- **PBIR (in the repo):** copy the file into the report's
  `StaticResources/RegisteredResources/` folder and register it in `definition/report.json`,
  under `themeCollection.customTheme` and in `resourcePackages`. The skill's
  `.claude/skills/powerbi-report-cli/references/authoring/theming.md` gives both entries and the
  file-naming rule. Then run `powerbi-report-author validate` on the report. See
  [Create custom report themes in Power BI Desktop](https://learn.microsoft.com/power-bi/create-reports/report-themes-create-custom).
- **Style presets:** after import, a textbox's **Style** dropdown in the Format pane lists
  `One fix` and `Premium`.
