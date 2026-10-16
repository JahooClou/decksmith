# Masters, layouts, theme, placeholders

`scripts/master.py` turns the base package (`assets/base-16x9.pptx`, an empty 16:9
deck saved by PowerPoint 365) into a deck with its own design system.

```python
from master import Theme, TextStyle, Layout, PH, new_deck, layout, show_slide_number, DEFAULT_LAYOUTS

theme = Theme(name="Client Q4", dk1="121214", lt1="F2F2F0", dk2="4A4A4F", lt2="E2E2DE",
              accent1="1F4BD8", major="Clash Display", minor="Work Sans", mono="IBM Plex Mono")
prs = new_deck(theme)                         # theme + master + DEFAULT_LAYOUTS
s = prs.slides.add_slide(layout(prs, "Content"))
show_slide_number(s)                           # python-pptx does not copy footer placeholders
```

## What each part controls

| Part | Holds | Written by |
| --- | --- | --- |
| Theme (`ppt/theme/theme1.xml`) | 12 colour slots, major (heading) and minor (body) fonts | `apply_theme` |
| Slide master | ground, title/body/other text styles for 9 levels, the master placeholders, fixed decoration | `style_master` |
| Layouts | named slide types with placeholder geometry and per-layout overrides | `build_layouts` |
| Slides | content; placeholders inherit from layout, layout from master | python-pptx |

Colour slots: `dk1` text, `lt1` ground, `dk2` secondary text, `lt2` hairlines and quiet
fills, `accent1` the accent, `accent2-6` chart series (neutral or accent tints),
`hlink`/`folHlink`. In XML, text uses `tx1/tx2/bg1/bg2` which the master's colour map
points at `dk1/dk2/lt1/lt2`.

Fonts: `+mj-lt` = theme major, `+mn-lt` = theme minor. Use those in styles so a font
change is one line. The mono face is not a theme slot; set it per run.

## Default layouts

| Name | Type | Placeholders |
| --- | --- | --- |
| Cover | title | ctrTitle bottom-left 60 pt, subTitle under it |
| Statement | titleOnly | title at top, the rest is free composition |
| Content | obj | title + body (7.2 in wide, so a picture or chart fits beside it) |
| Section | secHead | big title + one line |
| Blank | blank | nothing but date/footer/number |

Make your own with `Layout(name, type, [PH(kind, x, y, w, h, idx=, anchor=, style=TextStyle(...), prompt=)])`.
Placeholder kinds: `title`, `ctrTitle`, `subTitle`, `body`, `pic`, `chart`, `dt`,
`ftr`, `sldNum`. Body placeholders on one layout need distinct `idx` (1, 2, ...).
Layout `type` matters: PowerPoint uses it to re-map content when someone changes a
slide's layout.

## Master decoration

`style_master(prs, theme, decorate=fn)` calls `fn(spTree)`; append fixed shapes there
(a logo, a rule). Every layout shows them. Keep them out of Morph's way: master
shapes do not Morph, and they sit under every slide's content.

## Gotchas

- python-pptx copies no footer, date or slide-number placeholders to new slides:
  call `show_slide_number(slide)` on slides that should show one.
- Setting only `left`/`width` on a placeholder writes a zero-height frame at y=0.
  Set all four (`left, top, width, height`) or none.
- Fonts must exist where the deck is shown, or be embedded (File > Options > Save >
  Embed fonts). Whether a font can be embedded depends on its licence flags;
  render.py's report lists each font's `embeddable` flag as PowerPoint sees it. If a
  font is not embeddable, say so in the hand-over.
- PowerPoint's font list (render.py "fonts:") still names Calibri when the deck has a
  native chart: it comes from the chart's embedded Excel workbook, not from any slide.
  Nothing on screen uses it.
- A family name can map to the wrong face (here "Barlow" = ExtraBold Italic because only
  that weight is TrueType). Always look at the render.
- Base package: regenerate `assets/base-16x9.pptx` from PowerPoint if it is ever lost:
  `$p = $app.Presentations.Add(0); $p.SaveAs("base-16x9.pptx", 24); $p.Close()`.
  Its layout names come out in PowerPoint's UI language; `build_layouts` renames them.
- Polish UI language names layouts "Tytuł i zawartość" and the like; never look layouts up
  by the Office names, only by the names you gave them.
