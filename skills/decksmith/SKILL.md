---
name: decksmith
description: "Builds advanced, animated PowerPoint decks (.pptx) from code, with real PowerPoint motion and a deliberate design system. Use whenever the user wants a presentation, deck, slides, keynote, pitch deck or PowerPoint, or mentions .pptx, Morph, Morph transition, '!!' naming, slide animations, entrance/emphasis/exit effects, motion paths, animation timing (on click / with previous / after previous), triggers, slide transitions, Slide Zoom, slide masters, layouts, theme colours and fonts, speaker notes, sections, embedded video or audio, or wants an existing deck animated, restyled or checked. Writes the p:timing / transition / Morph / zoom XML that python-pptx cannot, applies taste-skill's design rules translated for slides, plans every picture with an image guide (exact px, topic, a Pexels.com proposition), and verifies the result in real PowerPoint: reopens the file, compares PowerPoint's effect counts with the XML, exports slide PNGs and an MP4 that plays the animation. Prefer it over the generic pptx skill whenever motion, Morph or design quality matter."
---

# decksmith

Animated PowerPoint, built from code and checked in PowerPoint. python-pptx builds
slides but has no animation, transition, Morph, zoom or section API. decksmith writes
that XML directly, copied from files PowerPoint 365 saved itself. It runs every deck
through PowerPoint to prove the result: the file opens, every effect loaded, and the
renders look right. Design follows taste-skill (`Z:\Skills\Old\taste-skill`),
restated for slides in `references/design-rules.md`.

Everything here ran for real on 16 October 2026 (PowerPoint 16.0 Click-to-Run,
python-pptx 1.0.2, Python 3.14). The demo is `Z:\Claude\decksmith-demo\`.

---

## Setup facts

- Python is `py -3.14` (plain `python` is the Windows Store stub). Needs python-pptx,
  lxml, Pillow; PyMuPDF only for the LibreOffice path; ffmpeg for video frames.
- PowerPoint is installed, LibreOffice is not. `render.py` uses PowerPoint through COM
  (PowerShell, no pywin32 needed) and falls back to LibreOffice where PowerPoint is missing.
- **PowerPoint is single-instance and the user often has decks open in it.**
  `ppt_render.ps1` attaches to the running PowerPoint, opens the deck hidden and
  read-only, closes only that deck, and quits PowerPoint only if it was not running
  before. Never call `Quit()` or close presentations you did not open. Never open the
  user's own decks with a window or for writing without being asked.
- Scripts live in `scripts/`. Import them with `sys.path.insert(0, <skill>/scripts)`.

| Script | Job |
| --- | --- |
| `master.py` | theme colours and fonts, slide master, named layouts (`new_deck`) |
| `motion.py` | animation timing, triggers, media playback, transitions, Morph names, Slide Zoom, sections, notes, links |
| `images.py` | picture slots, exact px, Pexels crop URLs and downloads, crop-to-fill placing |
| `image_guide.py` | writes the image guide (markdown + JSON) for a deck |
| `validate.py` | structural errors PowerPoint would repair or drop, plus the design-rule warnings |
| `render.py` + `ppt_render.ps1` | PowerPoint check, slide PNGs, contact sheet, optional MP4 and frames |
| `build_demo.py` | the demo deck, also the worked example to copy from |
| `selftest.py` | builds a deck with every effect and feature, validates and renders it |

---

## The workflow

**1. Read the brief and say the design read.** One line: deck kind, audience, room,
vibe, direction. Then the three dials (VARIANCE / MOTION / DENSITY) with the reason.
Ask one question only if the read truly forks (usually "presented live, or sent as a
file?", because a sent deck gets MOTION 1). See `references/design-rules.md` §1.

**2. Write the outline with motion in it.** Per slide: the one idea, the layout, what
moves, and *why* it moves (hierarchy, sequence, continuity or state). If you cannot
name the reason, the slide is static. Decide which objects Morph carries through a
section and give them `!!` names now.

**3. Set the design system.** `Theme` (one accent; accents 2-6 neutral or tints),
three faces at most, the master's text styles, and the layouts the outline needs
(`references/masters.md`).

**4. Plan every picture as a slot** (`images.add_slot`) with a key, a topic and a Pexels
query. Find a Pexels proposition for each (`references/images.md`): exact px for the
frame at 1080p and 4K, original size large enough, orientation right, colours in the
palette. **Ask before downloading**, naming the photos, the source and the rough size;
then `download_pexels` at the exact 4K crop and `place()`.

**5. Write the build script** in the user's project (copy the shape of
`scripts/build_demo.py`). Its docstring holds the design read and dials. Per slide:
layout, shapes, `Motion(slide)...save()`, `set_transition`, `set_notes`. Then
`set_sections`, `theme_notes` (gives the notes the deck's fonts), save.

**6. Validate.** `py -3.14 validate.py deck.pptx`: 0 errors, and read every warning.

**7. Render in PowerPoint.** `py -3.14 render.py deck.pptx --out <dir>`. It must print
"render check passed": PowerPoint opened the file and loaded exactly the effects,
triggers and Morph transitions the XML declares. Look at `sheet.png`, slide by slide,
against the checklist. The PNGs show every object in its final, built state.

**8. Second pass for Slide Zoom** (if used): rebuild with the rendered `slide-NN.png`
as the zoom thumbnail, then render again.

**9. Watch the motion.** `render.py ... --video --frames 0.25` exports an MP4 in which
clicks play as if pressed, Morph and transitions run, and pulls a still every 0.25 s.
Look at the frames around every Morph and every multi-step build: did the right thing
move, from the right place, in the right order? This is how the motion-path bug below
was found.

**10. Run the image guide** `py -3.14 image_guide.py deck.pptx --picks picks.json --out image-guide.md`
and the pre-delivery checklist (`references/design-rules.md` §8).

**11. Deliver**: the .pptx, the image guide, the contact sheet (and the video if made),
and one honest paragraph on where it plays (below). Say the file size if media is embedded.

---

## Animation in brief (details: `references/animation.md`)

```python
from motion import Motion, set_transition, morph_name, add_slide_zoom, set_sections, set_notes, link_to_slide

m = Motion(slide)
m.add(title, "float", start="after", delay=0.3, dur=0.9)        # first effect after previous = plays on slide start
m.add(body, "fade", start="after", paragraphs="each", stagger=0.3)
m.add(marker, "path", start="click", path=(0, 0.16))           # move down 16% of the slide height
m.add(words, "pulse", start="click")
m.on_click(question).add(answer, "wipe", direction="left")     # trigger
m.save()
set_transition(slide, "morph", dur=1.2)                         # the transition INTO this slide
morph_name(shape, "pop")                                        # name '!!pop' on both slides
```
Effects: appear, fade, fly, float, float-down, wipe, zoom; grow, spin, transparency,
pulse, fill-color; disappear, fade-out, fly-out, float-out, wipe-out, zoom-out; path.
Start modes: click / with / after. Ease-out by default. Transitions: morph, fade, push,
wipe, split, cover, pull, zoom, cut. Media: `m.play_media(movie, start="after", loop=True, length=s)`.

Things learned by rendering, not from docs:
- **Motion paths always start from the shape's place on the slide**, not from where the
  last path ended. `motion.py` chains tuple paths on the same shape for you.
- "Barlow" resolved to its ExtraBold Italic face, because only that weight is installed as
  TrueType here. A family name in XML proves nothing: check the render.
- Setting only `left`/`width` on a placeholder drops it to the top of the slide at zero height.
- python-pptx gives notes a Calibri theme (`theme_notes` fixes it) and never copies
  slide-number placeholders (`show_slide_number`).
- COM opens a file with out-of-order or unknown timing XML without complaint, but fails
  (0x80070570, the UI's repair prompt) on a target id that does not exist. So:
  validate.py for structure, render.py's counts for what PowerPoint actually loaded.

## Being honest about what plays where

Say this in the hand-over, in plain words:
- **Plays fully only in PowerPoint 365 / 2019 and later** (Windows, Mac, web): Morph, Slide
  Zoom, exact timing, triggers, media autoplay and loops.
- PowerPoint 2016 and older, Keynote, Google Slides and LibreOffice show Morph as a fade
  and Slide Zoom as a linked picture, and often get build timing wrong.
- PDF export shows every slide in its final state: no motion, no notes, links still work.
- The PNG renders here are the final built state; the MP4 is the motion. Triggers do not
  fire in the video.
- "Opens without a repair prompt" is checked by a hidden COM open plus validate.py. The
  only full proof is opening the file once in PowerPoint by hand, so say whether you did.
- Not built from code: SmartArt (use native shapes, which also Morph), Section and
  Summary Zoom, by-series chart builds, recorded narration. Point the user to PowerPoint
  for those.

## Design in brief (full rules and checklist: `references/design-rules.md`)

One idea per slide; 45 words at most on a slide, the rest in notes; every slide has notes.
One accent. Three faces at most, no Inter by default, no em or en dashes anywhere.
Left-aligned, asymmetric compositions above VARIANCE 4; no three equal cards; 0.75 in
margins. Real pictures, planned to the pixel, in the image guide. Every effect has a
reason; Morph first, builds only where order matters, one entrance family, ease-out,
0.4-0.9 s builds, 0.8-1.6 s Morph.

---

## References

- `references/design-rules.md`: taste-skill restated for slides, plus the pre-delivery checklist
- `references/animation.md`: timing model, every effect and option, Morph, zoom, triggers, media, the support matrix
- `references/images.md`: slots, exact px, Pexels search, download and licence, the image guide
- `references/masters.md`: theme slots, master, layouts, placeholders, font gotchas
- `references/xml-reference.md`: the XML PowerPoint writes, how to get more from PowerPoint itself (the oracle method), and how PowerPoint reacts to broken XML

For python-pptx basics (tables, charts, text frames, editing an existing deck), the
`anthropic-skills:pptx` skill is the general reference. decksmith adds motion, design
and verification on top of it.
