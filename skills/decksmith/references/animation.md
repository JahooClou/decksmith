# Animation, transitions, Morph, zoom, triggers, media

All of it is written by `scripts/motion.py`. python-pptx has no API for any of
this; motion.py writes the XML PowerPoint itself writes (references/xml-reference.md).

## The timing model in one picture

PowerPoint's Animation Pane is a tree, and `Motion.add()` builds it for you:

```
main sequence
 ├─ click group 1            "On Click" (or the slide start, if the first effect is With/After)
 │   ├─ time group @0.0 s    effects that start together
 │   │   ├─ effect A         start="click"
 │   │   └─ effect B         start="with", delay=0.3   (its own delay inside the group)
 │   └─ time group @A end    start="after": opens a new time group at the end of the previous one
 │       └─ effect C
 └─ click group 2 ...
trigger sequences (one per trigger shape), media nodes
```

- `start="click"`: a new click group. This is what the presenter's clicker advances.
- `start="with"`: plays with the previous effect. `delay` offsets it.
- `start="after"`: waits for the whole previous time group to finish, plus `delay`.
- If the **first** effect on a slide is `with` or `after`, the group starts when the slide
  starts (PowerPoint's "Start: With Previous / After Previous" on the first effect).
  Use it for a title that floats in by itself.
- A later `with`/`after` effect never opens a click group. It joins the current one.

## Effects

`Motion(slide).add(shape, name, start=, delay=, dur=, direction=, paragraphs=, ease=, amount=, path=, repeat=, rewind=, stagger=)`

| Class | Name | PowerPoint name | Options |
| --- | --- | --- | --- |
| entrance | `appear` | Appear | |
| | `fade` | Fade | |
| | `fly` | Fly In | `direction` top/right/bottom/left (where it comes from) |
| | `float` | Float In (rises 10% of slide height) | |
| | `float-down` | Descend | |
| | `wipe` | Wipe | `direction` |
| | `zoom` | Zoom (fades in, grows from the centre) | |
| emphasis | `grow` | Grow/Shrink | `amount=1.5` (scale) |
| | `spin` | Spin | `amount=360` (degrees) |
| | `transparency` | Transparency | `amount=0.5`, held until the slide ends |
| | `pulse` | Pulse | |
| | `fill-color` | Fill Colour | `amount="E2522B"` |
| exit | `disappear`, `fade-out`, `fly-out`, `float-out`, `wipe-out`, `zoom-out` | the exit twins | `direction` |
| path | `path` | Custom motion path | `path=(dx, dy)` as fractions of slide width/height, or a PowerPoint path string `"M 0 0 C ... E"` |

- `ease="out"` (default) = PowerPoint's *Smooth end* (decel 50%). `"in"`, `"inout"`, or `None` for linear.
  Paths default to smooth start and end, like PowerPoint.
- `paragraphs="each"` builds a text box by paragraph, one click each; `stagger=0.25`
  makes it one click with a cascade instead. `paragraphs=2` or `(1, 3)` targets a range.
- `repeat=2` or `"indefinite"`; `rewind=True` plays it back (autoReverse).
- **Motion paths are relative to where the shape sits on the slide, not to where the
  last path left it.** Found by rendering: a second path snapped the shape back to its
  start. motion.py chains tuple paths on the same shape automatically: each starts at
  the end of the previous one. If you pass a path string, you own that.
- Exit effects leave the shape hidden for the rest of the slide. If the next slide
  Morphs, leave the exited object off it, or it will seem to come back.

## Triggers

```python
m.on_click(button).add(panel, "wipe", direction="left").add(panel, "pulse", start="after")
```
The panel stays hidden until the presenter clicks `button` (not the slide). Use it for
optional answers, a reveal you might skip, or a non-linear menu. Anyone clicking through
with the arrow keys never sees them, so nothing essential goes behind a trigger. Do not
count on a trigger in a video export (the demo's video did not show the revealed text).

## Transitions

`set_transition(slide, kind, dur=0.7, direction=None, option="byObject", advance_after=None, on_click=True)`

| kind | notes |
| --- | --- |
| `morph` | `option` byObject / byWord / byChar. Written with a Fade fallback for old viewers. |
| `fade` | the default between sections |
| `push`, `wipe`, `cover`, `pull` | `direction` left/right/up/down |
| `split`, `zoom`, `cut` | |
| `none` | removes it |

`advance_after=4` makes the slide move on by itself after 4 s (kiosk decks).
The transition belongs to the slide it leads **into**.

## Morph

Morph animates between this slide and the previous one. It pairs objects:
1. objects with the same name that starts with `!!` are always paired, even when
   their type, text or geometry differ (`morph_name(shape, "pop")`);
2. otherwise PowerPoint guesses by type, text and name. Guesses fail on duplicates.

What it animates: position, size, rotation, crop, fill colour, line, text (byWord /
byChar morphs letters), and picture crop. Unpaired objects fade in or out.

Rules that keep Morph clean:
- Name every object that should travel with `!!`. validate.py lists the pairs and warns
  about `!!` names that have no partner on the previous slide.
- One file for a picture that travels. Two different crops of the same photo are two
  pictures to PowerPoint, and they cross-fade instead of moving.
- Keep z-order the same on both slides, or objects swap layers halfway.
- Text morphs best when the words are the same and only size/place change.
- 0.8-1.6 s. Shorter looks like a glitch; longer drags.
- Morph needs PowerPoint 2019 / 365 (Windows, Mac, web, iOS, Android). Keynote, Google
  Slides and LibreOffice show the fallback fade.

## Slide Zoom

`add_slide_zoom(slide, target_slide, image_path, left, top, width, height, return_to_parent=True)`

A thumbnail that zooms into the target slide in the show and, with return to zoom on,
comes back. It needs a picture of the target slide: build once, `render.py`, then build
again with the rendered `slide-NN.png` (build_demo.py `--thumbs` does this). PowerPoint
refreshes the thumbnail when someone edits the zoom. Older viewers get a picture with a
hyperlink to the slide. Section Zoom and Summary Zoom are not written by motion.py;
build those in PowerPoint itself (Insert > Zoom).

## Media

```python
mv = slide.shapes.add_movie("clip.mp4", left, top, w, h, poster_frame_image="poster.png", mime_type="video/mp4")
m = Motion(slide); m.play_media(mv, start="after", loop=True, length=3.0, volume=0.6); m.save()
```
- `start="after"` as the first effect means *Start: Automatically*; `"click"` = on click.
- `length` is the clip length in seconds (`ffprobe`). `loop=True` = *Loop until stopped*.
- python-pptx's own click-to-pause on the video is kept.
- Audio: `add_movie` with an audio file and `mime_type="audio/mpeg"` works the same way;
  hide the icon off-slide or under a shape if it should not show.
- Use H.264 MP4 (AAC audio). PowerPoint on Mac and web plays nothing else reliably.
- Embedded media makes the file large. Say how large when you deliver.

## Sections, notes, links

- `set_sections(prs, [("Opening", [s1]), ("Motion", [s2, s3])])`: every slide in exactly one section.
- `set_notes(slide, text)`: blank lines separate paragraphs. Every slide gets notes.
- `link_to_slide(shape, target)`: click jumps in the show and in a PDF export.
- Action buttons: any shape plus `link_to_slide`, or python-pptx's
  `shape.click_action.action` for first/last/next/previous slide.

## Charts and SmartArt

- Native charts: python-pptx `add_chart`. They inherit the theme's accents; style
  series explicitly. A chart animates as one object (wipe from the left is the classic).
  By-series or by-category chart builds have to be set in PowerPoint.
- SmartArt cannot be created from code. Build the diagram from native shapes (they can
  Morph, SmartArt cannot), or ask the user to convert it in PowerPoint.

## What plays where

| Feature | PowerPoint 365 / 2019+ | PowerPoint 2016 | Keynote / Google Slides / LibreOffice | PDF export | Video export |
| --- | --- | --- | --- | --- | --- |
| Layout, theme, fonts, notes, sections, charts | yes | yes | mostly (fonts must be installed) | yes, notes no | yes |
| Entrance / emphasis / exit / paths | yes | yes | approximations, often wrong timing | final state only | yes, as if clicked |
| Triggers | yes | yes | rarely | no | not played |
| Morph | yes | fallback fade | fallback fade | no | yes |
| Slide Zoom | yes | picture + link | picture + link | picture + link | yes |
| Media autoplay / loop | yes | yes | varies | poster frame | yes |

Never promise more than the first column. Test in PowerPoint (render.py), and tell the
user which column their audience will see.
