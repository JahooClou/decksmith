# Design rules for decks

These are the taste-skill rules restated for slides. The source is
`Z:\Skills\Old\taste-skill\skills\`. Read the file named in brackets when a
rule needs its full reasoning; nothing here replaces it.

- `taste-skill/SKILL.md`: brief inference, the three dials, colour, type, layout, AI tells, pre-flight
- `redesign-skill/SKILL.md`: audit order for improving an existing deck
- `minimalist-skill/SKILL.md`: the quiet editorial direction (monochrome, hairlines, flat)
- `soft-skill/SKILL.md`: the expensive soft direction (whitespace, diffused depth, slow easing)
- `brutalist-skill/SKILL.md`: the Swiss or terminal direction (one red, sharp grid, mono)
- `output-skill/SKILL.md`: no placeholders, finish every slide

The web-only parts of taste-skill do not apply to decks: React, Tailwind,
`100dvh`, dark-mode toggles, GSAP, Core Web Vitals, hover states, forms.

## 1. Read the brief first [taste-skill §0, §1]

Before you build anything, write one line:

> Reading this as: <deck kind> for <audience>, in <room / screen>, with a <vibe> language, leaning toward <direction>.

Deck kinds: pitch, keynote talk, teaching deck, internal update, sales leave-behind,
event loop (kiosk), portfolio. The room matters as much as the audience. A projector
in a lit room washes out low contrast. A 75-inch TV lets you use small type. A deck
read on a laptop as a PDF is a document, and it gets no animation.

Then set the dials, and write them into the build script's docstring:

| Dial | 1-3 | 4-7 | 8-10 |
| --- | --- | --- | --- |
| VARIANCE (layout) | one grid, centred titles | offset grid, asymmetric splits | broken grid, bleed, oversized type |
| MOTION | transitions only, no builds | Morph between slides plus builds where order matters | Morph choreography, motion paths, triggers |
| DENSITY | one idea, 0-15 words | 15-45 words, one visual | data slides, max 1 table or chart |

Starting points: keynote 7/6/2, pitch 6/5/3, teaching 5/5/4, internal update 4/3/5,
leave-behind PDF 4/1/6, kiosk loop 6/7/2. If the brief is ambiguous, ask **one**
question, for example "is this presented live or sent as a file?"

## 2. Colour [taste-skill §4.2, §8.B, §9.A]

- **One accent per deck.** It marks the one thing on a slide that matters most.
  If it is on everything, it marks nothing.
- The theme's `accent2`-`accent6` feed chart series and SmartArt. Set them to
  neutrals or tints of the accent (`master.Theme` does this), or a default chart
  drags in Office blue and orange.
- No pure `#000000` or `#FFFFFF` (validate.py warns). Use an off-black and an off-white.
- One grey family per deck. Do not mix warm and cool greys.
- No AI purple-blue glow, no neon, no gradient text, no outer glow.
- For premium-consumer briefs, the beige + brass + espresso palette is banned as a
  default [taste-skill §4.2]. Pick another family and say why.
- **Theme lock.** The deck is light or dark. One deliberate full-bleed photo or
  colour slide is a device. Light and dark alternating at random is a bug.
- Contrast is a floor: body text 4.5:1 against its ground, large type 3:1. Projectors
  lose contrast, so aim higher.
- Text on a photo needs a scrim (a block of the ground colour at 70-85%) or a
  quiet area of the photo. Never rely on luck.

## 3. Type [taste-skill §4.1, §9.B]

- Three faces at most: display, text, mono. validate.py warns at four.
- Inter is not the default (use it only on request). Fraunces and Instrument Serif
  are banned as default serifs. Serif only with a reason you can state.
- Emphasis inside a heading uses the same family in italic or bold, never a
  different family.
- Sizes for a 13.33 x 7.5 in slide:

  | Role | Size | Notes |
  | --- | --- | --- |
  | Cover title | 54-66 pt | 2 lines max, tracking -1 to -1.5 pt, line 0.9-0.95 |
  | Slide title | 36-44 pt | one line if possible, two at most |
  | Statement | 40-54 pt | the whole slide is the sentence |
  | Body | 20-24 pt | 18 pt is the floor for anything that must be read in the room |
  | Labels, axis, captions | 13-16 pt | |
  | Footer, slide number | 10-12 pt | the only place under 12 pt |

- Paragraph spacing comes from `space_before`, not empty paragraphs.
- Numbers, code and timings go in the mono face. That is the only place the mono belongs.
- **Polish copy:** no 1-2 letter word at the end of a line, and no number split from its unit.
  Use a no-break space (U+00A0) after `w z o u i a` and between `5 km`. Check that
  the font has the NBSP glyph (Metropolis does not; see memory "Polish typesetting").
- Check every font in the PowerPoint render, not only in the XML. A family name can
  resolve to the wrong face: on this machine "Barlow" resolves to its ExtraBold
  Italic face, because only that weight is TrueType. render.py's font list shows
  what PowerPoint actually used.

## 4. Layout [taste-skill §4.3, §4.4, §4.7, §9.C]

- **One idea per slide.** If the title needs "and", make two slides.
- Hero discipline applies to the cover: a title, a one-line promise of 20 words
  or fewer, and one visual. No tagline strip, no logo row, no date block on the cover.
- **No centred-everything** when VARIANCE is above 4: left-align and let one side
  carry a picture or a block of colour.
- **No three equal cards.** Use a two-column split, an asymmetric pair, a list that
  builds, or a single large number.
- **Repeat layouts with purpose.** A layout family (split, statement, chart and note)
  may repeat, but never three slides of the same composition in a row unless it is
  a deliberate sequence that Morph animates.
- **Eyebrows:** one small uppercase label per three slides at most. Usually none:
  the title is enough. Never number them ("01 / Intro").
- **No split header:** a big title on the left with a small paragraph floating at
  top right. Stack the title and the text instead.
- **Shape lock:** one corner-radius system per deck (all sharp, or all at one radius).
- **Margins:** 0.75 in at the sides, 0.7 in at the top. Nothing touches the edge
  unless it bleeds on purpose: a full-bleed photo or a colour block.
- Group with space and alignment before boxes. Draw a card only when its elevation means something.
- Bento grids: exactly as many cells as you have items.

## 5. Content [taste-skill §4.9, §9.D, §9.F, §9.G]

- 45 words per slide at most (validate.py warns). The rest goes into speaker notes,
  which every slide needs (validate.py warns).
- A list of more than 5 items is a different slide: group it, chart it, or split it.
- **No em dashes or en dashes anywhere** (validate.py warns). Use a comma, a full
  stop, a colon or parentheses. Ranges use a hyphen.
- No fake-precise numbers. A figure comes from the brief or is labelled as an example.
- No "Elevate / Seamless / Unleash / Next-gen", no "Jane Doe" and no "Acme".
- No "Step 1 / Phase 02" labels. The step's own verb is the label.
- No decorative status dots, no photo-credit captions as decoration, no version
  stamps, no "scroll" cues (decks have none anyway).
- Quotes: 3 lines at most, typographic quotes, attribution with name and role.
- Re-read every visible string before delivery and rewrite anything cute, vague or broken.

## 6. Pictures [taste-skill §4.8]

- A deck is a visual product. Text-only slides are fine for statements; a deck
  with no pictures at all is unfinished.
- Plan every picture as a slot with an exact frame, and keep the image guide
  (references/images.md): exact px, topic, and a Pexels proposition for each.
- No fake screenshots built from rectangles. Use a real screenshot, a real photo, or nothing.
- No labels or pills laid over photos. Put a caption under the photo, or leave it out.
- Icons come from one library with one stroke weight. Never draw icon paths by hand.

## 7. Motion [taste-skill §5 "motion must be motivated", §6.B]

Every animation answers one of four questions, and the speaker notes say which:

1. **Hierarchy:** what should the audience look at first?
2. **Sequence:** in what order should they read it?
3. **Continuity:** where did this come from? (Morph)
4. **State:** what changed? (emphasis, exit)

"It looked nice" is not a reason; drop the effect. The rules:

- **Morph is the main tool.** Two slides with the same named objects tell the eye
  what moved. Use it for zooming into a detail, re-arranging a diagram, or carrying
  an object (a photo, a colour block) through a section.
- **Builds only where order matters.** Steps, a sequence or an argument can build.
  A list of facts appears at once.
- **One entrance family per deck.** Fade and Float In cover nearly everything.
  Bounce, Spin, Boomerang and Swivel are banned unless the brief is playful.
- **Ease out.** Entrances decelerate (decel is the motion.py default). Linear only for
  continuous motion such as a marquee.
- **Durations:** builds 0.4-0.9 s, Morph 0.8-1.6 s, slide transitions 0.4-0.8 s.
  More than 2 s needs a reason. Cascades are staggered by 0.15-0.35 s.
- **Motion claimed, motion shown:** if the read says MOTION 5+, the deck moves.
  If you cannot make it work, set MOTION to 2 and deliver a clean static deck.
- **One transition family.** Morph within a section, Fade between sections. Push,
  Cube, Vortex and friends only on request.
- **Respect the room.** For a deck that is sent as a PDF or read on screen, MOTION
  is 1 and the meaning must survive with every build already visible (the static
  PNG export shows exactly this).
- Auto-advance (kiosk) decks: every slide holds at least 1.5 s after its last
  build finishes.

## 8. Pre-delivery checklist

Run it on every deck. One unticked box means the deck is not done.

**Brief and system**
- [ ] Design read written (one line), dials set and recorded in the build script
- [ ] Theme: one accent; accent2-6 neutral or tints; dk1/lt1 not pure black/white
- [ ] Master and layouts are the deck's own (no Office layout names, no Calibri)
- [ ] Three faces at most, checked in the PowerPoint render's font list

**Every slide**
- [ ] One idea; title reads as a sentence or a clear label
- [ ] 45 words or fewer on the slide; the rest in speaker notes
- [ ] Speaker notes present, saying what to say and why anything moves
- [ ] Body text 18 pt or larger; only footers under 12 pt
- [ ] No em/en dashes; no AI copy; no fake numbers; strings re-read
- [ ] Inside the 0.75 in margins unless bleeding on purpose; nothing overlaps
      (look at the PNG, not the code)
- [ ] Contrast holds: text on photos has a scrim or a quiet area
- [ ] Polish copy: no hanging 1-2 letter words, no number split from its unit

**Pictures**
- [ ] Every picture is in the image guide with exact px, topic and a Pexels proposition
- [ ] Every embedded picture is sharp at 4K ("sharp at 4K" in the guide), or the guide says why not
- [ ] One file per picture across Morph slides; alt text set on every picture

**Motion**
- [ ] Every effect has a reason (hierarchy, sequence, continuity, state)
- [ ] One entrance family; ease-out; durations in range
- [ ] Morph pairs reported by validate.py exist for every Morph slide
- [ ] Chained motion paths start where the last one ended (motion.py does this)
- [ ] render.py: PowerPoint effect counts equal the XML counts on every slide
- [ ] Video frames looked at for every Morph and every multi-step build

**File**
- [ ] validate.py: 0 errors
- [ ] render.py: "render check passed"
- [ ] Opens in PowerPoint with no repair prompt (render.py's COM open is the proxy;
      open it once by hand before it ships)
- [ ] Sections named; slide numbers on content slides; links tested in the show
