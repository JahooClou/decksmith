# Changelog

## 1.0.1 (2026-10-16)

- Installed as a plugin; SKILL.md says where the scripts live and that the repo is the source.
- Demo slide 7 points at the GitHub repo instead of a local folder; demo deck and README images re-rendered.

## 1.0.0 (2026-10-16)

First release.

- `motion.py`: 19 effects (entrance, emphasis, exit, motion path), click / with / after
  timing, easing, paragraph builds and cascades, triggers, media playback, chained motion
  paths, transitions including Morph (`!!` names), Slide Zoom, sections, notes, links.
- `master.py`: theme colours and fonts, rebuilt slide master, named layouts, notes theme.
- `images.py` + `image_guide.py`: picture slots, exact px at 1080p and 4K, Pexels crop
  URLs, crop-to-fill placing, and the image guide.
- `validate.py`, `render.py`, `ppt_render.ps1`: structural checks, design-rule warnings,
  and a PowerPoint COM check with effect counts, PNGs, a contact sheet, MP4 and stills.
- Demo deck (7 slides) and a self-test deck that uses every feature.
