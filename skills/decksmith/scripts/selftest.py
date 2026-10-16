"""Exercise every effect, transition and helper, then validate and render.

    py -3.14 selftest.py [--out <folder>] [--no-render]

Builds selftest.pptx with one slide per effect class, every transition kind,
a trigger, a chained motion path, a looping auto-playing video, a Slide Zoom,
sections, links and notes. Passes when validate.py finds no errors and, where
PowerPoint is installed, PowerPoint loads exactly the effects the XML declares.
Needs ffmpeg for the test clip (skipped without it).
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from pptx.util import Inches, Pt  # noqa: E402
from master import Theme, layout, new_deck  # noqa: E402
from motion import CATALOG, Motion, add_slide_zoom, link_to_slide, morph_name, set_notes, set_sections, set_transition  # noqa: E402


def box(slide, i, label, cols=6):
    x = 0.6 + (i % cols) * 2.05
    y = 2.0 + (i // cols) * 1.6
    sh = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(1.85), Inches(1.2))
    sh.text_frame.text = label
    sh.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    return sh


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(Path(tempfile.gettempdir()) / "decksmith-selftest"))
    ap.add_argument("--no-render", action="store_true")
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    prs = new_deck(Theme())
    content, blank = layout(prs, "Statement"), layout(prs, "Blank")
    slides = []
    groups = {"entr": [], "emph": [], "exit": [], "path": []}
    for name, (klass, *_r) in CATALOG.items():
        groups[klass].append(name)
    kinds = ["fade", "push", "wipe", "split", "cover", "pull", "zoom", "cut"]
    for gi, (klass, names) in enumerate(groups.items()):
        s = prs.slides.add_slide(content)
        s.shapes.title.text = f"{klass}: {len(names)} effects"
        m = Motion(s)
        for i, n in enumerate(names):
            b = box(s, i, n)
            kw = {}
            if n in ("fly", "wipe", "fly-out", "wipe-out"):
                kw["direction"] = "left"
            if n == "fill-color":
                kw["amount"] = "E2522B"
            if n == "path":
                m.add(b, n, start="click", path=(0.1, 0))
                m.add(b, n, start="after", path=(0, 0.1))   # chained: starts where the first ended
                continue
            m.add(b, n, start="click" if i == 0 else "after", **kw)
        m.save()
        set_transition(s, kinds[gi % len(kinds)], dur=0.6, direction="left" if kinds[gi] in ("push", "wipe") else None)
        set_notes(s, f"Self-test: {', '.join(names)}.")
        slides.append(s)
    # paragraphs, triggers, morph, transitions
    s = prs.slides.add_slide(content)
    s.shapes.title.text = "Paragraph build and a trigger"
    tb = s.shapes.add_textbox(Inches(0.75), Inches(2.2), Inches(6), Inches(3))
    tb.text_frame.text = "One"
    for t in ("Two", "Three"):
        tb.text_frame.add_paragraph().text = t
    trig = box(s, 9, "click me")
    reveal = box(s, 10, "revealed")
    m = Motion(s)
    m.add(tb, "float", start="after", paragraphs="each", stagger=0.2)
    m.on_click(trig).add(reveal, "zoom").add(reveal, "pulse", start="after")
    m.save()
    morph_name(box(s, 5, "morph me"), "probe")
    set_transition(s, "push", dur=0.5, direction="up")
    set_notes(s, "Self-test: paragraph cascade, trigger.")
    slides.append(s)
    s2 = prs.slides.add_slide(content)
    s2.shapes.title.text = "Morph target"
    p = morph_name(box(s2, 0, "morph me"), "probe")
    p.left, p.top = Inches(9), Inches(5)
    set_transition(s2, "morph", dur=1.0, option="byWord", advance_after=4)
    set_notes(s2, "Self-test: morph byWord, auto-advance after 4 s.")
    slides.append(s2)
    # media
    clip = out / "clip.mp4"
    ff = shutil.which("ffmpeg")
    s3 = prs.slides.add_slide(blank)
    if ff:
        if not clip.exists():
            subprocess.run([ff, "-v", "error", "-y", "-f", "lavfi", "-i", "testsrc2=size=640x360:rate=25:duration=3",
                            "-c:v", "libx264", "-pix_fmt", "yuv420p", str(clip)], check=True)
        poster = out / "poster.png"
        if not poster.exists():
            subprocess.run([ff, "-v", "error", "-y", "-i", str(clip), "-frames:v", "1", str(poster)], check=True)
        mv = s3.shapes.add_movie(str(clip), Inches(1), Inches(1), Inches(6.4), Inches(3.6),
                                 poster_frame_image=str(poster), mime_type="video/mp4")
        m = Motion(s3)
        m.play_media(mv, start="after", loop=True, length=3, volume=0.5)
        m.save()
    zoom_img = out / "poster.png" if ff else None
    if zoom_img and zoom_img.exists():
        add_slide_zoom(s3, slides[0], str(zoom_img), Inches(8), Inches(1), Inches(4), Inches(2.25))
    home = box(s3, 11, "back to 1")
    link_to_slide(home, slides[0])
    set_transition(s3, "fade", dur=0.4)
    set_notes(s3, "Self-test: video autoplay and loop, slide zoom, link.")
    slides.append(s3)
    set_sections(prs, [("Effects", slides[:4]), ("Interaction", slides[4:6]), ("Media", slides[6:])])
    pptx = out / "selftest.pptx"
    prs.save(pptx)
    print(pptx)
    py = sys.executable
    r = subprocess.run([py, str(HERE / "validate.py"), str(pptx)])
    if r.returncode:
        sys.exit("validate failed")
    if not a.no_render:
        r = subprocess.run([py, str(HERE / "render.py"), str(pptx), "--out", str(out / "render")])
        if r.returncode:
            sys.exit("render check failed")
    print("selftest passed")


if __name__ == "__main__":
    main()
