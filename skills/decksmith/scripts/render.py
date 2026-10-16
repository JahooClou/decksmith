"""Render a deck and check that the renderer read what the XML says.

    py -3.14 render.py deck.pptx --out renders [--video] [--frames 0.5] [--engine auto|powerpoint|libreoffice]

PowerPoint engine (Windows with PowerPoint installed):
  - opens the file hidden and read-only; a COM error on open is the automated
    face of the "PowerPoint found a problem with content" repair prompt
  - compares PowerPoint's own count of animation effects, triggers and the
    transition type per slide with what the XML declares; if PowerPoint
    silently dropped a broken timing tree, the counts differ and this fails
  - exports slide PNGs, and with --video an MP4 in which click animations,
    Morph and transitions actually play; --frames N pulls a still every N s
LibreOffice engine (anywhere soffice exists): PDF -> PNG. It shows layout,
theme, fonts and notes only. It does not play Morph, timing or zoom.

Writes <out>/sheet.png (contact sheet) and <out>/check.json. Exit code 1 when a
check fails.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from motion import NS, q  # noqa: E402

MORPH = 3954  # PowerPoint's EntryEffect value for Morph by object (3955 word, 3956 char)


def expected(pptx: Path):
    """Per slide, from the XML: main-sequence effects, trigger sequences, transition kind."""
    z = zipfile.ZipFile(pptx)
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    out = []
    for sid in pres.find(q("p:sldIdLst")):
        path = "ppt/" + target[sid.get(q("r:id"))].lstrip("/").replace("ppt/", "")
        sld = etree.fromstring(z.read(path))
        main = trig = 0
        for seq in sld.iter(q("p:seq")):
            c = seq.find(q("p:cTn"))
            if c is None:
                continue
            n = sum(1 for x in c.iter(q("p:cTn")) if x.get("presetClass"))
            if c.get("nodeType") == "mainSeq":
                main += n
            elif c.get("nodeType") == "interactiveSeq":
                trig += 1 if n else 0
        kind = None
        for node in sld:
            tr = node if node.tag == q("p:transition") else node.find(".//" + q("p:transition"))
            if tr is not None and len(tr):
                kind = etree.QName(tr[0]).localname
                break
        out.append(dict(path=path, main=main, triggers=trig, transition=kind))
    return out


def find_soffice():
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def has_powerpoint():
    if os.name != "nt":
        return False
    for p in (r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
              r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"):
        if os.path.exists(p):
            return True
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application")
        return True
    except OSError:
        return False


def run_powerpoint(pptx, out, video, timeout):
    shell = shutil.which("pwsh") or shutil.which("powershell")
    cmd = [shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "ppt_render.ps1"),
           "-Path", str(pptx), "-OutDir", str(out)]
    if video:
        cmd.append("-Video")
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return None, ("PowerPoint did not answer within %ss. A dialog (often the repair prompt) "
                      "may be waiting in PowerPoint; look at its window." % timeout)
    rep = out / "report.json"
    data = json.loads(rep.read_text(encoding="utf-8-sig")) if rep.exists() else None
    if r.returncode != 0:
        return data, (r.stdout + r.stderr).strip()
    return data, None


def run_libreoffice(pptx, out, timeout):
    soffice = find_soffice()
    if not soffice:
        return None, "LibreOffice not found"
    r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(pptx)],
                       capture_output=True, text=True, timeout=timeout)
    pdf = out / (pptx.stem + ".pdf")
    if not pdf.exists():
        return None, r.stdout + r.stderr
    import fitz  # PyMuPDF
    doc = fitz.open(pdf)
    slides = []
    for i, page in enumerate(doc, 1):
        png = out / f"slide-{i:02d}.png"
        page.get_pixmap(dpi=144).save(png)
        slides.append(dict(index=i, png=str(png)))
    return dict(opened=True, slides=slides, engine="libreoffice"), None


def contact_sheet(pngs, dest, cols=3, width=640):
    from PIL import Image, ImageDraw
    ims = [Image.open(p).convert("RGB") for p in pngs]
    if not ims:
        return None
    h = int(width * ims[0].height / ims[0].width)
    rows = (len(ims) + cols - 1) // cols
    pad = 16
    sheet = Image.new("RGB", (cols * (width + pad) + pad, rows * (h + pad + 22) + pad), (60, 60, 64))
    d = ImageDraw.Draw(sheet)
    for i, im in enumerate(ims):
        x = pad + (i % cols) * (width + pad)
        y = pad + (i // cols) * (h + pad + 22)
        sheet.paste(im.resize((width, h), Image.LANCZOS), (x, y + 22))
        d.text((x, y + 4), f"{i + 1}", fill=(230, 230, 230))
    sheet.save(dest)
    return dest


def frames(mp4, out, every):
    ff = shutil.which("ffmpeg")
    if not ff or not Path(mp4).exists():
        return []
    fdir = out / "frames"
    fdir.mkdir(exist_ok=True)
    for f in fdir.glob("*.png"):
        f.unlink()
    subprocess.run([ff, "-v", "error", "-y", "-i", str(mp4), "-vf", f"fps=1/{every}",
                    str(fdir / "f-%03d.png")], check=False)
    return sorted(fdir.glob("f-*.png"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None)
    ap.add_argument("--engine", default="auto", choices=["auto", "powerpoint", "libreoffice"])
    ap.add_argument("--video", action="store_true", help="PowerPoint only: export an MP4 that plays the animation")
    ap.add_argument("--frames", type=float, default=0, help="with --video: one still every N seconds")
    ap.add_argument("--timeout", type=int, default=900)
    a = ap.parse_args()
    pptx = Path(a.pptx).resolve()
    out = Path(a.out or pptx.with_name(pptx.stem + "-render")).resolve()
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("slide-*.png"):
        f.unlink()
    exp = expected(pptx)
    engine = a.engine
    if engine == "auto":
        engine = "powerpoint" if has_powerpoint() else "libreoffice"
    problems = []
    if engine == "powerpoint":
        data, err = run_powerpoint(pptx, out, a.video, a.timeout)
    else:
        data, err = run_libreoffice(pptx, out, a.timeout)
    if err:
        problems.append(f"{engine}: {err}")
    result = dict(engine=engine, pptx=str(pptx), slides=[], problems=problems)
    if data and data.get("slides"):
        for e, s in zip(exp, data["slides"]):
            row = dict(index=s["index"], png=s.get("png"), expected=e)
            if engine == "powerpoint":
                got_main = len(s.get("mainEffects") or [])
                got_trig = len(s.get("interactive") or [])
                row.update(powerpoint_main=got_main, powerpoint_triggers=got_trig,
                           powerpoint_transition=s.get("transition"), layout=s.get("layout"),
                           notes_chars=s.get("notesChars"))
                if got_main != e["main"]:
                    problems.append(f"slide {s['index']}: XML has {e['main']} main-sequence effects, "
                                    f"PowerPoint loaded {got_main} (it drops timing it cannot read)")
                if got_trig != e["triggers"]:
                    problems.append(f"slide {s['index']}: XML has {e['triggers']} trigger sequences, PowerPoint loaded {got_trig}")
                if e["transition"] == "morph" and s.get("transition") not in (3954, 3955, 3956):
                    problems.append(f"slide {s['index']}: Morph in XML, PowerPoint reports transition {s.get('transition')}")
            result["slides"].append(row)
        if engine == "powerpoint":
            result["sections"] = data.get("sections")
            result["fonts"] = data.get("fonts")
            result["video"] = data.get("video")
        if len(data["slides"]) != len(exp):
            problems.append(f"renderer saw {len(data['slides'])} slides, XML lists {len(exp)}")
    pngs = sorted(out.glob("slide-*.png"))
    sheet = contact_sheet(pngs, out / "sheet.png")
    result["sheet"] = str(sheet) if sheet else None
    if a.video and a.frames and result.get("video"):
        result["frames"] = [str(p) for p in frames(result["video"]["path"], out, a.frames)]
    (out / "check.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(f"engine: {engine}   slides: {len(pngs)}   sheet: {sheet}")
    if engine == "powerpoint":
        for r in result["slides"]:
            print(f"  slide {r['index']:>2}  {str(r.get('layout')):<10} effects {r['expected']['main']:>2}/{r.get('powerpoint_main')!s:<3}"
                  f" triggers {r['expected']['triggers']}/{r.get('powerpoint_triggers')}"
                  f"  transition {r['expected']['transition'] or '-'}/{r.get('powerpoint_transition')}"
                  f"  notes {r.get('notes_chars')}")
        if result.get("sections"):
            print("  sections:", ", ".join(f"{s['name']}({s['count']})" for s in result["sections"]))
        if result.get("fonts"):
            print("  fonts:", ", ".join(f"{f['name']}{'' if f['embeddable'] else ' (not embeddable)'}" for f in result["fonts"]))
    if result.get("video"):
        print("  video:", result["video"])
    if result.get("frames"):
        print(f"  frames: {len(result['frames'])} in {out / 'frames'}")
    if problems:
        print("PROBLEMS:")
        for p in problems:
            print("  -", p)
        sys.exit(1)
    print("render check passed")


if __name__ == "__main__":
    main()
