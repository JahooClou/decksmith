# XML reference

What motion.py writes, copied from files PowerPoint 365 (16.0) saved itself.
Namespaces: `p` presentationml, `a` drawingml, `r` relationships,
`mc` markup-compatibility, `p14` `http://schemas.microsoft.com/office/powerpoint/2010/main`,
`p159` `http://schemas.microsoft.com/office/powerpoint/2015/09/main`,
`pslz` `http://schemas.microsoft.com/office/powerpoint/2016/slidezoom`,
`p166` `http://schemas.microsoft.com/office/powerpoint/2016/6/main`.

## Getting the truth from PowerPoint (the oracle method)

When an effect or option is not in motion.py, do not guess the XML. Make PowerPoint
write it, then copy it. In PowerShell (PowerPoint installed):

```powershell
$app = New-Object -ComObject PowerPoint.Application
$p = $app.Presentations.Add(0)                      # no window
$s = $p.Slides.Add(1, 12)                           # blank layout
$sh = $s.Shapes.AddShape(1, 50, 50, 200, 100)
$e = $s.TimeLine.MainSequence.AddEffect($sh, 39, 0, 1)   # 39 = msoAnimEffectAscend (Float In), 1 = on click
$e.Timing.Duration = 0.8; $e.Timing.SmoothEnd = -1
$p.SaveAs("$env:TEMP\oracle.pptx", 24); $p.Close()   # never $app.Quit() if PowerPoint was already open
```
Unzip and read `ppt/slides/slide1.xml`. Enum values: load
`C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll`
with `Add-Type -Path` and list `[Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]`.
Morph is `EntryEffect = 3954` (by object), 3955 (by word), 3956 (by character); the
2013 interop does not name them.

## Slide child order

`p:sld` = `p:cSld`, `p:clrMapOvr`, transition (often wrapped in `mc:AlternateContent`),
`p:timing`, `p:extLst`. Out of order: COM still opens it, but it is invalid; validate.py fails it.

## Timing skeleton

```xml
<p:timing><p:tnLst><p:par>
 <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
  <p:seq concurrent="1" nextAc="seek">
   <p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
    <p:par><p:cTn id="3" fill="hold">                         <!-- click group -->
      <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>   <!-- + <p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond> to auto-start -->
      <p:childTnLst>
       <p:par><p:cTn id="4" fill="hold">                      <!-- time group -->
         <p:stCondLst><p:cond delay="0"/></p:stCondLst>         <!-- start offset within the click (ms) -->
         <p:childTnLst>
          <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" decel="50000"
                        fill="hold" grpId="0" nodeType="clickEffect">   <!-- withEffect / afterEffect -->
            <p:stCondLst><p:cond delay="0"/></p:stCondLst>      <!-- the effect's own delay -->
            <p:childTnLst> behaviours </p:childTnLst></p:cTn></p:par>
       </p:childTnLst></p:cTn></p:par>
    </p:childTnLst></p:cTn></p:par>
   </p:childTnLst></p:cTn>
   <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
   <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
  </p:seq>
 </p:childTnLst></p:cTn></p:par></p:tnLst>
 <p:bldLst><p:bldP spid="4" grpId="0" animBg="1"/></p:bldLst>
</p:timing>
```
- `cTn/@id` unique per slide. `spTgt/@spid` = the shape's `cNvPr/@id`.
- `grpId` counts the effects on one shape (0, 1, 2 ...); a by-paragraph build shares one.
- `bldLst`: `bldP` for shapes with text bodies (`build="p"` for paragraph builds,
  `animBg="1"` when the shape has a fill), `bldGraphic` + `bldAsOne` for charts and
  tables, nothing for pictures.
- Paragraph target: `<p:spTgt spid="4"><p:txEl><p:pRg st="1" end="1"/></p:txEl></p:spTgt>`.
- `repeatCount="2000"` = 2 times (thousandths). Smooth end = `decel="50000"`.

## Behaviour templates (PowerPoint's own)

| Effect | presetID / class / subtype | Behaviours |
| --- | --- | --- |
| Appear | 1 entr 0 | `set style.visibility=visible` |
| Fade | 10 entr 0 | set visible + `animEffect transition="in" filter="fade"` |
| Fly In | 2 entr 1/2/4/8 (top/right/bottom/left) | set + `anim ppt_x`, `anim ppt_y` (`additive="base"`), from `0-#ppt_h/2`, `1+#ppt_w/2`, `1+#ppt_h/2`, `0-#ppt_w/2` |
| Float In | 42 entr 0 | set + fade + `anim ppt_x #ppt_x→#ppt_x` + `anim ppt_y #ppt_y+.1→#ppt_y` |
| Wipe | 22 entr 1/2/4/8 | set + `animEffect filter="wipe(up/right/down/left)"` |
| Zoom | 53 entr 16 | set + `anim ppt_w 0→#ppt_w`, `anim ppt_h 0→#ppt_h` + fade |
| Grow/Shrink | 6 emph 0 | `animScale` `<p:by x="150000" y="150000"/>` |
| Spin | 8 emph 0 | `animRot by="21600000"` (60000 per degree), attr `r` |
| Transparency | 9 emph 0 | `set style.opacity` + `animEffect filter="image" prLst="opacity: 0.5"`, `dur="indefinite"` |
| Pulse | 26 emph 0 | fade with `tmFilter="0, 0; .2, .5; .8, .5; 1, 0"` + `animScale by 105%` `autoRev="1"` |
| Fill Colour | 1 emph 2 | `animClr clrSpc="rgb" dir="cw"` attr `fillcolor` + `set fill.type=solid`, `set fill.on=true` (no grpId) |
| exits | same IDs, class `exit` | the entrance behaviours reversed, values without `#` (`ppt_y` to `1+ppt_h/2`), then `set style.visibility=hidden` at `dur-1` ms |
| Motion path | 0 path 0 | `animMotion origin="layout" path="M 0 0 L 0.25 0 E" pathEditMode="relative"`, attrs `ppt_x ppt_y`; coordinates are fractions of the slide, relative to the layout position |
| Media play | 1 mediacall 0 | `<p:cmd type="call" cmd="playFrom(0.0)">`, plus a `p:video/p:cMediaNode` with `repeatCount="indefinite"` for loop and `vol="50000"` |

Trigger sequence: a second `p:seq` whose `cTn` has `nodeType="interactiveSeq"`,
`restart="whenNotActive"`, `evtFilter="cancelBubble"`, starts on
`<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="TRIGGER"/></p:tgtEl></p:cond>`,
an `endSync`, and the same condition again in `nextCondLst`.

## Transitions

```xml
<mc:AlternateContent xmlns:mc="...markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1600" advTm="3000">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback><p:transition spd="slow" advTm="3000"><p:fade/></p:transition></mc:Fallback>
</mc:AlternateContent>
```
Other transitions use `Requires="p14"` so `p14:dur` can carry the exact duration, with
the same transition in the fallback. `advClick="0"` disables advance on click.

## Slide Zoom

```xml
<mc:AlternateContent><mc:Choice Requires="pslz">
 <p:graphicFrame> nvGraphicFramePr (cNvPr id, graphicFrameLocks noGrp noChangeAspect) + xfrm
  <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom">
   <pslz:sldZm><pslz:sldZmObj sldId="257">
     <pslz:zmPr id="{GUID}" returnToParent="1" transitionDur="1000">
       <p166:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>
       <p166:spPr> xfrm + prstGeom rect + a:ln </p166:spPr>
     </pslz:zmPr></pslz:sldZmObj></pslz:sldZm>
  </a:graphicData></a:graphic></p:graphicFrame>
</mc:Choice><mc:Fallback>
 <p:pic> cNvPr with <a:hlinkClick r:id="rIdSlide" action="ppaction://hlinksldjump"/>, same blip and geometry </p:pic>
</mc:Fallback></mc:AlternateContent>
```
Schema: MS-PPTX 2.8.3.1 `CT_ZoomObjectProperties`
(learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/059e3722-139d-4e41-9841-d53eecaf73f6)
and `CT_SlideZoomObject` (…/c1683604-6df4-4859-9e9d-dac99077d340). `sldId` is the
target's id in `p:sldIdLst`, not its position. Choice and Fallback share one shape id.

## Sections

In `ppt/presentation.xml`, `p:extLst`:
```xml
<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
 <p14:sectionLst><p14:section name="Intro" id="{GUID}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst>
</p:ext>
```

## Morph names

`p:cNvPr/@name` starting with `!!`. Alt text lives in `@descr`, which Morph ignores;
images.py keeps picture metadata in `@title`, so the name stays free.

## What PowerPoint does with bad XML (tested through COM)

| Damage | `Presentations.Open` via COM |
| --- | --- |
| animation targets a missing shape id | fails, 0x80070570 ("file or directory is corrupted"): the UI's repair prompt |
| slide children out of order | opens, all effects loaded |
| unknown element inside `p:timing` | opens, all effects loaded |
| missing `bldP` entry | opens, all effects loaded |
| duplicate `cTn` ids | opens, all effects loaded |

So a clean COM open is necessary, not sufficient. validate.py catches the rest, and
render.py compares PowerPoint's effect counts with the XML so a silently dropped timing
tree cannot pass.
