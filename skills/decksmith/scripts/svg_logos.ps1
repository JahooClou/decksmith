# svg_logos.ps1: swap pictures for SVG logos through PowerPoint (logos stay whole and vector: PowerPoint keeps the
# SVG with its own PNG fallback). Each SVG takes the old picture's box, name, stacking order and, on slides, its
# animation (effect, timing, trigger). Works on slides and on layouts. Trim the SVG first (svgshapes.trim_svg) so
# the box is the artwork's box. PowerPoint is single-instance: quits only if it was not running before.
#   powershell -File svg_logos.ps1 -Deck deck.pptx -Cfg logos.json
# logos.json: { "jobs": [ { "scope": "slide", "slide": 1, "shape": "logo", "svg": "logo_trim.svg" },
#                         { "scope": "layout", "layout": "Title", "shape": "Picture 3", "svg": "logo_trim.svg" } ] }
param([string]$Deck, [string]$Cfg)
$c = Get-Content -Raw -Encoding UTF8 $Cfg | ConvertFrom-Json
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)   # never quit a PowerPoint the user had open
$pp = New-Object -ComObject PowerPoint.Application
function Swap($shapes, $seq, $name, $svg) {
    $old = $shapes.Item($name)
    $z = $old.ZOrderPosition
    $new = $shapes.AddPicture((Resolve-Path $svg).Path, 0, -1, $old.Left, $old.Top, $old.Width, $old.Height)
    while ($new.ZOrderPosition -gt $z) { $new.ZOrder(3) }
    $idx = 0
    if ($seq) { for ($i = 1; $i -le $seq.Count; $i++) { if ($seq.Item($i).Shape.Name -eq $name) { $idx = $i; break } } }
    if ($idx -gt 0) { $o = $seq.Item($idx); $fx = $o.EffectType; $dur = $o.Timing.Duration; $dly = $o.Timing.TriggerDelayTime; $trg = $o.Timing.TriggerType }
    $old.Delete()
    $new.Name = $name
    if ($idx -gt 0) { $e = $seq.AddEffect($new, $fx, 0, $trg, $idx); $e.Timing.Duration = $dur; $e.Timing.TriggerDelayTime = $dly }
    return $new.Type
}
try {
    $p = $pp.Presentations.Open((Resolve-Path $Deck).Path, $false, $false, $false)
    foreach ($j in $c.jobs) {
        if ($j.scope -eq 'slide') {
            $sl = $p.Slides([int]$j.slide)
            $t = Swap $sl.Shapes $sl.TimeLine.MainSequence $j.shape $j.svg
            "slide $($j.slide) $($j.shape): type $t"
        } else {
            $hit = 0
            foreach ($d in $p.Designs) { foreach ($cl in $d.SlideMaster.CustomLayouts) {
                if ($cl.Name -eq $j.layout) { $t = Swap $cl.Shapes $null $j.shape $j.svg; $hit++ } } }
            if ($hit -ne 1) { throw "layout $($j.layout) matched $hit times" }
            "layout $($j.layout): type $t"
        }
    }
    $p.Save(); $p.Close()
} finally {
    if (-not $wasRunning -and $pp.Presentations.Count -eq 0) { $pp.Quit() }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pp) | Out-Null
}
