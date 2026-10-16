<#
decksmith: open a .pptx in the real PowerPoint (hidden), report what PowerPoint
parsed, export slide PNGs and optionally an MP4 that plays the animations.

  pwsh -File ppt_render.ps1 -Path deck.pptx -OutDir renders [-Video] [-Width 1920]

Safety: PowerPoint is single-instance, so this attaches to a PowerPoint the user
may already have open. It opens the deck read-only without a window, closes only
that deck, and quits PowerPoint only if it was not running before.
Output: renders\report.json plus slide-NN.png (and deck.mp4 with -Video).
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [switch]$Video,
  [int]$Width = 1920,
  [int]$VideoHeight = 720,
  [int]$Fps = 25,
  [int]$SlideSeconds = 3
)
$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path $Path).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
$report = [ordered]@{ path = $Path; opened = $false; error = $null; powerpoint = $app.Version; slides = @() }
$pres = $null
try {
  try {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $app.Presentations.Open($Path, -1, 0, 0)
    $report.opened = $true
  } catch {
    # 0x80070570 here is what the UI shows as "PowerPoint found a problem with content"
    $report.error = $_.Exception.Message
    throw
  }
  $report.slideWidth = $pres.PageSetup.SlideWidth
  $report.slideHeight = $pres.PageSetup.SlideHeight
  $report.sections = @()
  for ($i = 1; $i -le $pres.SectionProperties.Count; $i++) {
    $report.sections += [ordered]@{ name = $pres.SectionProperties.Name($i); first = $pres.SectionProperties.FirstSlide($i); count = $pres.SectionProperties.SlidesCount($i) }
  }
  $report.fonts = @()
  foreach ($f in $pres.Fonts) { $report.fonts += [ordered]@{ name = $f.Name; embedded = [bool]$f.Embedded; embeddable = [bool]$f.Embeddable } }
  $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  foreach ($s in $pres.Slides) {
    $t = $s.SlideShowTransition
    $effects = @()
    foreach ($e in $s.TimeLine.MainSequence) {
      $effects += [ordered]@{
        shape = $e.Shape.Name; type = [int]$e.EffectType; exit = [bool]$e.Exit
        trigger = [int]$e.Timing.TriggerType; delay = $e.Timing.TriggerDelayTime; duration = $e.Timing.Duration
        paragraph = $e.Paragraph
      }
    }
    $triggers = @()
    foreach ($seq in $s.TimeLine.InteractiveSequences) {
      $n = $seq.Count
      if ($n -gt 0) { $triggers += [ordered]@{ trigger = $seq.Item(1).Timing.TriggerShape.Name; effects = $n } }
    }
    $shapes = @()
    foreach ($sh in $s.Shapes) {
      $shapes += [ordered]@{ name = $sh.Name; type = [int]$sh.Type; left = $sh.Left; top = $sh.Top; width = $sh.Width; height = $sh.Height
        hasText = [bool]($sh.HasTextFrame -and $sh.TextFrame.HasText) }
    }
    $notes = ''
    try { $notes = $s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text } catch {}
    $png = Join-Path $OutDir ('slide-{0:D2}.png' -f $s.SlideIndex)
    $s.Export($png, 'PNG', $Width, $h)
    $report.slides += [ordered]@{
      index = $s.SlideIndex; id = $s.SlideID; layout = $s.CustomLayout.Name
      transition = [int]$t.EntryEffect; transitionDuration = $t.Duration
      advanceOnTime = [bool]$t.AdvanceOnTime; advanceTime = $t.AdvanceTime
      mainEffects = $effects; interactive = $triggers; shapes = $shapes; notesChars = $notes.Length; png = $png
    }
  }
  if ($Video) {
    $mp4 = Join-Path $OutDir 'deck.mp4'
    if (Test-Path $mp4) { Remove-Item $mp4 -Force }
    # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality)
    $pres.CreateVideo($mp4, $true, $SlideSeconds, $VideoHeight, $Fps, 85)
    $deadline = (Get-Date).AddMinutes(15)
    while ($pres.CreateVideoStatus -eq 1 -or $pres.CreateVideoStatus -eq 2) {
      if ((Get-Date) -gt $deadline) { throw 'video export timed out' }
      Start-Sleep -Milliseconds 700
    }
    $report.video = [ordered]@{ path = $mp4; status = [int]$pres.CreateVideoStatus }  # 3 = done, 4 = failed
  }
} catch {
  if (-not $report.error) { $report.error = $_.Exception.Message }
} finally {
  if ($pres) { $pres.Close() }
  if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
  $json = $report | ConvertTo-Json -Depth 8
  [IO.File]::WriteAllText((Join-Path $OutDir 'report.json'), $json, [Text.UTF8Encoding]::new($false))
  if ($report.error) { Write-Output ('ERROR: ' + $report.error); exit 2 }
  Write-Output ('OK ' + $report.slides.Count + ' slides')
}
