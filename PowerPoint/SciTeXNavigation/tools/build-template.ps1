<#
.SYNOPSIS
Build the distributable SciTeX Navigation template (.pptm).

.DESCRIPTION
build-and-test.ps1 builds a SANDBOX to exercise the macro, and it uses the
single-column SCITEX_TOC_BODY -- so it neither tests nor demonstrates the
columned index. This builds the thing we hand to someone else: a deck they can
open, read, press one button in, and then copy their own slides into.

What it contains, in the order a new reader meets it:

  1  cover
  2  what this is, and the one thing to press
  3  the four shape names the macro looks for
  4  how the index columns work
  5  an index page (two columns, empty until the macro runs)
  6  two content slides, so the index has something to list
  7  a second index page, to show the current section highlighted
  8  one more content slide
  9  the configuration page

The usage is IN the file rather than in a README beside it, because a .pptm
travels alone -- it gets mailed, dropped in Teams, copied to a stick. Anything
kept outside it is lost by the second hop.

.PARAMETER ModulePath
The exported SciTeXNavigation.bas to embed.

.PARAMETER Output
Where to write the template. Must be .pptm.

.NOTES
From WSL, launch through cmd.exe -- calling powershell.exe directly across the
interop socket fails intermittently and prints only a vsock error.
#>
param(
    [Parameter(Mandatory = $true)][string]$ModulePath,
    [Parameter(Mandatory = $true)][string]$Output
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $ModulePath)) { throw "Not found: $ModulePath" }
if (-not $Output.EndsWith(".pptm")) { throw "Output must be .pptm" }

$INK    = 27 + (38 * 256) + (53 * 65536)
$MUTED  = 120 + (128 * 256) + (136 * 65536)
$ACCENT = 25 + (157 * 256) + (179 * 65536)
$WHITE  = 255 + (255 * 256) + (255 * 65536)

function Add-Box($Slide, $Name, $Text, $L, $T, $W, $H, $Size, $Color, $Bold) {
    $shape = $Slide.Shapes.AddTextbox(1, $L, $T, $W, $H)
    $shape.Name = $Name
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.MarginLeft = 6
    $shape.TextFrame.MarginRight = 6
    $shape.TextFrame.MarginTop = 3
    $shape.TextFrame.MarginBottom = 3
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Size = $Size
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    return $shape
}

function Add-Page($Presentation, $Title) {
    $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, 12)
    [void](Add-Box $slide "SCITEX_TITLE" $Title 32 24 660 44 28 $INK $true)
    return $slide
}

$app = $null; $pres = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $pres = $app.Presentations.Add(-1)
    $pres.PageSetup.SlideWidth = 720
    $pres.PageSetup.SlideHeight = 540

    # 1 --- cover. Tagged so the macro leaves it out of the index.
    $cover = $pres.Slides.Add(1, 12)
    $cover.Tags.Add("SCITEX_COVER", "1") | Out-Null
    [void](Add-Box $cover "COVER_TITLE" "SciTeX Navigation" 60 190 600 60 40 $INK $true)
    [void](Add-Box $cover "COVER_SUB" "A table of contents that keeps itself correct." 60 256 600 36 20 $MUTED $false)

    # 2 --- what it is
    $what = Add-Page $pres "What this does"
    [void](Add-Box $what "BODY" @"
Every index page in this deck is rebuilt from the slides themselves:
numbering, titles, links, and the split across columns.

To run it: Alt+F8 -> RunSciTeXNavigation -> Run.

A copy of the file is saved beside it first, named
  <name>.before-navigation-<date>.pptm
because PowerPoint's undo does not survive a macro.
"@ 32 96 640 320 18 $INK $false)

    # 3 --- the names it looks for
    $names = Add-Page $pres "The names it looks for"
    [void](Add-Box $names "BODY" @"
SCITEX_TITLE            the slide's title (every slide)
SCITEX_TOC_BODY_*       one index column; any suffix, ordered left to right
SCITEX_STATUS           where the macro reports what it did
SCITEX_CFG_*            the settings on the last page

Slide tags mark the roles:
SCITEX_COVER on the cover, SCITEX_TOC on each index page.

Rename a shape in the Selection Pane (Home -> Select -> Selection Pane).
"@ 32 96 640 320 16 $INK $false)

    # 4 --- how columns behave
    $cols = Add-Page $pres "How the columns fill"
    [void](Add-Box $cols "BODY" @"
Entries fill the first column from the top, then carry on into the next.
The split is measured, not written down, so it moves when slides do.

Two columns or three -- add another SCITEX_TOC_BODY_* box and the macro
uses it. More columns means the type can stay larger: 54 entries need
12pt across two columns and fit at 15pt across three.

If nothing fits even at the smallest configured size, the macro says so
on the status line rather than letting entries fall off the page.
"@ 32 96 640 320 16 $INK $false)

    # 5 --- first index page, two columns
    $toc1 = Add-Page $pres "Contents"
    $toc1.Tags.Add("SCITEX_TOC", "1") | Out-Null
    [void](Add-Box $toc1 "SCITEX_TOC_BODY_LEFT"  "Run the macro to fill this." 32 96 310 400 18 $MUTED $false)
    [void](Add-Box $toc1 "SCITEX_TOC_BODY_RIGHT" "" 378 96 310 400 18 $MUTED $false)

    # 6 --- content
    $c1 = Add-Page $pres "A content slide"
    [void](Add-Box $c1 "BODY" "Ordinary slides need no tags. The macro numbers them from the index page above." 32 96 640 100 18 $MUTED $false)
    $c2 = Add-Page $pres "Another content slide"
    [void](Add-Box $c2 "BODY" "Add, delete or reorder slides, run the macro again, and the index follows." 32 96 640 100 18 $MUTED $false)

    # 7 --- second index page
    $toc2 = Add-Page $pres "Contents"
    $toc2.Tags.Add("SCITEX_TOC", "1") | Out-Null
    [void](Add-Box $toc2 "SCITEX_TOC_BODY_LEFT"  "Run the macro to fill this." 32 96 310 400 18 $MUTED $false)
    [void](Add-Box $toc2 "SCITEX_TOC_BODY_RIGHT" "" 378 96 310 400 18 $MUTED $false)

    $c3 = Add-Page $pres "A slide in the second section"
    [void](Add-Box $c3 "BODY" "Each index page highlights the section it introduces and greys the rest." 32 96 640 100 18 $MUTED $false)

    # 8 --- configuration, last page
    $cfg = Add-Page $pres "Configuration"
    $cfg.Tags.Add("SCITEX_CONFIG", "1") | Out-Null
    [void](Add-Box $cfg "CFG_HELP" "Edit the boxes on the right. The macro reads them on every run." 32 92 640 28 14 $MUTED $false)
    $labels = @(
        @("Latin font",      "SCITEX_CFG_FONT_LATIN", "Segoe UI"),
        @("Japanese font",   "SCITEX_CFG_FONT_CJK",   "Yu Gothic"),
        @("Smallest type",   "SCITEX_CFG_FONT_MIN",   "12"),
        @("Largest type",    "SCITEX_CFG_FONT_MAX",   "32"),
        @("Skip hidden",     "SCITEX_CFG_HIDE_HIDDEN","Yes"),
        @("Macro version",   "SCITEX_CFG_VERSION",    "0.4.0")
    )
    $y = 132
    foreach ($row in $labels) {
        [void](Add-Box $cfg ("LBL_" + $row[1]) $row[0] 32 $y 200 30 16 $MUTED $false)
        $v = Add-Box $cfg $row[1] $row[2] 240 $y 220 30 16 $INK $true
        $v.Line.Visible = -1
        $v.Line.ForeColor.RGB = $ACCENT
        $y += 40
    }
    [void](Add-Box $cfg "SCITEX_STATUS" "The macro writes what it did here." 32 470 640 28 12 $MUTED $false)

    $pres.VBProject.VBComponents.Import($ModulePath) | Out-Null
    Write-Output ("embedded " + (Split-Path -Leaf $ModulePath))

    $pres.SaveAs($Output, 25)
    Write-Output ("wrote " + $Output + " (" + $pres.Slides.Count + " slides)")
}
catch { Write-Output ("ERROR: " + $_.Exception.Message) }
finally {
    if ($null -ne $pres) { try { $pres.Saved = $true; $pres.Close() } catch { } }
}
