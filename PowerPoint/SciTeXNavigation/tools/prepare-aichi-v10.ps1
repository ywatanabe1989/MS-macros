param(
    [Parameter(Mandatory = $true)][string]$SourceDeck,
    [Parameter(Mandatory = $true)][string]$CleanDeck,
    [Parameter(Mandatory = $true)][string]$MacroDeck,
    [Parameter(Mandatory = $true)][string]$ModulePath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

function ConvertTo-Rgb([int]$Red, [int]$Green, [int]$Blue) {
    return $Red + (256 * $Green) + (65536 * $Blue)
}

function Assert-Equal($Actual, $Expected, [string]$Label) {
    if ($Actual -ne $Expected) {
        throw "$Label expected '$Expected' but got '$Actual'"
    }
}

function Assert-True([bool]$Value, [string]$Label) {
    if (-not $Value) {
        throw "$Label was false"
    }
}

function Get-NamedShape($Slide, [string]$Name) {
    foreach ($shape in $Slide.Shapes) {
        if ($shape.Name -ieq $Name) {
            return $shape
        }
    }
    return $null
}

function Get-TitleShape($Slide) {
    $title = Get-NamedShape $Slide "SCITEX_TITLE"
    if ($null -ne $title) {
        return $title
    }

    $candidates = @()
    foreach ($shape in $Slide.Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1 -and $shape.Top -lt 45) {
            $candidates += $shape
        }
    }
    $title = $candidates | Sort-Object Top, Left | Select-Object -First 1
    if ($null -eq $title) {
        throw "Slide $($Slide.SlideIndex) has no deterministic title candidate."
    }
    $title.Name = "SCITEX_TITLE"
    return $title
}

function Remove-AllSlideShapes($Slide) {
    for ($shapeIndex = $Slide.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
        $Slide.Shapes.Item($shapeIndex).Delete()
    }
}

function Add-Rectangle($Slide, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [int]$Color) {
    $shape = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.ForeColor.RGB = $Color
    $shape.Line.Visible = 0
    return $shape
}

function Add-TextBox($Slide, [string]$Name, [string]$Text, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [double]$FontSize, [int]$Color, [bool]$Bold = $false) {
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.TextFrame.MarginLeft = 5
    $shape.TextFrame.MarginRight = 5
    $shape.TextFrame.MarginTop = 3
    $shape.TextFrame.MarginBottom = 3
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = "Arial"
    $shape.TextFrame.TextRange.Font.NameFarEast = "Yu Gothic"
    $shape.TextFrame.TextRange.Font.Size = $FontSize
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    return $shape
}

function Add-ConfigValue($Slide, [string]$Name, [string]$Text, [double]$Left, [double]$Top, [double]$Width, [double]$Height) {
    $ink = ConvertTo-Rgb 27 38 53
    $accent = ConvertTo-Rgb 25 157 179
    $white = ConvertTo-Rgb 255 255 255
    $shape = Add-TextBox $Slide $Name $Text $Left $Top $Width $Height 17 $ink $true
    $shape.Fill.Visible = -1
    $shape.Fill.ForeColor.RGB = $white
    $shape.Line.Visible = -1
    $shape.Line.ForeColor.RGB = $accent
    $shape.Line.Weight = 1.25
    return $shape
}

function Invoke-PowerPointMacro($Application, [string]$MacroName) {
    $method = $Application.GetType().GetMethod("Run", [type[]]@([string], [object[]].MakeByRefType()))
    [object[]]$macroParameters = @()
    $invokeArguments = [object[]]@($MacroName, $macroParameters)
    [void]$method.Invoke($Application, $invokeArguments)
}

function Restore-AccessVbom([string]$Path, [bool]$Existed, $OriginalValue) {
    if ($Existed) {
        New-ItemProperty -Path $Path -Name "AccessVBOM" -PropertyType DWord -Value $OriginalValue -Force | Out-Null
    }
    else {
        Remove-ItemProperty -Path $Path -Name "AccessVBOM" -ErrorAction SilentlyContinue
    }
}

function Set-StaticTocBody($Presentation, $TocSlide, $Body, $Mappings, [bool]$LeftColumn, [int]$CurrentSection) {
    $entries = @()
    foreach ($mapping in $Mappings) {
        $target = $Presentation.Slides.Item($mapping.Slide)
        $section = [int]([regex]::Match($mapping.Code, "^\d+").Value)
        $includeColumn = $(if ($LeftColumn) { $section -le 3 } else { $section -gt 3 })
        if ($includeColumn -and $target.SlideShowTransition.Hidden -eq 0) {
            $entries += [pscustomobject]@{ Slide = $target; Code = $mapping.Code; Section = $section; Title = $target.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text }
        }
    }

    $Body.TextFrame.TextRange.Text = ($entries.Title -join "`r")
    $Body.TextFrame.TextRange.Font.Name = "Arial"
    $Body.TextFrame.TextRange.Font.NameFarEast = "Yu Gothic"
    $Body.TextFrame.TextRange.Font.Size = 18
    $Body.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 2
    $Body.TextFrame.Ruler.Levels.Item(1).FirstMargin = 0
    $Body.TextFrame.Ruler.Levels.Item(1).LeftMargin = 0
    $Body.TextFrame.Ruler.Levels.Item(2).FirstMargin = 18
    $Body.TextFrame.Ruler.Levels.Item(2).LeftMargin = 18

    for ($entryIndex = 1; $entryIndex -le $entries.Count; $entryIndex++) {
        $entry = $entries[$entryIndex - 1]
        $paragraph = $Body.TextFrame.TextRange.Paragraphs($entryIndex, 1)
        $linkLength = $paragraph.Text.Length
        while ($linkLength -gt 0 -and ($paragraph.Text.Substring($linkLength - 1, 1) -eq "`r" -or $paragraph.Text.Substring($linkLength - 1, 1) -eq "`n")) {
            $linkLength--
        }
        $linkRange = $paragraph.Characters(1, $linkLength)
        $linkRange.ActionSettings.Item(1).Action = 7
        $linkRange.ActionSettings.Item(1).Hyperlink.SubAddress = "$($entry.Slide.SlideID),$($entry.Slide.SlideIndex),$($entry.Title)"
        $linkRange.Font.Color.RGB = $(if ($entry.Section -eq $CurrentSection) { ConvertTo-Rgb 27 38 53 } else { ConvertTo-Rgb 170 179 188 })
        if ($entry.Code -match "^\d+$") {
            $paragraph.IndentLevel = 1
            $linkRange.Font.Bold = -1
        }
        else {
            $paragraph.IndentLevel = 2
        }
    }
}

$mappings = @(
    [pscustomobject]@{ Slide = 2; Code = "1" },
    [pscustomobject]@{ Slide = 3; Code = "1a" },
    [pscustomobject]@{ Slide = 4; Code = "2" },
    [pscustomobject]@{ Slide = 5; Code = "2a" },
    [pscustomobject]@{ Slide = 6; Code = "2b" },
    [pscustomobject]@{ Slide = 7; Code = "2c" },
    [pscustomobject]@{ Slide = 9; Code = "3" },
    [pscustomobject]@{ Slide = 10; Code = "3a" },
    [pscustomobject]@{ Slide = 11; Code = "3b" },
    [pscustomobject]@{ Slide = 12; Code = "3c" },
    [pscustomobject]@{ Slide = 13; Code = "3d" },
    [pscustomobject]@{ Slide = 14; Code = "3e" },
    [pscustomobject]@{ Slide = 15; Code = "3f" },
    [pscustomobject]@{ Slide = 16; Code = "3g" },
    [pscustomobject]@{ Slide = 17; Code = "3h" },
    [pscustomobject]@{ Slide = 18; Code = "3i" },
    [pscustomobject]@{ Slide = 20; Code = "4" },
    [pscustomobject]@{ Slide = 21; Code = "4a" },
    [pscustomobject]@{ Slide = 22; Code = "4b" },
    [pscustomobject]@{ Slide = 23; Code = "4c" },
    [pscustomobject]@{ Slide = 24; Code = "4d" },
    [pscustomobject]@{ Slide = 25; Code = "4e" },
    [pscustomobject]@{ Slide = 26; Code = "4f" },
    [pscustomobject]@{ Slide = 28; Code = "5" }
)

$powerPoint = $null
$presentation = $null
$existingPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
if ($null -ne $existingPowerPoint) {
    throw "PowerPoint is already running. Close it before preparing the AICHI deck."
}
$securityPath = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
$securityItem = Get-ItemProperty -Path $securityPath -ErrorAction SilentlyContinue
$accessVbomExisted = ($null -ne $securityItem -and $securityItem.PSObject.Properties.Name -contains "AccessVBOM")
$originalAccessVbom = $(if ($accessVbomExisted) { $securityItem.AccessVBOM } else { $null })

try {
    if (-not (Test-Path -LiteralPath $securityPath)) {
        [void](New-Item -Path $securityPath -Force)
    }
    New-ItemProperty -Path $securityPath -Name "AccessVBOM" -PropertyType DWord -Value 1 -Force | Out-Null
    if (Test-Path -LiteralPath $CleanDeck) { Remove-Item -LiteralPath $CleanDeck -Force }
    if (Test-Path -LiteralPath $MacroDeck) { Remove-Item -LiteralPath $MacroDeck -Force }

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.WindowState = 2
    $powerPoint.DisplayAlerts = 1
    $presentation = $powerPoint.Presentations.Open($SourceDeck, $true, $false, $true)

    Assert-Equal $presentation.Slides.Count 29 "source slide count"
    Assert-Equal $presentation.Designs.Count 1 "source design count"
    Assert-Equal $presentation.SlideMaster.CustomLayouts.Count 1 "source layout count"
    Assert-Equal $presentation.PageSetup.SlideWidth 720 "source slide width"
    Assert-Equal $presentation.PageSetup.SlideHeight 540 "source slide height"
    $sentinelBodySize = $presentation.Slides.Item(25).Shapes.Item("Box133").TextFrame.TextRange.Font.Size

    $presentation.Slides.Item(1).Tags.Add("SCITEX_COVER", "1")
    foreach ($mapping in $mappings) {
        $slide = $presentation.Slides.Item($mapping.Slide)
        $title = Get-TitleShape $slide
        $baseTitle = [regex]::Replace($title.TextFrame.TextRange.Text.Trim(), "^\d+[A-Za-z]*\.\s*", "")
        $title.TextFrame.TextRange.Text = "$($mapping.Code). $baseTitle"
        $slide.Tags.Add("SCITEX_NAV_CODE", $mapping.Code)
    }

    $tocSpecifications = @(
        [pscustomobject]@{ Slide = 8; Current = 3 },
        [pscustomobject]@{ Slide = 19; Current = 4 },
        [pscustomobject]@{ Slide = 27; Current = 5 }
    )
    foreach ($tocSpec in $tocSpecifications) {
        $tocSlide = $presentation.Slides.Item($tocSpec.Slide)
        $tocTitle = Get-TitleShape $tocSlide
        for ($shapeIndex = $tocSlide.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
            $shape = $tocSlide.Shapes.Item($shapeIndex)
            if ($shape.Name -ne $tocTitle.Name -and $shape.HasTextFrame -eq -1) {
                $shape.Delete()
            }
        }
        $leftBody = Add-TextBox $tocSlide "SCITEX_TOC_BODY_LEFT" "" 32.4 56.2 363.6 405 18 (ConvertTo-Rgb 27 38 53) $false
        $rightBody = Add-TextBox $tocSlide "SCITEX_TOC_BODY_RIGHT" "" 406.8 56.2 291.6 300 18 (ConvertTo-Rgb 27 38 53) $false
        $tocSlide.Tags.Add("SCITEX_TOC", "1")
        $tocSlide.Tags.Add("SCITEX_CURRENT_SECTION", [string]$tocSpec.Current)
        $tocSlide.Tags.Add("SCITEX_TOC_SPLIT_AFTER", "3")
        Set-StaticTocBody $presentation $tocSlide $leftBody $mappings $true $tocSpec.Current
        Set-StaticTocBody $presentation $tocSlide $rightBody $mappings $false $tocSpec.Current
    }

    $config = $presentation.Slides.Item(29)
    Remove-AllSlideShapes $config
    $navy = ConvertTo-Rgb 21 40 66
    $ink = ConvertTo-Rgb 27 38 53
    $muted = ConvertTo-Rgb 93 107 124
    $accent = ConvertTo-Rgb 25 157 179
    $white = ConvertTo-Rgb 255 255 255
    $panel = ConvertTo-Rgb 241 246 249
    [void](Add-Rectangle $config "SCITEX_CONFIG_BG" 0 0 720 540 $white)
    [void](Add-Rectangle $config "SCITEX_CONFIG_HEADER" 0 0 720 70 $navy)
    [void](Add-Rectangle $config "SCITEX_CONFIG_ACCENT" 0 70 720 5 $accent)
    [void](Add-TextBox $config "SCITEX_CONFIG_TITLE" "SciTeX Navigation - Configuration" 38 13 640 42 24 $white $true)
    [void](Add-TextBox $config "SCITEX_CONFIG_HELP" "Edit the value boxes, then run RunSciTeXNavigation again." 45 86 630 30 14 $muted $false)
    [void](Add-Rectangle $config "SCITEX_CONFIG_TYPOGRAPHY_PANEL" 42 125 636 180 $panel)
    [void](Add-TextBox $config "SCITEX_CONFIG_TYPOGRAPHY_TITLE" "Navigation typography" 60 139 260 30 19 $ink $true)
    [void](Add-TextBox $config "SCITEX_CONFIG_LATIN_LABEL" "Latin font" 60 184 115 30 14 $muted $false)
    [void](Add-ConfigValue $config "SCITEX_CFG_FONT_LATIN" "Arial" 175 177 170 38)
    [void](Add-TextBox $config "SCITEX_CONFIG_CJK_LABEL" "CJK font" 60 234 115 30 14 $muted $false)
    [void](Add-ConfigValue $config "SCITEX_CFG_FONT_CJK" "Yu Gothic" 175 227 170 38)
    [void](Add-TextBox $config "SCITEX_CONFIG_MIN_LABEL" "Minimum size (pt)" 370 184 170 30 14 $muted $false)
    [void](Add-ConfigValue $config "SCITEX_CFG_FONT_MIN" "18" 585 177 60 38)
    [void](Add-TextBox $config "SCITEX_CONFIG_MAX_LABEL" "Maximum size (pt)" 370 234 170 30 14 $muted $false)
    [void](Add-ConfigValue $config "SCITEX_CFG_FONT_MAX" "32" 585 227 60 38)
    [void](Add-Rectangle $config "SCITEX_CONFIG_TOC_PANEL" 42 325 636 110 $panel)
    [void](Add-TextBox $config "SCITEX_CONFIG_TOC_TITLE" "Table of contents" 60 339 220 30 19 $ink $true)
    [void](Add-TextBox $config "SCITEX_CONFIG_HIDE_LABEL" "Hide hidden slides from TOC" 60 385 265 30 14 $muted $false)
    [void](Add-ConfigValue $config "SCITEX_CFG_HIDE_HIDDEN" "Yes" 310 378 65 38)
    [void](Add-TextBox $config "SCITEX_CONFIG_VERSION_LABEL" "Version" 425 385 75 30 14 $muted $false)
    [void](Add-ConfigValue $config "SCITEX_CFG_VERSION" "0.1.1" 500 378 110 38)
    [void](Add-TextBox $config "SCITEX_CONFIG_FOOTER" "Accepted values: Yes / No. This slide is always excluded from the TOC and slide show." 45 458 630 42 13 $muted $false)
    $config.Tags.Add("SCITEX_CONFIG", "1")
    $config.Tags.Add("SCITEX_ALWAYS_SKIP", "1")
    $config.SlideShowTransition.Hidden = -1

    $presentation.SaveAs($CleanDeck, 24)
    $presentation.SaveAs($MacroDeck, 25)
    [void]$presentation.VBProject.VBComponents.Import($ModulePath)
    $presentation.Save()
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"

    Assert-Equal $presentation.Slides.Count 29 "macro deck slide count"
    Assert-Equal $presentation.Designs.Count 1 "macro deck design count"
    Assert-Equal $presentation.SlideMaster.CustomLayouts.Count 1 "macro deck layout count"
    Assert-Equal $presentation.Slides.Item(25).Shapes.Item("Box133").TextFrame.TextRange.Font.Size $sentinelBodySize "non-navigation body typography preserved"
    Assert-Equal $presentation.VBProject.VBComponents.Count 1 "macro deck VBA component count"
    Assert-Equal $presentation.VBProject.VBComponents.Item(1).Name "SciTeXNavigation" "macro deck VBA module name"

    foreach ($mapping in $mappings) {
        $titleText = $presentation.Slides.Item($mapping.Slide).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text
        Assert-True ($titleText.StartsWith("$($mapping.Code). ")) "slide $($mapping.Slide) explicit navigation code"
    }

    $leftExpected = @()
    $rightExpected = @()
    foreach ($mapping in $mappings) {
        $target = $presentation.Slides.Item($mapping.Slide)
        if ($target.SlideShowTransition.Hidden -eq 0) {
            $section = [int]([regex]::Match($mapping.Code, "^\d+").Value)
            if ($section -le 3) { $leftExpected += $target.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text }
            else { $rightExpected += $target.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text }
        }
    }
    foreach ($tocSpec in $tocSpecifications) {
        $tocSlide = $presentation.Slides.Item($tocSpec.Slide)
        Assert-Equal $tocSlide.Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text ($leftExpected -join "`r") "TOC slide $($tocSpec.Slide) left full contents"
        Assert-Equal $tocSlide.Shapes.Item("SCITEX_TOC_BODY_RIGHT").TextFrame.TextRange.Text ($rightExpected -join "`r") "TOC slide $($tocSpec.Slide) right full contents"
        Assert-Equal $tocSlide.Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Paragraphs(2, 1).IndentLevel 2 "TOC slide $($tocSpec.Slide) child indentation"
    }
    Assert-True (($leftExpected -join "|") -notmatch "3h\.") "hidden slide excluded from TOC"
    Assert-Equal $presentation.Slides.Item(29).Shapes.Item("SCITEX_CFG_VERSION").TextFrame.TextRange.Text "0.1.1" "configuration version"
    $configText = ""
    foreach ($shape in $presentation.Slides.Item(29).Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) { $configText += $shape.TextFrame.TextRange.Text }
    }
    Assert-True ($configText -notmatch "[^\x00-\x7F]") "configuration slide is English ASCII only"
    $presentation.Save()

    [ordered]@{
        source = $SourceDeck
        clean_pptx = $CleanDeck
        macro_pptm = $MacroDeck
        slides = $presentation.Slides.Count
        navigation_entries = $mappings.Count
        toc_slides = $tocSpecifications.Count
        explicit_numbering = "passed"
        full_two_column_toc = "passed"
        hierarchical_indentation = "passed"
        hidden_slide_filter = "passed"
        english_configuration = "passed"
        version = "0.1.1"
        non_navigation_content_preserved = "passed"
        master_and_layout_counts_preserved = "passed"
    } | ConvertTo-Json
}
catch {
    Write-Error ("AICHI_PREPARE_FAILED: " + $_.Exception.Message + "`n" + $_.ScriptStackTrace)
    throw
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) } catch { }
        $presentation = $null
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) } catch { }
        $powerPoint = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Restore-AccessVbom $securityPath $accessVbomExisted $originalAccessVbom
    $restoreScript = Join-Path $PSScriptRoot "restore-access-vbom.ps1"
    $originalExistedNumber = $(if ($accessVbomExisted) { 1 } else { 0 })
    $originalValueNumber = $(if ($accessVbomExisted) { [int]$originalAccessVbom } else { 0 })
    $restoreArguments = @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $restoreScript, "-SecurityPath", $securityPath, "-OriginalExisted", $originalExistedNumber, "-OriginalValue", $originalValueNumber)
    Start-Process -FilePath "powershell.exe" -ArgumentList $restoreArguments -WindowStyle Hidden | Out-Null
}
