param(
    [Parameter(Mandatory = $true)][string]$Deck
)

$ErrorActionPreference = "Stop"

function ConvertTo-Rgb([int]$Red, [int]$Green, [int]$Blue) {
    return $Red + (256 * $Green) + (65536 * $Blue)
}

function Assert-Equal($Actual, $Expected, [string]$Label) {
    if ($Actual -ne $Expected) { throw "$Label expected '$Expected' but got '$Actual'" }
}

function Assert-True([bool]$Value, [string]$Label) {
    if (-not $Value) { throw "$Label was false" }
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

function Get-ExpectedEntries($Presentation, $Mappings, [bool]$LeftColumn, [bool]$IncludeHidden) {
    $entries = @()
    foreach ($mapping in $Mappings) {
        $target = $Presentation.Slides.Item($mapping.Slide)
        $section = [int]([regex]::Match($mapping.Code, "^\d+").Value)
        $includeColumn = $(if ($LeftColumn) { $section -le 3 } else { $section -gt 3 })
        if ($includeColumn -and ($IncludeHidden -or $target.SlideShowTransition.Hidden -eq 0)) {
            $entries += [pscustomobject]@{
                Slide = $target
                Code = $mapping.Code
                Section = $section
                Title = $target.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text
            }
        }
    }
    return $entries
}

function Assert-TocColumn($Body, $Entries, [int]$CurrentSection, [string]$Label) {
    Assert-Equal $Body.TextFrame.TextRange.Text ($Entries.Title -join "`r") "$Label complete text"
    for ($entryIndex = 1; $entryIndex -le $Entries.Count; $entryIndex++) {
        $entry = $Entries[$entryIndex - 1]
        $paragraph = $Body.TextFrame.TextRange.Paragraphs($entryIndex, 1)
        $character = $paragraph.Characters(1, 1)
        $action = $character.ActionSettings.Item(1)
        Assert-Equal $action.Action 7 "$Label line $entryIndex hyperlink action"
        Assert-True ($action.Hyperlink.SubAddress -match ",$($entry.Slide.SlideIndex),") "$Label line $entryIndex hyperlink target"
        $expectedIndent = $(if ($entry.Code -match "^\d+$") { 1 } else { 2 })
        Assert-Equal $paragraph.IndentLevel $expectedIndent "$Label line $entryIndex indentation"
        $expectedColor = $(if ($entry.Section -eq $CurrentSection) { ConvertTo-Rgb 27 38 53 } else { ConvertTo-Rgb 170 179 188 })
        Assert-Equal $character.Font.Color.RGB $expectedColor "$Label line $entryIndex current-section emphasis"
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
$tocSpecifications = @(
    [pscustomobject]@{ Slide = 8; Current = 3 },
    [pscustomobject]@{ Slide = 19; Current = 4 },
    [pscustomobject]@{ Slide = 27; Current = 5 }
)

$powerPoint = $null
$presentation = $null
$existingPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
if ($null -ne $existingPowerPoint) { throw "PowerPoint is already running. Close it before validation." }
$securityPath = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
$securityItem = Get-ItemProperty -Path $securityPath -ErrorAction SilentlyContinue
$accessVbomExisted = ($null -ne $securityItem -and $securityItem.PSObject.Properties.Name -contains "AccessVBOM")
$originalAccessVbom = $(if ($accessVbomExisted) { $securityItem.AccessVBOM } else { $null })

try {
    if (-not (Test-Path -LiteralPath $securityPath)) { [void](New-Item -Path $securityPath -Force) }
    New-ItemProperty -Path $securityPath -Name "AccessVBOM" -PropertyType DWord -Value 1 -Force | Out-Null
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.WindowState = 2
    $powerPoint.DisplayAlerts = 1
    $presentation = $powerPoint.Presentations.Open($Deck, $false, $false, $true)

    Assert-Equal $presentation.Slides.Count 29 "fresh reopen slide count"
    Assert-Equal $presentation.Designs.Count 1 "fresh reopen design count"
    Assert-Equal $presentation.SlideMaster.CustomLayouts.Count 1 "fresh reopen layout count"
    Assert-Equal $presentation.VBProject.VBComponents.Count 1 "fresh reopen VBA component count"
    Assert-Equal $presentation.VBProject.VBComponents.Item(1).Name "SciTeXNavigation" "fresh reopen VBA module"
    $sourceModule = $presentation.VBProject.VBComponents.Item(1).CodeModule
    $source = $sourceModule.Lines(1, $sourceModule.CountOfLines)
    Assert-Equal ([regex]::Matches($source, "(?im)^\s*Public\s+Sub\s+").Count) 1 "public macro count"
    $versionMatch = [regex]::Match($source, 'NAVIGATION_VERSION\s+As\s+String\s*=\s*"([^"]+)"')
    Assert-True $versionMatch.Success "source version constant"
    Assert-Equal $versionMatch.Groups.Item(1).Value "0.1.1" "source version"

    $sentinelBodySize = $presentation.Slides.Item(25).Shapes.Item("Box133").TextFrame.TextRange.Font.Size
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"

    foreach ($mapping in $mappings) {
        $title = $presentation.Slides.Item($mapping.Slide).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text
        Assert-True ($title.StartsWith("$($mapping.Code). ")) "slide $($mapping.Slide) stable explicit code"
    }
    $leftEntries = Get-ExpectedEntries $presentation $mappings $true $false
    $rightEntries = Get-ExpectedEntries $presentation $mappings $false $false
    Assert-Equal $leftEntries.Count 15 "visible left-column entry count"
    Assert-Equal $rightEntries.Count 8 "visible right-column entry count"
    foreach ($tocSpec in $tocSpecifications) {
        $tocSlide = $presentation.Slides.Item($tocSpec.Slide)
        Assert-TocColumn $tocSlide.Shapes.Item("SCITEX_TOC_BODY_LEFT") $leftEntries $tocSpec.Current "TOC $($tocSpec.Slide) left"
        Assert-TocColumn $tocSlide.Shapes.Item("SCITEX_TOC_BODY_RIGHT") $rightEntries $tocSpec.Current "TOC $($tocSpec.Slide) right"
    }
    Assert-Equal $presentation.Slides.Item(25).Shapes.Item("Box133").TextFrame.TextRange.Font.Size $sentinelBodySize "body typography unchanged after navigation run"

    $config = $presentation.Slides.Item(29)
    Assert-True ($config.SlideShowTransition.Hidden -ne 0) "configuration slide hidden"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_FONT_LATIN").TextFrame.TextRange.Text "Arial" "Latin font configuration"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_FONT_CJK").TextFrame.TextRange.Text "Yu Gothic" "CJK font configuration"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_FONT_MIN").TextFrame.TextRange.Text "18" "minimum font configuration"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_FONT_MAX").TextFrame.TextRange.Text "32" "maximum font configuration"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text "Yes" "hidden-slide configuration"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_VERSION").TextFrame.TextRange.Text $versionMatch.Groups.Item(1).Value "displayed version matches source"
    $configText = ""
    foreach ($shape in $config.Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) { $configText += $shape.TextFrame.TextRange.Text }
    }
    Assert-True ($configText -notmatch "[^\x00-\x7F]") "configuration slide English ASCII only"

    $config.Shapes.Item("SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text = "No"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $leftWithHidden = Get-ExpectedEntries $presentation $mappings $true $true
    Assert-Equal $leftWithHidden.Count 16 "hidden slide included when configured No"
    Assert-Equal $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text ($leftWithHidden.Title -join "`r") "hidden slide toggle No"
    $config.Shapes.Item("SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text = "Yes"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text ($leftEntries.Title -join "`r") "hidden slide toggle restored Yes"

    $originalTitleSize = $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Font.Size
    $originalTocSize = $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Font.Size
    $config.Shapes.Item("SCITEX_CFG_FONT_MIN").TextFrame.TextRange.Text = "20"
    $config.Shapes.Item("SCITEX_CFG_FONT_MAX").TextFrame.TextRange.Text = "24"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Font.Size 24 "maximum font size applied"
    Assert-Equal $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Font.Size 20 "minimum font size applied"
    Assert-Equal $presentation.Slides.Item(25).Shapes.Item("Box133").TextFrame.TextRange.Font.Size $sentinelBodySize "body typography unchanged during configuration test"
    $config.Shapes.Item("SCITEX_CFG_FONT_MIN").TextFrame.TextRange.Text = "18"
    $config.Shapes.Item("SCITEX_CFG_FONT_MAX").TextFrame.TextRange.Text = "32"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Font.Size $originalTitleSize "title font size reversibly restored"
    Assert-Equal $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Font.Size $originalTocSize "TOC font size reversibly restored"

    $stateBeforeFinalRun = @(
        $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text,
        $presentation.Slides.Item(28).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text,
        $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text,
        $presentation.Slides.Item(27).Shapes.Item("SCITEX_TOC_BODY_RIGHT").TextFrame.TextRange.Text
    ) -join "|"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $stateAfterFinalRun = @(
        $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text,
        $presentation.Slides.Item(28).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text,
        $presentation.Slides.Item(8).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text,
        $presentation.Slides.Item(27).Shapes.Item("SCITEX_TOC_BODY_RIGHT").TextFrame.TextRange.Text
    ) -join "|"
    Assert-Equal $stateAfterFinalRun $stateBeforeFinalRun "final idempotent run"

    # Regression: PowerPoint copies slide tags, so a TOC duplicated from the
    # section-3 position used to keep section 3 emphasized after moving to slide 2.
    $duplicateRange = $presentation.Slides.Item(8).Duplicate()
    $copiedToc = $duplicateRange.Item(1)
    $copiedToc.MoveTo(2)
    Assert-Equal $presentation.Slides.Count 30 "copied TOC slide count"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal $copiedToc.SlideIndex 2 "copied TOC moved position"
    Assert-Equal $copiedToc.Tags.Item("SCITEX_CURRENT_SECTION") "1" "copied TOC inferred current section"
    $copiedLeftBody = $copiedToc.Shapes.Item("SCITEX_TOC_BODY_LEFT")
    $copiedRightBody = $copiedToc.Shapes.Item("SCITEX_TOC_BODY_RIGHT")
    Assert-Equal $copiedLeftBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Color.RGB (ConvertTo-Rgb 27 38 53) "copied TOC section 1 emphasized"
    Assert-Equal $copiedLeftBody.TextFrame.TextRange.Paragraphs(7, 1).Characters(1, 1).Font.Color.RGB (ConvertTo-Rgb 170 179 188) "copied TOC stale section 3 dimmed"
    Assert-Equal $copiedRightBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Color.RGB (ConvertTo-Rgb 170 179 188) "copied TOC right column dimmed"
    $copiedToc.Delete()
    Assert-Equal $presentation.Slides.Count 29 "copied TOC cleanup slide count"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $presentation.Save()

    [ordered]@{
        deck = $Deck
        fresh_reopen = "passed"
        repeated_macro_runs = "passed"
        explicit_numbering = "passed"
        toc_links_checked = 69
        hierarchical_indentation = "passed"
        current_section_emphasis = "passed"
        copied_toc_section_inference = "passed"
        hidden_slide_toggle = "passed"
        typography_limits = "passed"
        typography_reversible = "passed"
        body_content_unchanged = "passed"
        english_configuration = "passed"
        version = $versionMatch.Groups.Item(1).Value
        public_macro_count = 1
    } | ConvertTo-Json
}
catch {
    Write-Error ("AICHI_VALIDATION_FAILED: " + $_.Exception.Message + "`n" + $_.ScriptStackTrace)
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
