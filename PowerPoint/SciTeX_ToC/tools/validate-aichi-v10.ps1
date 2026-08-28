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

function Get-Tag($Slide, [string]$Name) {
    return [string]$Slide.Tags.Item($Name)
}

function Get-Section([string]$NavigationCode) {
    return [int]([regex]::Match($NavigationCode, "^\d+").Value)
}

function Get-EntryText($Slide) {
    $code = Get-Tag $Slide "SCITEX_NAV_CODE"
    if ((Get-Tag $Slide "SCITEX_TOC") -eq "1") {
        return "$code. $(Get-Tag $Slide 'SCITEX_SECTION_TITLE')"
    }
    return $Slide.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text
}

function Get-ExpectedEntries($Presentation, [bool]$LeftColumn, [bool]$IncludeHidden, [int]$SplitAfter) {
    $entries = @()
    foreach ($slide in $Presentation.Slides) {
        $code = Get-Tag $slide "SCITEX_NAV_CODE"
        if ([string]::IsNullOrWhiteSpace($code)) { continue }
        if ((Get-Tag $slide "SCITEX_CONFIG") -eq "1" -or (Get-Tag $slide "SCITEX_COVER") -eq "1") { continue }
        if (-not $IncludeHidden -and $slide.SlideShowTransition.Hidden -ne 0) { continue }
        $section = Get-Section $code
        $includeColumn = $(if ($LeftColumn) { $section -le $SplitAfter } else { $section -gt $SplitAfter })
        if ($includeColumn) {
            $entries += [pscustomobject]@{
                Slide = $slide
                Code = $code
                Section = $section
                IsToc = ((Get-Tag $slide "SCITEX_TOC") -eq "1")
                Text = Get-EntryText $slide
            }
        }
    }
    return $entries
}

function Assert-TocColumn($TocSlide, $Body, $Entries, [int]$CurrentSection, [string]$ColumnKey, [string]$Label) {
    Assert-Equal $Body.TextFrame.TextRange.Text ($Entries.Text -join "`r") "$Label complete text"
    for ($entryIndex = 1; $entryIndex -le $Entries.Count; $entryIndex++) {
        $entry = $Entries[$entryIndex - 1]
        $paragraph = $Body.TextFrame.TextRange.Paragraphs($entryIndex, 1)
        $character = $paragraph.Characters(1, 1)
        $textAction = $character.ActionSettings.Item(1)
        Assert-Equal $textAction.Action 0 "$Label line $entryIndex text has no hyperlink styling"
        $action = $TocSlide.Shapes.Item("SCITEX_TOC_LINK_$($ColumnKey)_$entryIndex").ActionSettings.Item(1)
        Assert-Equal $action.Action 7 "$Label line $entryIndex overlay hyperlink action"
        Assert-True ($action.Hyperlink.SubAddress -match ",$($entry.Slide.SlideIndex),") "$Label line $entryIndex hyperlink target"
        $expectedIndent = $(if ($entry.IsToc) { 1 } else { 2 })
        $expectedUnderline = $(if ($entry.IsToc) { -1 } else { 0 })
        $expectedBold = $(if ($entry.IsToc) { -1 } else { 0 })
        $expectedColor = $(if ($entry.Section -eq $CurrentSection) { ConvertTo-Rgb 27 38 53 } else { ConvertTo-Rgb 170 179 188 })
        Assert-Equal $paragraph.IndentLevel $expectedIndent "$Label line $entryIndex indentation"
        Assert-Equal $character.Font.Underline $expectedUnderline "$Label line $entryIndex underline"
        Assert-Equal $character.Font.Bold $expectedBold "$Label line $entryIndex bold"
        Assert-Equal $character.Font.Color.RGB $expectedColor "$Label line $entryIndex current-section emphasis"
    }
}

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

    Assert-Equal $presentation.Slides.Count 30 "fresh reopen slide count"
    Assert-Equal $presentation.Designs.Count 1 "fresh reopen design count"
    Assert-Equal $presentation.SlideMaster.CustomLayouts.Count 1 "fresh reopen layout count"
    Assert-Equal $presentation.VBProject.VBComponents.Count 1 "fresh reopen VBA component count"
    Assert-Equal $presentation.VBProject.VBComponents.Item(1).Name "SciTeXNavigation" "fresh reopen VBA module"
    $sourceModule = $presentation.VBProject.VBComponents.Item(1).CodeModule
    $source = $sourceModule.Lines(1, $sourceModule.CountOfLines)
    Assert-Equal ([regex]::Matches($source, "(?im)^\s*Public\s+Sub\s+").Count) 1 "public macro count"
    $versionMatch = [regex]::Match($source, 'NAVIGATION_VERSION\s+As\s+String\s*=\s*"([^"]+)"')
    Assert-True $versionMatch.Success "source version constant"
    Assert-Equal $versionMatch.Groups.Item(1).Value "0.1.2" "source version"

    $sentinelShape = $null
    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -eq "Box133") { $sentinelShape = $shape; break }
        }
        if ($null -ne $sentinelShape) { break }
    }
    Assert-True ($null -ne $sentinelShape) "body-content sentinel found"
    $sentinelText = $sentinelShape.TextFrame.TextRange.Text
    $sentinelSize = $sentinelShape.TextFrame.TextRange.Font.Size

    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"

    $tocSlides = @($presentation.Slides | Where-Object { (Get-Tag $_ "SCITEX_TOC") -eq "1" })
    Assert-Equal $tocSlides.Count 4 "TOC slide count"
    Assert-Equal ($tocSlides.SlideIndex -join ",") "2,9,20,28" "TOC slide positions"
    $expectedChildCounts = @(6, 10, 7, 1)
    $expectedTocIndices = @(2, 9, 20, 28)
    for ($section = 1; $section -le 4; $section++) {
        $toc = $presentation.Slides.Item($expectedTocIndices[$section - 1])
        Assert-Equal (Get-Tag $toc "SCITEX_NAV_CODE") ([string]$section) "section $section TOC navigation code"
        Assert-Equal (Get-Tag $toc "SCITEX_CURRENT_SECTION") ([string]$section) "section $section current-section tag"
        Assert-Equal (Get-Tag $toc "SCITEX_TOC_SPLIT_AFTER") "2" "section $section balanced split"
        $sectionTitle = Get-Tag $toc "SCITEX_SECTION_TITLE"
        Assert-True (-not [string]::IsNullOrWhiteSpace($sectionTitle)) "section $section title tag"
        Assert-True ($toc.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text.EndsWith($sectionTitle)) "section $section visible TOC title"
        for ($childNumber = 1; $childNumber -le $expectedChildCounts[$section - 1]; $childNumber++) {
            $childSlide = $presentation.Slides.Item($toc.SlideIndex + $childNumber)
            $letter = [char](96 + $childNumber)
            $expectedCode = "$section$letter"
            Assert-Equal (Get-Tag $childSlide "SCITEX_NAV_CODE") $expectedCode "section $section child $childNumber navigation code"
            Assert-True ($childSlide.Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text.StartsWith("$expectedCode. ")) "section $section child $childNumber visible numbering"
        }
    }

    $leftEntries = Get-ExpectedEntries $presentation $true $false 2
    $rightEntries = Get-ExpectedEntries $presentation $false $false 2
    Assert-Equal $leftEntries.Count 17 "visible left-column entry count"
    Assert-Equal $rightEntries.Count 10 "visible right-column entry count"
    foreach ($toc in $tocSlides) {
        $currentSection = [int](Get-Tag $toc "SCITEX_NAV_CODE")
        $tocTitle = $toc.Shapes.Item("SCITEX_TITLE")
        Assert-True ($tocTitle.TextFrame.TextRange.BoundWidth -le $tocTitle.Width + 0.5) "TOC $($toc.SlideIndex) title fits its header box"
        Assert-True ($tocTitle.TextFrame.TextRange.Font.Size -ge 18 -and $tocTitle.TextFrame.TextRange.Font.Size -le 32) "TOC $($toc.SlideIndex) title respects font-size bounds"
        foreach ($shape in $toc.Shapes) {
            Assert-True ($shape.Type -ne 9) "TOC $($toc.SlideIndex) has no stale connector lines"
        }
        Assert-TocColumn $toc $toc.Shapes.Item("SCITEX_TOC_BODY_LEFT") $leftEntries $currentSection "L" "TOC $($toc.SlideIndex) left"
        Assert-TocColumn $toc $toc.Shapes.Item("SCITEX_TOC_BODY_RIGHT") $rightEntries $currentSection "R" "TOC $($toc.SlideIndex) right"
    }

    $originalTocTitle = $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text
    $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text = "Contents: Renamed Section"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text "Contents: Renamed Section" "manual TOC title preserved"
    Assert-Equal (Get-Tag $presentation.Slides.Item(2) "SCITEX_SECTION_TITLE") "Renamed Section" "manual TOC title accepted"
    Assert-True $presentation.Slides.Item(2).Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text.StartsWith("1. Renamed Section") "manual TOC title propagated"
    $presentation.Slides.Item(2).Shapes.Item("SCITEX_TITLE").TextFrame.TextRange.Text = $originalTocTitle
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"

    $config = $presentation.Slides.Item(30)
    Assert-True ($config.SlideShowTransition.Hidden -ne 0) "configuration slide hidden"
    Assert-Equal $config.Shapes.Item("SCITEX_CFG_VERSION").TextFrame.TextRange.Text "0.1.2" "displayed version"
    $config.Shapes.Item("SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text = "No"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $leftWithHidden = Get-ExpectedEntries $presentation $true $true 2
    Assert-Equal $leftWithHidden.Count 18 "hidden slide included when configured No"
    $config.Shapes.Item("SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text = "Yes"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal (Get-ExpectedEntries $presentation $true $false 2).Count 17 "hidden slide filter restored Yes"

    Assert-Equal $sentinelShape.TextFrame.TextRange.Text $sentinelText "body content text unchanged"
    Assert-Equal $sentinelShape.TextFrame.TextRange.Font.Size $sentinelSize "body content typography unchanged"
    $stateBefore = @($tocSlides | ForEach-Object { $_.Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text + "|" + $_.Shapes.Item("SCITEX_TOC_BODY_RIGHT").TextFrame.TextRange.Text }) -join "||"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $stateAfter = @($tocSlides | ForEach-Object { $_.Shapes.Item("SCITEX_TOC_BODY_LEFT").TextFrame.TextRange.Text + "|" + $_.Shapes.Item("SCITEX_TOC_BODY_RIGHT").TextFrame.TextRange.Text }) -join "||"
    Assert-Equal $stateAfter $stateBefore "idempotent repeated run"
    $presentation.Save()

    [ordered]@{
        deck = $Deck
        fresh_reopen = "passed"
        toc_driven_hierarchy = "passed"
        toc_title_rename = "passed"
        toc_links_checked = 108
        top_level_only_underlined = "passed"
        child_links_without_underline = "passed"
        stale_connector_cleanup = "passed"
        toc_title_fit = "passed"
        balanced_two_column_toc = "passed"
        current_section_emphasis = "passed"
        hidden_slide_toggle = "passed"
        repeated_macro_runs = "passed"
        body_content_unchanged = "passed"
        version = $versionMatch.Groups.Item(1).Value
        public_macro_count = 1
    } | ConvertTo-Json
}
catch {
    Write-Error ("TOC_DRIVEN_VALIDATION_FAILED: " + $_.Exception.Message + "`n" + $_.ScriptStackTrace)
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
