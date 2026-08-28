param(
    [Parameter(Mandatory = $true)][string]$Deck
)

$ErrorActionPreference = "Stop"

function Get-NamedShape($Slide, [string]$Name) {
    return $Slide.Shapes.Item($Name)
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

function ConvertTo-Rgb([int]$Red, [int]$Green, [int]$Blue) {
    return $Red + (256 * $Green) + (65536 * $Blue)
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

$powerPoint = $null
$presentation = $null
$existingPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
if ($null -ne $existingPowerPoint) {
    throw "PowerPoint is already running. Close it before testing so no user presentation is affected."
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

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.WindowState = 2
    $powerPoint.DisplayAlerts = 1
    $presentation = $powerPoint.Presentations.Open($Deck, $false, $false, $true)

    Assert-Equal $presentation.Slides.Count 7 "reopened slide count before run"
    Assert-Equal $presentation.VBProject.VBComponents.Count 1 "VBA component count"
    Assert-Equal $presentation.VBProject.VBComponents.Item(1).Name "SciTeXNavigation" "VBA module name"

    $module = $presentation.VBProject.VBComponents.Item(1).CodeModule
    $source = $module.Lines(1, $module.CountOfLines)
    $publicSubCount = [regex]::Matches($source, "(?im)^\s*Public\s+Sub\s+").Count
    Assert-Equal $publicSubCount 1 "public macro count"

    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"

    Assert-Equal $presentation.Slides.Count 7 "slide count after fresh reopen and second run"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TITLE").TextFrame.TextRange.Text "Company Overview" "reopened section 1 TOC title"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(3) "SCITEX_TITLE").TextFrame.TextRange.Text "1a. Company Profile" "idempotent child numbering"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(4) "SCITEX_TITLE").TextFrame.TextRange.Text "1b. Problem and Solution" "idempotent second child numbering"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TITLE").TextFrame.TextRange.Text "Product" "reopened section 2 TOC title"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(6) "SCITEX_TITLE").TextFrame.TextRange.Text "2a. SciTeX Platform" "reopened section 2 child title"
    $fullToc = "1. Company Overview`r1a. Company Profile`r1b. Problem and Solution`r2. Product`r2a. SciTeX Platform"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "reopened section 1 full TOC"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "reopened section 2 full TOC"
    $currentColor = ConvertTo-Rgb 27 38 53
    $dimmedColor = ConvertTo-Rgb 170 179 188
    $tocOneBody = Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY"
    $tocTwoBody = Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TOC_BODY"
    Assert-Equal $tocOneBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Color.RGB $currentColor "reopened TOC 1 current section color"
    Assert-Equal $tocOneBody.TextFrame.TextRange.Paragraphs(4, 1).Characters(1, 1).Font.Color.RGB $dimmedColor "reopened TOC 1 other section dimmed"
    Assert-Equal $tocTwoBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Color.RGB $dimmedColor "reopened TOC 2 other section dimmed"
    Assert-Equal $tocTwoBody.TextFrame.TextRange.Paragraphs(4, 1).Characters(1, 1).Font.Color.RGB $currentColor "reopened TOC 2 current section color"
    Assert-True ($presentation.Slides.Item(7).SlideShowTransition.Hidden -ne 0) "settings slide stays hidden"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_LATIN").TextFrame.TextRange.Text "Aptos" "reopened Latin font configuration"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_CJK").TextFrame.TextRange.Text "Yu Gothic" "reopened CJK font configuration"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_MIN").TextFrame.TextRange.Text "18" "reopened minimum font size configuration"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_MAX").TextFrame.TextRange.Text "32" "reopened maximum font size configuration"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text "Yes" "reopened hidden-slide configuration"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_VERSION").TextFrame.TextRange.Text "0.1.2" "reopened configuration version"
    $configText = ""
    foreach ($shape in $presentation.Slides.Item(7).Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
            $configText += $shape.TextFrame.TextRange.Text
        }
    }
    Assert-True ($configText -notmatch "[^\x00-\x7F]") "reopened configuration page contains English ASCII text only"

    $expectedTargetSlides = @(2, 3, 4, 5, 6)
    foreach ($tocSlideIndex in @(2, 5)) {
        $tocBody = Get-NamedShape $presentation.Slides.Item($tocSlideIndex) "SCITEX_TOC_BODY"
        $tocTitle = Get-NamedShape $presentation.Slides.Item($tocSlideIndex) "SCITEX_TITLE"
        Assert-True ($tocTitle.TextFrame.TextRange.BoundWidth -le $tocTitle.Width + 0.5) "reopened TOC slide $tocSlideIndex title fits its header box"
        Assert-True ($tocTitle.TextFrame.TextRange.Font.Size -ge 18 -and $tocTitle.TextFrame.TextRange.Font.Size -le 32) "reopened TOC slide $tocSlideIndex title size bounds"
        for ($lineIndex = 1; $lineIndex -le $expectedTargetSlides.Count; $lineIndex++) {
            $textAction = $tocBody.TextFrame.TextRange.Paragraphs($lineIndex, 1).Characters(1, 1).ActionSettings.Item(1)
            Assert-Equal $textAction.Action 0 "reopened TOC slide $tocSlideIndex line $lineIndex text has no hyperlink styling"
            $link = $presentation.Slides.Item($tocSlideIndex).Shapes.Item("SCITEX_TOC_LINK_B_$lineIndex").ActionSettings.Item(1)
            Assert-Equal $link.Action 7 "reopened TOC slide $tocSlideIndex line $lineIndex overlay hyperlink action"
            Assert-True ($link.Hyperlink.SubAddress -match ",$($expectedTargetSlides[$lineIndex - 1]),") "reopened TOC slide $tocSlideIndex line $lineIndex target"
        }
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(1, 1).IndentLevel 1 "reopened TOC slide $tocSlideIndex section 1 indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(2, 1).IndentLevel 2 "reopened TOC slide $tocSlideIndex child 1a indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(3, 1).IndentLevel 2 "reopened TOC slide $tocSlideIndex child 1b indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(4, 1).IndentLevel 1 "reopened TOC slide $tocSlideIndex section 2 indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(5, 1).IndentLevel 2 "reopened TOC slide $tocSlideIndex child 2a indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Underline -1 "reopened TOC slide $tocSlideIndex section 1 heading underlined"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(2, 1).Characters(1, 1).Font.Underline 0 "reopened TOC slide $tocSlideIndex child 1a not underlined"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(4, 1).Characters(1, 1).Font.Underline -1 "reopened TOC slide $tocSlideIndex section 2 heading underlined"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(5, 1).Characters(1, 1).Font.Underline 0 "reopened TOC slide $tocSlideIndex child 2a not underlined"
    }

    $hiddenToc = "1. Company Overview`r1a. Company Profile`r2. Product`r2a. SciTeX Platform"
    $presentation.Slides.Item(4).SlideShowTransition.Hidden = -1
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $hiddenToc "reopened hidden slide excluded from TOC"
    $presentation.Slides.Item(4).SlideShowTransition.Hidden = 0
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "reopened visible slide restored to TOC"

    $stateBeforeThirdRun = @(
        (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(3) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(4) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(6) "SCITEX_TITLE").TextFrame.TextRange.Text,
        $tocOneBody.TextFrame.TextRange.Text,
        $tocTwoBody.TextFrame.TextRange.Text
    ) -join "|"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $stateAfterThirdRun = @(
        (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(3) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(4) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(6) "SCITEX_TITLE").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text,
        (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TOC_BODY").TextFrame.TextRange.Text
    ) -join "|"
    Assert-Equal $presentation.Slides.Count 7 "slide count after third run"
    Assert-Equal $stateAfterThirdRun $stateBeforeThirdRun "third run idempotence"

    $presentation.Save()
    $result = [ordered]@{
        deck = $Deck
        fresh_powerpoint_reopen = "passed"
        second_macro_run = "passed"
        third_macro_run = "passed"
        idempotent_numbering = "passed"
        toc_links_persist = "passed"
        full_toc_on_every_toc_slide = "passed"
        current_section_emphasis = "passed"
        all_other_sections_dimmed = "passed"
        hierarchical_indentation = "passed"
        toc_title_fit = "passed"
        typography_configuration = "passed"
        font_size_bounds = "passed"
        hidden_slide_toc_toggle = "passed"
        english_only_configuration_page = "passed"
        version = "0.1.2"
        settings_slide_hidden = "passed"
        public_macro_count = $publicSubCount
    }
    $result | ConvertTo-Json
}
catch {
    Write-Error ("REOPEN_TEST_FAILED: " + $_.Exception.Message + "`n" + $_.ScriptStackTrace)
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
    $remainingPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($null -ne $remainingPowerPoint) {
        $remainingPowerPoint | Wait-Process -Timeout 15 -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 1000
    Restore-AccessVbom $securityPath $accessVbomExisted $originalAccessVbom
    $finalSecurityItem = Get-ItemProperty -Path $securityPath -ErrorAction SilentlyContinue
    $finalAccessVbomExists = ($null -ne $finalSecurityItem -and $finalSecurityItem.PSObject.Properties.Name -contains "AccessVBOM")
    if ($finalAccessVbomExists -ne $accessVbomExisted) {
        throw "AccessVBOM cleanup validation failed."
    }
    if ($accessVbomExisted -and $finalSecurityItem.AccessVBOM -ne $originalAccessVbom) {
        throw "AccessVBOM original value was not restored."
    }
    $restoreScript = Join-Path $PSScriptRoot "restore-access-vbom.ps1"
    $originalExistedNumber = $(if ($accessVbomExisted) { 1 } else { 0 })
    $originalValueNumber = $(if ($accessVbomExisted) { [int]$originalAccessVbom } else { 0 })
    $restoreArguments = @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $restoreScript, "-SecurityPath", $securityPath, "-OriginalExisted", $originalExistedNumber, "-OriginalValue", $originalValueNumber)
    Start-Process -FilePath "powershell.exe" -ArgumentList $restoreArguments -WindowStyle Hidden | Out-Null
}
