param(
    [Parameter(Mandatory = $true)][string]$SourceDeck,
    [Parameter(Mandatory = $true)][string]$OutputDeck,
    [Parameter(Mandatory = $true)][string]$ModulePath
)

$ErrorActionPreference = "Stop"
$script:ScaleX = 1.0
$script:ScaleY = 1.0

function ConvertTo-Rgb([int]$Red, [int]$Green, [int]$Blue) {
    return $Red + (256 * $Green) + (65536 * $Blue)
}

function Remove-AllSlideShapes($Slide) {
    for ($shapeIndex = $Slide.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
        $Slide.Shapes.Item($shapeIndex).Delete()
    }
}

function Add-Rectangle($Slide, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [int]$Color, [double]$Transparency = 0) {
    $shape = $Slide.Shapes.AddShape(1, $Left * $script:ScaleX, $Top * $script:ScaleY, $Width * $script:ScaleX, $Height * $script:ScaleY)
    $shape.Name = $Name
    $shape.Fill.ForeColor.RGB = $Color
    $shape.Fill.Transparency = $Transparency
    $shape.Line.Visible = 0
    return $shape
}

function Add-TextBox($Slide, [string]$Name, [string]$Text, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [double]$FontSize, [int]$Color, [bool]$Bold = $false) {
    $shape = $Slide.Shapes.AddTextbox(1, $Left * $script:ScaleX, $Top * $script:ScaleY, $Width * $script:ScaleX, $Height * $script:ScaleY)
    $shape.Name = $Name
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 4
    $shape.TextFrame.MarginBottom = 4
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = "Aptos"
    $shape.TextFrame.TextRange.Font.Size = $FontSize
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    return $shape
}

function Add-ConfigValue($Slide, [string]$Name, [string]$Text, [double]$Left, [double]$Top, [double]$Width, [double]$Height) {
    $ink = ConvertTo-Rgb 27 38 53
    $accent = ConvertTo-Rgb 25 157 179
    $white = ConvertTo-Rgb 255 255 255
    $shape = Add-TextBox $Slide $Name $Text $Left $Top $Width $Height 18 $ink $true
    $shape.Fill.Visible = -1
    $shape.Fill.ForeColor.RGB = $white
    $shape.Line.Visible = -1
    $shape.Line.ForeColor.RGB = $accent
    $shape.Line.Weight = 1.5
    return $shape
}

function Add-TestSlide($Presentation, [string]$Title, [string]$Body, [bool]$IsToc, [string]$SectionTitle = "") {
    $slide = $Presentation.Slides.Add($Presentation.Slides.Count, 12)
    $slide.FollowMasterBackground = -1

    $navy = ConvertTo-Rgb 21 40 66
    $ink = ConvertTo-Rgb 27 38 53
    $muted = ConvertTo-Rgb 93 107 124
    $accent = ConvertTo-Rgb 25 157 179
    $white = ConvertTo-Rgb 255 255 255

    [void](Add-Rectangle $slide "SCITEX_HEADER" 0 0 960 92 $navy)
    [void](Add-Rectangle $slide "SCITEX_ACCENT" 0 92 960 5 $accent)
    [void](Add-TextBox $slide "SCITEX_TITLE" $Title 54 20 850 58 28 $white $true)

    if ($IsToc) {
        $slide.Tags.Add("SCITEX_TOC", "1")
        $slide.Tags.Add("SCITEX_SECTION_TITLE", $SectionTitle)
        $bodyShape = Add-TextBox $slide "SCITEX_TOC_BODY" "Run the macro to build this table of contents." 100 155 750 260 25 $ink $false
        $bodyShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 14
        [void](Add-TextBox $slide "SCITEX_HINT" "Each item becomes a clickable internal slide link." 100 430 750 35 15 $muted $false)
    }
    else {
        [void](Add-TextBox $slide "SCITEX_BODY" $Body 100 160 760 220 26 $ink $false)
        [void](Add-Rectangle $slide "SCITEX_MARK" 100 410 150 8 $accent)
    }

    return $slide
}

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
    throw "PowerPoint is already running. Close it before building so no user presentation is affected."
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

    if (Test-Path -LiteralPath $OutputDeck) {
        Remove-Item -LiteralPath $OutputDeck -Force
    }

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.WindowState = 2
    $powerPoint.DisplayAlerts = 1

    $presentation = $powerPoint.Presentations.Open($SourceDeck, $true, $false, $true)
    $script:ScaleX = $presentation.PageSetup.SlideWidth / 960.0
    $script:ScaleY = $presentation.PageSetup.SlideHeight / 540.0
    $originalSlideCount = $presentation.Slides.Count
    if ($originalSlideCount -lt 2) {
        throw "The source deck needs a cover and a final settings slide."
    }

    for ($slideIndex = $originalSlideCount - 1; $slideIndex -ge 2; $slideIndex--) {
        $presentation.Slides.Item($slideIndex).Delete()
    }

    $cover = $presentation.Slides.Item(1)
    $settings = $presentation.Slides.Item(2)

    Remove-AllSlideShapes $cover
    $navy = ConvertTo-Rgb 21 40 66
    $accent = ConvertTo-Rgb 25 157 179
    $white = ConvertTo-Rgb 255 255 255
    $light = ConvertTo-Rgb 220 230 239

    [void](Add-Rectangle $cover "SCITEX_COVER_BG" 0 0 960 540 $navy)
    [void](Add-Rectangle $cover "SCITEX_COVER_ACCENT" 0 0 18 540 $accent)
    [void](Add-TextBox $cover "SCITEX_COVER_TITLE" "SciTeX Navigation Sandbox" 75 130 800 75 34 $white $true)
    [void](Add-TextBox $cover "SCITEX_COVER_SUBTITLE" "Small, isolated test deck - the AICHI contest deck is untouched" 77 220 780 45 19 $light $false)
    [void](Add-TextBox $cover "SCITEX_RUN_INSTRUCTION" "Alt+F8  >  RunSciTeXNavigation  >  Run" 77 315 650 45 20 $white $true)
    $status = Add-TextBox $cover "SCITEX_STATUS" "Ready" 77 385 650 35 16 $light $false
    $runButton = $cover.Shapes.AddShape(5, 730 * $script:ScaleX, 310 * $script:ScaleY, 155 * $script:ScaleX, 58 * $script:ScaleY)
    $runButton.Name = "SCITEX_RUN_BUTTON"
    $runButton.Fill.ForeColor.RGB = $accent
    $runButton.Line.Visible = 0
    $runButton.TextFrame.TextRange.Text = "Run Navigation"
    $runButton.TextFrame.TextRange.Font.Name = "Aptos"
    $runButton.TextFrame.TextRange.Font.Size = 17
    $runButton.TextFrame.TextRange.Font.Bold = -1
    $runButton.TextFrame.TextRange.Font.Color.RGB = $white
    $runButton.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $cover.Tags.Add("SCITEX_COVER", "1")

    [void](Add-TestSlide $presentation "Company Overview" "" $true "Company Overview")
    [void](Add-TestSlide $presentation "Company Profile" "A compact content slide used to verify automatic numbering." $false)
    [void](Add-TestSlide $presentation "Problem and Solution" "A second content slide used to verify a/b child numbering." $false)
    [void](Add-TestSlide $presentation "Product" "" $true "Product")
    [void](Add-TestSlide $presentation "SciTeX Platform" "A product slide used to verify the second section." $false)

    Remove-AllSlideShapes $settings
    $ink = ConvertTo-Rgb 27 38 53
    $muted = ConvertTo-Rgb 93 107 124
    $panel = ConvertTo-Rgb 241 246 249
    [void](Add-Rectangle $settings "SCITEX_CONFIG_BG" 0 0 960 540 $white)
    [void](Add-Rectangle $settings "SCITEX_CONFIG_HEADER" 0 0 960 76 $navy)
    [void](Add-Rectangle $settings "SCITEX_CONFIG_ACCENT" 0 76 960 5 $accent)
    [void](Add-TextBox $settings "SCITEX_CONFIG_TITLE" "SciTeX Navigation - Configuration" 48 14 860 48 26 $white $true)
    [void](Add-TextBox $settings "SCITEX_CONFIG_HELP" "Edit the value boxes, then run RunSciTeXNavigation again." 60 92 840 35 16 $muted $false)
    [void](Add-Rectangle $settings "SCITEX_CONFIG_TYPOGRAPHY_PANEL" 55 132 850 180 $panel)
    [void](Add-TextBox $settings "SCITEX_CONFIG_TYPOGRAPHY_TITLE" "Typography" 78 145 250 34 20 $ink $true)
    [void](Add-TextBox $settings "SCITEX_CONFIG_LATIN_LABEL" "Latin font" 78 194 155 34 16 $muted $false)
    [void](Add-ConfigValue $settings "SCITEX_CFG_FONT_LATIN" "Aptos" 230 187 245 40)
    [void](Add-TextBox $settings "SCITEX_CONFIG_CJK_LABEL" "CJK font" 78 246 155 34 16 $muted $false)
    [void](Add-ConfigValue $settings "SCITEX_CFG_FONT_CJK" "Yu Gothic" 230 239 245 40)
    [void](Add-TextBox $settings "SCITEX_CONFIG_MIN_LABEL" "Minimum size (pt)" 505 194 220 34 16 $muted $false)
    [void](Add-ConfigValue $settings "SCITEX_CFG_FONT_MIN" "18" 770 187 85 40)
    [void](Add-TextBox $settings "SCITEX_CONFIG_MAX_LABEL" "Maximum size (pt)" 505 246 220 34 16 $muted $false)
    [void](Add-ConfigValue $settings "SCITEX_CFG_FONT_MAX" "32" 770 239 85 40)
    [void](Add-Rectangle $settings "SCITEX_CONFIG_TOC_PANEL" 55 332 850 105 $panel)
    [void](Add-TextBox $settings "SCITEX_CONFIG_TOC_TITLE" "Table of contents" 78 345 300 34 20 $ink $true)
    [void](Add-TextBox $settings "SCITEX_CONFIG_HIDE_LABEL" "Hide hidden slides from TOC" 78 392 390 34 16 $muted $false)
    [void](Add-TextBox $settings "SCITEX_CONFIG_VERSION_LABEL" "Version" 505 392 100 34 16 $muted $false)
    [void](Add-ConfigValue $settings "SCITEX_CFG_VERSION" "0.1.1" 605 385 120 40)
    [void](Add-ConfigValue $settings "SCITEX_CFG_HIDE_HIDDEN" "Yes" 770 385 85 40)
    [void](Add-TextBox $settings "SCITEX_CONFIG_FOOTER" "Accepted values: Yes / No. This configuration slide is always excluded from the TOC and slide show." 60 458 840 45 14 $muted $false)
    $settings.Tags.Add("SCITEX_CONFIG", "1")
    $settings.Tags.Add("SCITEX_ALWAYS_SKIP", "1")
    $settings.SlideShowTransition.Hidden = -1

    $presentation.SaveAs($OutputDeck, 25)
    $importedModule = $presentation.VBProject.VBComponents.Import($ModulePath)
    $runButton.ActionSettings.Item(1).Action = 8
    $runButton.ActionSettings.Item(1).Run = "RunSciTeXNavigation"
    $presentation.Save()

    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"

    Assert-Equal $presentation.Slides.Count 7 "slide count after first run"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TITLE").TextFrame.TextRange.Text "1. Company Overview" "section 1 title"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(3) "SCITEX_TITLE").TextFrame.TextRange.Text "1a. Company Profile" "section 1 child a title"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(4) "SCITEX_TITLE").TextFrame.TextRange.Text "1b. Problem and Solution" "section 1 child b title"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TITLE").TextFrame.TextRange.Text "2. Product" "section 2 title"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(6) "SCITEX_TITLE").TextFrame.TextRange.Text "2a. SciTeX Platform" "section 2 child title"
    $fullToc = "1. Company Overview`r1a. Company Profile`r1b. Problem and Solution`r2. Product`r2a. SciTeX Platform"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "section 1 full TOC"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "section 2 full TOC"
    $currentColor = ConvertTo-Rgb 27 38 53
    $dimmedColor = ConvertTo-Rgb 170 179 188
    $tocOneBody = Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY"
    $tocTwoBody = Get-NamedShape $presentation.Slides.Item(5) "SCITEX_TOC_BODY"
    Assert-Equal $tocOneBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Color.RGB $currentColor "TOC 1 current section color"
    Assert-Equal $tocOneBody.TextFrame.TextRange.Paragraphs(4, 1).Characters(1, 1).Font.Color.RGB $dimmedColor "TOC 1 other section dimmed"
    Assert-Equal $tocTwoBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Color.RGB $dimmedColor "TOC 2 other section dimmed"
    Assert-Equal $tocTwoBody.TextFrame.TextRange.Paragraphs(4, 1).Characters(1, 1).Font.Color.RGB $currentColor "TOC 2 current section color"
    $expectedTargetSlides = @(2, 3, 4, 5, 6)
    foreach ($tocSlideIndex in @(2, 5)) {
        $tocBody = Get-NamedShape $presentation.Slides.Item($tocSlideIndex) "SCITEX_TOC_BODY"
        for ($lineIndex = 1; $lineIndex -le $expectedTargetSlides.Count; $lineIndex++) {
            $link = $tocBody.TextFrame.TextRange.Paragraphs($lineIndex, 1).Characters(1, 1).ActionSettings.Item(1)
            Assert-Equal $link.Action 7 "TOC slide $tocSlideIndex line $lineIndex hyperlink action"
            Assert-True ($link.Hyperlink.SubAddress -match ",$($expectedTargetSlides[$lineIndex - 1]),") "TOC slide $tocSlideIndex line $lineIndex target"
        }
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).Font.Bold -1 "TOC slide $tocSlideIndex section 1 heading bold"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(4, 1).Characters(1, 1).Font.Bold -1 "TOC slide $tocSlideIndex section 2 heading bold"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(1, 1).IndentLevel 1 "TOC slide $tocSlideIndex section 1 indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(2, 1).IndentLevel 2 "TOC slide $tocSlideIndex child 1a indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(3, 1).IndentLevel 2 "TOC slide $tocSlideIndex child 1b indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(4, 1).IndentLevel 1 "TOC slide $tocSlideIndex section 2 indent level"
        Assert-Equal $tocBody.TextFrame.TextRange.Paragraphs(5, 1).IndentLevel 2 "TOC slide $tocSlideIndex child 2a indent level"
    }
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(1) "SCITEX_RUN_BUTTON").ActionSettings.Item(1).Action 8 "cover run button action"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(1) "SCITEX_RUN_BUTTON").ActionSettings.Item(1).Run "RunSciTeXNavigation" "cover run button macro"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(1) "SCITEX_STATUS").TextFrame.TextRange.Text "Navigation v0.1.1 updated - 2 sections" "cover run status"
    Assert-True ($presentation.Slides.Item(7).SlideShowTransition.Hidden -ne 0) "settings slide hidden"
    Assert-Equal $presentation.Slides.Item(7).Tags.Item("SCITEX_CONFIG") "1" "settings slide tag"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_LATIN").TextFrame.TextRange.Text "Aptos" "configured Latin font"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_CJK").TextFrame.TextRange.Text "Yu Gothic" "configured CJK font"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_MIN").TextFrame.TextRange.Text "18" "configured minimum font size"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_FONT_MAX").TextFrame.TextRange.Text "32" "configured maximum font size"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text "Yes" "configured hidden-slide behavior"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_VERSION").TextFrame.TextRange.Text "0.1.1" "configuration version"
    $configText = ""
    foreach ($shape in $presentation.Slides.Item(7).Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
            $configText += $shape.TextFrame.TextRange.Text
        }
    }
    Assert-True ($configText -notmatch "[^\x00-\x7F]") "configuration page contains English ASCII text only"

    $managedTypographyShapes = @("SCITEX_TITLE", "SCITEX_TOC_BODY", "SCITEX_TOC_BODY_LEFT", "SCITEX_TOC_BODY_RIGHT", "SCITEX_STATUS", "SCITEX_RUN_BUTTON")
    for ($testSlideIndex = 1; $testSlideIndex -le 6; $testSlideIndex++) {
        foreach ($shape in $presentation.Slides.Item($testSlideIndex).Shapes) {
            if ($managedTypographyShapes -contains $shape.Name -and $shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $textRange = $shape.TextFrame.TextRange
                for ($characterIndex = 1; $characterIndex -le $textRange.Length; $characterIndex++) {
                    $character = $textRange.Characters($characterIndex, 1)
                    Assert-True ($character.Font.Size -ge 18 -and $character.Font.Size -le 32) "slide $testSlideIndex shape $($shape.Name) typography size bounds"
                    Assert-Equal $character.Font.Name "Aptos" "slide $testSlideIndex shape $($shape.Name) Latin font"
                }
            }
        }
    }

    $hiddenToc = "1. Company Overview`r1a. Company Profile`r2. Product`r2a. SciTeX Platform"
    $presentation.Slides.Item(4).SlideShowTransition.Hidden = -1
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $hiddenToc "hidden slide excluded with Yes"
    (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text = "No"
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "hidden slide included with No"
    (Get-NamedShape $presentation.Slides.Item(7) "SCITEX_CFG_HIDE_HIDDEN").TextFrame.TextRange.Text = "Yes"
    $presentation.Slides.Item(4).SlideShowTransition.Hidden = 0
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    Assert-Equal (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Text $fullToc "visible slide restored in final TOC"

    $firstLink = (Get-NamedShape $presentation.Slides.Item(2) "SCITEX_TOC_BODY").TextFrame.TextRange.Paragraphs(1, 1).Characters(1, 1).ActionSettings.Item(1)
    Assert-Equal $firstLink.Action 7 "first TOC hyperlink action"
    Assert-True (-not [string]::IsNullOrWhiteSpace($firstLink.Hyperlink.SubAddress)) "first TOC hyperlink target"

    $presentation.Save()

    $result = [ordered]@{
        output = $OutputDeck
        source_slides = $originalSlideCount
        sandbox_slides = $presentation.Slides.Count
        public_macro = "RunSciTeXNavigation"
        vba_components = $presentation.VBProject.VBComponents.Count
        first_run = "passed"
        numbering = "passed"
        toc_rebuild = "passed"
        internal_links = "passed"
        current_section_emphasis = "passed"
        all_other_sections_dimmed = "passed"
        hierarchical_indentation = "passed"
        typography_configuration = "passed"
        font_size_bounds = "passed"
        hidden_slide_toc_toggle = "passed"
        english_only_configuration_page = "passed"
        version = "0.1.1"
        run_button = "passed"
        settings_slide_preserved_hidden = "passed"
    }
    $result | ConvertTo-Json
}
catch {
    Write-Error ("BUILD_FAILED: " + $_.Exception.Message + "`n" + $_.ScriptStackTrace)
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
