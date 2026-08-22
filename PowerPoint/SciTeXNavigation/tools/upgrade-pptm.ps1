param(
    [Parameter(Mandatory = $true)][string]$SourceDeck,
    [Parameter(Mandatory = $true)][string]$OutputDeck,
    [Parameter(Mandatory = $true)][string]$ModulePath,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

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
    throw "PowerPoint is already running. Close it before upgrading so the open presentation is not affected."
}
if (-not (Test-Path -LiteralPath $SourceDeck)) { throw "Source deck does not exist: $SourceDeck" }
if (-not (Test-Path -LiteralPath $ModulePath)) { throw "VBA module does not exist: $ModulePath" }
if ($SourceDeck -eq $OutputDeck) { throw "SourceDeck and OutputDeck must be different paths." }
if ((Test-Path -LiteralPath $OutputDeck) -and -not $Force) { throw "Output deck already exists. Use -Force to replace it: $OutputDeck" }

$securityPath = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
$securityItem = Get-ItemProperty -Path $securityPath -ErrorAction SilentlyContinue
$accessVbomExisted = ($null -ne $securityItem -and $securityItem.PSObject.Properties.Name -contains "AccessVBOM")
$originalAccessVbom = $(if ($accessVbomExisted) { $securityItem.AccessVBOM } else { $null })

try {
    if (-not (Test-Path -LiteralPath $securityPath)) { [void](New-Item -Path $securityPath -Force) }
    New-ItemProperty -Path $securityPath -Name "AccessVBOM" -PropertyType DWord -Value 1 -Force | Out-Null

    $outputDirectory = Split-Path -Parent $OutputDeck
    if (-not (Test-Path -LiteralPath $outputDirectory)) { [void](New-Item -ItemType Directory -Path $outputDirectory -Force) }
    if (Test-Path -LiteralPath $OutputDeck) { Remove-Item -LiteralPath $OutputDeck -Force }
    Copy-Item -LiteralPath $SourceDeck -Destination $OutputDeck
    Unblock-File -LiteralPath $OutputDeck

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.WindowState = 2
    $powerPoint.DisplayAlerts = 1
    $presentation = $powerPoint.Presentations.Open($OutputDeck, $false, $false, $true)

    for ($componentIndex = $presentation.VBProject.VBComponents.Count; $componentIndex -ge 1; $componentIndex--) {
        $component = $presentation.VBProject.VBComponents.Item($componentIndex)
        if ($component.Name -like "SciTeXNavigation*") {
            $presentation.VBProject.VBComponents.Remove($component)
        }
    }
    $importedModule = $presentation.VBProject.VBComponents.Import($ModulePath)
    Invoke-PowerPointMacro $powerPoint "RunSciTeXNavigation"
    $presentation.Save()

    $sourceModule = $importedModule.CodeModule
    $source = $sourceModule.Lines(1, $sourceModule.CountOfLines)
    $versionMatch = [regex]::Match($source, 'NAVIGATION_VERSION\s+As\s+String\s*=\s*"([^"]+)"')
    $publicMacroCount = [regex]::Matches($source, "(?im)^\s*Public\s+Sub\s+").Count
    $tocSlideCount = 0
    foreach ($slide in $presentation.Slides) {
        if ($slide.Tags.Item("SCITEX_TOC") -eq "1") { $tocSlideCount++ }
    }
    [ordered]@{
        source = $SourceDeck
        output = $OutputDeck
        slides = $presentation.Slides.Count
        toc_slides = $tocSlideCount
        module = $importedModule.Name
        version = $(if ($versionMatch.Success) { $versionMatch.Groups.Item(1).Value } else { "unknown" })
        public_macro_count = $publicMacroCount
        macro_run = "passed"
    } | ConvertTo-Json
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
    if ($null -ne $remainingPowerPoint) { $remainingPowerPoint | Wait-Process -Timeout 15 -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 1000
    Restore-AccessVbom $securityPath $accessVbomExisted $originalAccessVbom
    $restoreScript = Join-Path $PSScriptRoot "restore-access-vbom.ps1"
    $originalExistedNumber = $(if ($accessVbomExisted) { 1 } else { 0 })
    $originalValueNumber = $(if ($accessVbomExisted) { [int]$originalAccessVbom } else { 0 })
    $restoreArguments = @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $restoreScript, "-SecurityPath", $securityPath, "-OriginalExisted", $originalExistedNumber, "-OriginalValue", $originalValueNumber)
    Start-Process -FilePath "powershell.exe" -ArgumentList $restoreArguments -WindowStyle Hidden | Out-Null
}
