<#
.SYNOPSIS
Apply SciTeXNavigation to a real deck and save the result beside it.

.DESCRIPTION
build-and-test.ps1 builds a synthetic sandbox to exercise the macro. This runs
it against a deck someone actually presents.

Two rules are enforced here rather than left to the caller, because both have
already cost something:

  NEVER TOUCH AN INSTANCE WE DID NOT START.  Attaching to a running PowerPoint
  and calling Quit() closed the operator's two open presentations and lost
  their unsaved work (2026-08-27).  This script always starts its own instance,
  records that it did, and quits only that one.

  NEVER LEAVE AccessVBOM ON.  Importing a module needs the Trust Center gate
  open.  It is recorded, opened, and put back -- and the restore also runs from
  the finally block, so a crash mid-run still closes it.

The source deck is never modified: it is copied to -Output first and the macro
runs on the copy.  The macro takes its own timestamped backup as well.

.PARAMETER Deck
The .pptx/.pptm to lay out.  Read only.

.PARAMETER ModulePath
The exported SciTeXNavigation.bas to import.

.PARAMETER Output
Where to write the laid-out deck.  Must be .pptm -- a deck with a macro in it
cannot be saved as .pptx.

.NOTES
LAUNCHING FROM WSL.  Calling powershell.exe directly across the WSL interop
socket fails intermittently -- it times out with

    WSL ERROR: UtilAcceptVsock:273: accept4 failed 110

and writes NOTHING but that line, so the run looks like a script that produced
no output rather than one that never started.  cmd.exe crosses the same socket
fine, and PowerShell launched underneath it works, so go through cmd:

    cd /mnt/c/Users/<user>          # cmd cannot start in a \\wsl.localhost path
    cmd.exe /c "powershell -NoProfile -ExecutionPolicy Bypass -File C:\...\apply-navigation.ps1 ..."

Observed 2026-08-28: direct launch failed twice in a row while
`cmd.exe /c powershell -Command Write-Output ok` returned ok immediately.

.EXAMPLE
powershell -File apply-navigation.ps1 `
  -Deck   C:\Users\wyusu\Downloads\AICHI_v18.pptx `
  -Module C:\Users\wyusu\Downloads\SciTeXNavigation.bas `
  -Output C:\Users\wyusu\Downloads\AICHI_v18_nav.pptm
#>
param(
    [Parameter(Mandatory = $true)][string]$Deck,
    [Parameter(Mandatory = $true)][string]$ModulePath,
    [Parameter(Mandatory = $true)][string]$Output
)

$ErrorActionPreference = "Stop"

foreach ($required in @($Deck, $ModulePath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Not found: $required"
    }
}
if (-not $Output.EndsWith(".pptm")) {
    throw "Output must be .pptm; a deck carrying a macro cannot be saved as .pptx."
}

$securityPath = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
if (-not (Test-Path $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
$securityItem = Get-ItemProperty -Path $securityPath -ErrorAction SilentlyContinue
$vbomExisted = ($null -ne $securityItem -and
                $securityItem.PSObject.Properties.Name -contains "AccessVBOM")
$vbomOriginal = $(if ($vbomExisted) { $securityItem.AccessVBOM } else { 0 })

$app = $null
$startedByUs = $false
$presentation = $null

try {
    Copy-Item -LiteralPath $Deck -Destination $Output -Force
    New-ItemProperty -Path $securityPath -Name "AccessVBOM" -PropertyType DWord `
        -Value 1 -Force | Out-Null

    # Always our own instance. See the note above about Quit().
    $app = New-Object -ComObject PowerPoint.Application
    $startedByUs = $true
    Write-Output "opened our own PowerPoint (existing windows untouched)"

    $presentation = $app.Presentations.Open($Output, $false, $false, $false)

    # A stale copy of the module would shadow the one we are testing.
    for ($i = $presentation.VBProject.VBComponents.Count; $i -ge 1; $i--) {
        $component = $presentation.VBProject.VBComponents.Item($i)
        if ($component.Name -eq "SciTeXNavigation") {
            $presentation.VBProject.VBComponents.Remove($component)
        }
    }
    $presentation.VBProject.VBComponents.Import($ModulePath) | Out-Null
    Write-Output "imported $(Split-Path -Leaf $ModulePath)"

    $app.Run("RunSciTeXNavigation") | Out-Null
    Write-Output "RunSciTeXNavigation returned"

    # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
    $presentation.SaveAs($Output, 25)
    Write-Output "saved $Output"
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Saved = $true; $presentation.Close() } catch { }
    }
    if ($startedByUs -and $null -ne $app) {
        try { $app.Quit() } catch { }
    }
    $restore = Join-Path $PSScriptRoot "restore-access-vbom.ps1"
    & $restore -SecurityPath $securityPath `
               -OriginalExisted $(if ($vbomExisted) { 1 } else { 0 }) `
               -OriginalValue $vbomOriginal
}
