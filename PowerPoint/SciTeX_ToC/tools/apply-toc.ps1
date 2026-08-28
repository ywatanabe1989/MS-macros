<#
.SYNOPSIS
Apply SciTeX_ToC to a real deck and save the result beside it.

.DESCRIPTION
build-and-test.ps1 builds a synthetic sandbox to exercise the macro. This runs
it against a deck someone actually presents.

Four things are enforced here rather than left to the caller, because each one
has already cost something:

  NEVER QUIT POWERPOINT.  Not "quit only the instance we started" -- there is
  no such thing.  PowerPoint's COM server is a SINGLETON per user session, so
  New-Object PowerPoint.Application ATTACHES to a running instance rather than
  creating a second one, and a script cannot distinguish the two cases.  A
  startedByUs flag therefore records a belief, not a fact, and acting on it is
  how this closed the operator's open presentations twice (2026-08-27, and
  again 2026-08-28 through exactly that flag).  Close the presentation this
  script opened; leave the application alone.  A stray window costs one click,
  a Quit() costs unsaved work.

  OPEN THE SOURCE, SAVE AS THE OUTPUT.  Copying a .pptx to a .pptm name first
  fails: PowerPoint checks the container format against the extension and
  refuses -- "PowerPoint can't open this file because its file extension has
  changed".  The source is never saved, so it is left exactly as found.

  OPEN WITH A WINDOW.  A windowless presentation is never ActivePresentation,
  which is how the macro used to find its target: it broke into the VBE and sat
  there, invisible to this script, and the run simply never returned.

  NAME THE TARGET EXPLICITLY.  The macro is invoked as RefreshToCIn
  with the presentation as an argument, so "which deck" is never inferred.

  AccessVBOM is recorded, opened, and put back -- from the finally block, so a
  crash mid-run still closes it.

.PARAMETER Deck
The .pptx/.pptm to lay out.  Never modified.

.PARAMETER ModulePath
The exported SciTeX_ToC.bas to import.

.PARAMETER Output
Where to write the laid-out deck.  Must be .pptm: a deck carrying a macro
cannot be saved as .pptx.

.NOTES
LAUNCHING FROM WSL.  Calling powershell.exe directly across the WSL interop
socket fails intermittently -- it times out with

    WSL ERROR: UtilAcceptVsock:273: accept4 failed 110

and writes NOTHING but that line, so the run looks like a script that produced
no output rather than one that never started.  cmd.exe crosses the same socket
fine, so go through it:

    cd /mnt/c/Users/<user>          # cmd cannot start in a \\wsl.localhost path
    cmd.exe /c "powershell -NoProfile -ExecutionPolicy Bypass -File C:\...\apply-navigation.ps1 ..."

.EXAMPLE
powershell -File apply-navigation.ps1 `
  -Deck   C:\Users\wyusu\Downloads\AICHI_v18.pptx `
  -ModulePath C:\Users\wyusu\Template\SciTeX_ToC.bas `
  -Output C:\Users\wyusu\Downloads\AICHI_v18_nav.pptm
#>
param(
    [Parameter(Mandatory = $true)][string]$Deck,
    [Parameter(Mandatory = $true)][string]$ModulePath,
    [Parameter(Mandatory = $true)][string]$Output
)

$ErrorActionPreference = "Stop"

foreach ($required in @($Deck, $ModulePath)) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Not found: $required" }
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

$failure = Join-Path (Split-Path -Parent $Deck) "SciTeX_ToC.failure.txt"
Remove-Item -LiteralPath $failure -ErrorAction SilentlyContinue

$app = $null
$presentation = $null

try {
    New-ItemProperty -Path $securityPath -Name "AccessVBOM" -PropertyType DWord `
        -Value 1 -Force | Out-Null

    $app = New-Object -ComObject PowerPoint.Application
    Write-Output ("attached to PowerPoint (" + $app.Presentations.Count + " already open)")

    $presentation = $app.Presentations.Open($Deck, $false, $false, $true)
    Write-Output ("opened " + (Split-Path -Leaf $Deck))

    # A stale copy of the module would shadow the one under test.
    for ($i = $presentation.VBProject.VBComponents.Count; $i -ge 1; $i--) {
        $component = $presentation.VBProject.VBComponents.Item($i)
        if ($component.Name -eq "SciTeX_ToC") {
            $presentation.VBProject.VBComponents.Remove($component)
        }
    }
    $presentation.VBProject.VBComponents.Import($ModulePath) | Out-Null
    Write-Output ("imported " + (Split-Path -Leaf $ModulePath))

    # NOT $app.Run("..."): PowerShell cannot bind that, because Run takes the
    # macro name plus a parameter array and PowerShell will not supply the
    # optional half -- "Cannot find an overload for Run and the argument
    # count: 1". InvokeMember hands COM the argument array directly.
    # Qualify the macro with its presentation.  Application.Run resolves a BARE
    # name against every open presentation, so an older copy of this module sitting
    # in another open deck (the distributable template, say) can win the lookup and
    # run instead -- silently, with no error, producing output from stale code.
    $qualified = "'" + $presentation.Name + "'!SciTeX_ToC.RefreshToCIn"
    Write-Output ("invoking " + $qualified)
    $app.GetType().InvokeMember(
        "Run",
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null, $app, @($qualified, $presentation)) | Out-Null
    Write-Output "RefreshToCIn returned"

    # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
    $presentation.SaveAs($Output, 25)
    Write-Output ("saved " + $Output)
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Saved = $true; $presentation.Close() } catch { }
    }
    # Deliberately no $app.Quit(). See the note at the top.
    $restore = Join-Path $PSScriptRoot "restore-access-vbom.ps1"
    & $restore -SecurityPath $securityPath `
               -OriginalExisted $(if ($vbomExisted) { 1 } else { 0 }) `
               -OriginalValue $vbomOriginal
    if (Test-Path -LiteralPath $failure) {
        Write-Output "--- the macro reported a failure ---"
        Get-Content -LiteralPath $failure | Write-Output
    }
}
