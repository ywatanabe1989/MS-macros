param(
    [Parameter(Mandatory = $true)][string]$SecurityPath,
    [Parameter(Mandatory = $true)][ValidateSet(0, 1)][int]$OriginalExisted,
    [Parameter(Mandatory = $true)][int]$OriginalValue
)

$ErrorActionPreference = "Stop"

for ($attempt = 0; $attempt -lt 30; $attempt++) {
    if ($null -eq (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)) {
        break
    }
    Start-Sleep -Milliseconds 500
}

# Office can flush Trust Center values shortly after POWERPNT exits.
Start-Sleep -Milliseconds 1500

function Restore-Value {
    if ($OriginalExisted -eq 1) {
        New-ItemProperty -Path $SecurityPath -Name "AccessVBOM" -PropertyType DWord -Value $OriginalValue -Force | Out-Null
    }
    else {
        Remove-ItemProperty -Path $SecurityPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
    }
}

Restore-Value
Start-Sleep -Milliseconds 500
Restore-Value

$item = Get-ItemProperty -Path $SecurityPath -ErrorAction SilentlyContinue
$exists = ($null -ne $item -and $item.PSObject.Properties.Name -contains "AccessVBOM")
if ($exists -ne ($OriginalExisted -eq 1)) {
    throw "Delayed AccessVBOM restoration failed."
}
if ($OriginalExisted -eq 1 -and $item.AccessVBOM -ne $OriginalValue) {
    throw "Delayed AccessVBOM value restoration failed."
}

Write-Output "AccessVBOM restoration verified"
