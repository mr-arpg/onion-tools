param(
    [Parameter(Mandatory = $false)]
    [string]$AddInFileName = "OnionPPT.ppam"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$targetPath = Join-Path (Join-Path $env:APPDATA "Microsoft\AddIns") $AddInFileName

$powerPoint = $null
$createdInstance = $false

try {
    try {
        $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    }
    catch {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $createdInstance = $true
    }

    $found = $false
    foreach ($item in $powerPoint.AddIns) {
        if ($item.FullName -ieq $targetPath) {
            $item.Loaded = $false
            $found = $true
            break
        }
    }

    if (Test-Path -LiteralPath $targetPath) {
        Remove-Item -LiteralPath $targetPath -Force
        Write-Host "Removed add-in file: $targetPath"
    }
    elseif (-not $found) {
        Write-Host "Add-in was not installed at: $targetPath"
    }
}
finally {
    if ($createdInstance -and $null -ne $powerPoint) {
        $powerPoint.Quit()
    }
}
