param(
    [Parameter(Mandatory = $false)]
    [string]$AddInPath = ".\OnionPPT.ppam"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-AbsolutePath([string]$PathValue) {
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }
    return (Join-Path -Path (Get-Location) -ChildPath $PathValue)
}

$sourcePath = Resolve-AbsolutePath $AddInPath
if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Add-in file not found: $sourcePath"
}

if ([System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant() -ne ".ppam") {
    throw "Expected a .ppam file. Received: $sourcePath"
}

$addInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path -LiteralPath $addInsDir)) {
    New-Item -ItemType Directory -Path $addInsDir | Out-Null
}

$targetPath = Join-Path $addInsDir ([System.IO.Path]::GetFileName($sourcePath))
Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force

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

    $powerPoint.DisplayAlerts = 1

    $existing = $null
    foreach ($item in $powerPoint.AddIns) {
        if ($item.FullName -ieq $targetPath) {
            $existing = $item
            break
        }
    }

    if ($null -eq $existing) {
        $existing = $powerPoint.AddIns.Add($targetPath, $true)
    }

    $existing.Loaded = $true

    Write-Host "Installed and loaded successfully:"
    Write-Host "  $targetPath"
    Write-Host ""
    Write-Host "Open PowerPoint > Onion Tools tab to verify."
}
finally {
    if ($createdInstance -and $null -ne $powerPoint) {
        $powerPoint.Quit()
    }
}
