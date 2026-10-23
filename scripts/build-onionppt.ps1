# Builds OnionPPT.ppam from src/ and registers it.
# PowerPoint writes the VBA project. After it quits, patch_ribbon.py injects the ribbon.
# Requires PowerPoint and "Trust access to the VBA project object model"
# (this script turns that setting on for the installed Office version).

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$basSrc = Join-Path $repoRoot "src\vba\OnionSkinTools.bas"
$clsSrc = Join-Path $repoRoot "src\vba\OnionSkinEventBridge.cls"
$patchScript = Join-Path $repoRoot "scripts\patch_ribbon.py"
$addInDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$ppamPath = Join-Path $addInDir "OnionPPT.ppam"

$ppAlertsNone = 1
$ppSaveAsMacroEnabled = 25
$ppSaveAsOpenXmlAddin = 30
$msoTrue = -1

function Enable-VbaAccess {
    $versions = @("16.0", "15.0", "14.0")
    $found = $false
    foreach ($version in $versions) {
        $key = "HKCU:\Software\Microsoft\Office\$version\PowerPoint\Security"
        if (Test-Path $key) {
            New-ItemProperty -Path $key -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
            $found = $true
        }
    }
    if (-not $found) {
        $key = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
        New-Item -Path $key -Force | Out-Null
        New-ItemProperty -Path $key -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
    }
}

function Close-PowerPointIfIdle {
    $procs = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
    if ($procs.Count -eq 0) { return }

    $ppt = $null
    try {
        $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        throw "PowerPoint is open and this script cannot attach to it. Close it, then run the script again."
    }

    foreach ($pres in @($ppt.Presentations)) {
        if (-not $pres.Saved) {
            throw "PowerPoint has unsaved work. Save or close those files, then run the script again."
        }
    }

    foreach ($addIn in @($ppt.AddIns)) {
        if ($addIn.Name -eq "OnionPPT.ppam" -or $addIn.FullName -eq $ppamPath) {
            $addIn.Loaded = $false
        }
    }

    $ppt.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)

    foreach ($proc in $procs) {
        if (-not $proc.HasExited) { $proc.WaitForExit(20000) | Out-Null }
    }
    $still = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
    if ($still.Count -gt 0) {
        $still | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
}

function New-ImportableBas {
    $tempBas = Join-Path $env:TEMP "OnionSkinTools.import.bas"
    $text = [System.IO.File]::ReadAllText($basSrc)
    if ($text -notmatch "Attribute VB_Name") {
        $text = "Attribute VB_Name = `"OnionSkinTools`"`r`n" + $text
    }
    [System.IO.File]::WriteAllText($tempBas, $text, [System.Text.Encoding]::Default)
    return $tempBas
}

function Wait-FileReady([string]$path) {
    for ($i = 0; $i -lt 30; $i++) {
        try {
            $stream = [System.IO.File]::Open($path, "Open", "ReadWrite", "None")
            $stream.Close()
            return
        } catch {
            Start-Sleep -Seconds 1
        }
    }
    throw "Timed out waiting for $path to be unlocked."
}

function Invoke-RibbonPatch {
    $python = $null
    foreach ($candidate in @("python", "py")) {
        if (Get-Command $candidate -ErrorAction SilentlyContinue) {
            $python = $candidate
            break
        }
    }
    if (-not $python) { throw "Python was not found on PATH." }

    if ($python -eq "py") {
        & py -3 $patchScript $ppamPath
    } else {
        & python $patchScript $ppamPath
    }
    if ($LASTEXITCODE -ne 0) { throw "Ribbon patch failed." }
}

if (-not (Test-Path $basSrc)) { throw "Missing $basSrc" }
if (-not (Test-Path $clsSrc)) { throw "Missing $clsSrc" }
New-Item -ItemType Directory -Force -Path $addInDir | Out-Null

Enable-VbaAccess
Close-PowerPointIfIdle

$tempBas = New-ImportableBas
$tempPptm = Join-Path $env:TEMP ("OnionPPT-build-" + [guid]::NewGuid().ToString("N") + ".pptm")
$ppt = $null
$pres = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue
    $ppt.DisplayAlerts = $ppAlertsNone

    foreach ($addIn in @($ppt.AddIns)) {
        if ($addIn.Name -eq "OnionPPT.ppam" -or $addIn.FullName -eq $ppamPath) {
            $addIn.Loaded = $false
        }
    }

    $pres = $ppt.Presentations.Add($msoTrue)
    $pres.SaveAs($tempPptm, $ppSaveAsMacroEnabled) | Out-Null
    if (-not (Test-Path -LiteralPath $tempPptm)) {
        throw "PowerPoint did not save the temporary macro-enabled presentation."
    }

    $components = $pres.VBProject.VBComponents
    foreach ($name in @("OnionSkinTools", "OnionSkinEventBridge")) {
        foreach ($component in @($components)) {
            if ($component.Name -eq $name) {
                $components.Remove($component)
            }
        }
    }
    [void]$components.Import($tempBas)
    [void]$components.Import($clsSrc)

    if (Test-Path -LiteralPath $ppamPath) {
        Remove-Item -LiteralPath $ppamPath -Force
    }
    $pres.SaveAs($ppamPath, $ppSaveAsOpenXmlAddin) | Out-Null
    if (-not (Test-Path -LiteralPath $ppamPath)) {
        throw "PowerPoint did not save $ppamPath"
    }

    $pres.Saved = $true
    $pres.Close()
    $pres = $null
    $ppt.Quit()
} finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
    }
    if ($ppt -ne $null) {
        try { $ppt.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path $tempBas) { Remove-Item -LiteralPath $tempBas -Force -ErrorAction SilentlyContinue }
    if (Test-Path $tempPptm) { Remove-Item -LiteralPath $tempPptm -Force -ErrorAction SilentlyContinue }
}

$leftover = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
if ($leftover.Count -gt 0) {
    $leftover | Stop-Process -Force
    Start-Sleep -Seconds 1
}

Wait-FileReady $ppamPath
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($ppamPath)
$hasVba = @($zip.Entries | Where-Object { $_.FullName -eq "ppt/vbaProject.bin" }).Count -gt 0
$zip.Dispose()
if (-not $hasVba) {
    throw "The saved file has no VBA project. PowerPoint did not write an add-in."
}
Invoke-RibbonPatch

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue
    $ppt.DisplayAlerts = $ppAlertsNone
    $alreadyListed = $false
    foreach ($addIn in @($ppt.AddIns)) {
        if ($addIn.FullName -eq $ppamPath -or $addIn.Name -eq "OnionPPT.ppam") {
            $addIn.Loaded = $true
            $alreadyListed = $true
        }
    }
    if (-not $alreadyListed) {
        $added = $ppt.AddIns.Add($ppamPath)
        $added.Loaded = $true
    }
} finally {
    if ($ppt -ne $null) {
        try { $ppt.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Built and registered: $ppamPath"
Write-Output "Open PowerPoint. The Onion Tools tab should be there."
