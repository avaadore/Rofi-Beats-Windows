# SPDX-License-Identifier: GPL-3.0-only
# Rofi Beats - Windows launcher
[CmdletBinding()]
param(
    [switch]$ResetProfile
)

$target = Join-Path $PSScriptRoot "windows/rofi-beats-windows.ps1"
$stations = Join-Path $PSScriptRoot "windows/stations.json"
if (-not (Test-Path -Path $target)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "This launcher needs the full release package.`n`nPlease download and extract the .zip release so the 'windows' folder stays next to this file.",
        "Rofi Beats - Windows",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    throw "Windows script not found: $target"
}

if (-not (Test-Path -Path $stations)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "stations.json is missing.`n`nPlease download and extract the full .zip release instead of running the standalone launcher file.",
        "Rofi Beats - Windows",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    throw "Station catalog not found: $stations"
}

$forwardArgs = @()
if ($ResetProfile) {
    $forwardArgs += "-ResetProfile"
}

# WinForms requires STA. Relaunch with Windows PowerShell STA when needed.
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
    $windowsPs = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $hostExe = if (Test-Path -Path $windowsPs) { $windowsPs } else { (Get-Process -Id $PID).Path }

    $relaunchArgs = @(
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-STA",
        "-File", "`"$PSCommandPath`""
    )
    $relaunchArgs += $forwardArgs

    Start-Process -FilePath $hostExe -ArgumentList $relaunchArgs | Out-Null
    return
}

if ($forwardArgs.Count -gt 0) {
    & $target @forwardArgs
} else {
    & $target
}
