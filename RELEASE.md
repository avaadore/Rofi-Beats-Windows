# Release Flow

This repository already has CI and release workflows. Use this flow for normal patch/minor releases.

## 1) Validate locally

PowerShell syntax:

```powershell
[void][scriptblock]::Create((Get-Content -Path ".\rofi-beats-windows.ps1" -Raw))
[void][scriptblock]::Create((Get-Content -Path ".\windows\rofi-beats-windows.ps1" -Raw))
```

Station catalog:

```powershell
$stations = Get-Content -Path ".\windows\stations.json" -Raw | ConvertFrom-Json
if (-not $stations -or @($stations).Count -eq 0) { throw "stations.json is empty." }
```

Manual smoke test on Windows:
- tray icon appears correctly
- station playback starts
- setup wizard still applies a recommendation
- tray `Session volume` slider changes only app volume
- mute button next to the slider works
- stop and exit both shut playback down cleanly

## 2) Commit the release-ready changes

```powershell
git add .
git commit -m "Describe the release change"
```

## 3) Push `main`

```powershell
git push origin main
```

## 4) Create and push a version tag

Example:

```powershell
git tag -a v1.0.9 -m "Rofi Beats - Windows v1.0.9"
git push origin v1.0.9
```

## 5) Verify GitHub Actions

- `Windows CI` should pass on `main`
- `Windows Release` should run for the new tag
- the release should publish a `.zip` package asset

## 6) Add release notes

Keep release notes short and user-facing:
- what changed
- any bug fixes worth calling out
- any download/run guidance if behavior changed
