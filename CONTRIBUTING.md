# Contributing

Thanks for helping improve Rofi Beats - Windows.

## Development Scope
- Windows-only repository
- Main app script: `windows/rofi-beats-windows.ps1`
- Station catalog: `windows/stations.json`

## Local Validation
Run these before opening a PR:

```powershell
[void][scriptblock]::Create((Get-Content -Raw .\rofi-beats-windows.ps1))
[void][scriptblock]::Create((Get-Content -Raw .\windows\rofi-beats-windows.ps1))
```

```powershell
$stations = Get-Content -Raw .\windows\stations.json | ConvertFrom-Json
if (-not $stations -or @($stations).Count -eq 0) { throw "stations.json is empty" }
```

## Station Contribution Rules
- `id`, `name`, `url` are required.
- `url` must be absolute `http` or `https`.
- Keep `id` unique.
- Prefer stable public streams.
- Add useful tags in `moods` and `genres`.

## Pull Requests
- Keep changes focused.
- Explain user-visible behavior changes.
- Include test/repro notes when fixing bugs.
