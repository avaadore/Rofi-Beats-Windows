# Security Policy

## Supported Scope
This repository currently supports the Windows app files:
- `rofi-beats-windows.cmd`
- `rofi-beats-windows.ps1`
- `windows/rofi-beats-windows.ps1`
- `windows/stations.json`

## Reporting a Vulnerability
Please open a private security report if possible, or open an issue without
publishing exploit details.

Include:
- reproduction steps
- impacted file/function
- expected vs actual behavior
- PowerShell version and Windows version

## Security Notes
- Stream URLs are loaded from local `stations.json` and validated as
  `http`/`https`.
- Automatic player install via `winget` is optional and user-confirmed.
- The app does not include telemetry or background network beacons beyond
  opening/playing selected station streams.
- Tray volume control uses Windows Core Audio session APIs and is intended to
  affect only the app session, not the system master volume.
- While a station is playing, the app may periodically read ICY metadata from
  the same stream URL to show bitrate and current song title in the tray.
