# SPDX-License-Identifier: GPL-3.0-only
# Rofi Beats - Windows
# Copyright (C) contributors
# Windows rewrite inspired by Carbon-Bl4ck/Rofi-Beats
[CmdletBinding()]
param(
    [switch]$ResetProfile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$hotkeyTypeSource = @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class RofiHotkeyWindow : Form
{
    public const int WM_HOTKEY = 0x0312;

    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);

    public event EventHandler HotkeyPressed;

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && HotkeyPressed != null)
        {
            HotkeyPressed(this, EventArgs.Empty);
        }

        base.WndProc(ref m);
    }
}
"@

$hotkeyRefSet = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)

function Add-HotkeyAssemblyRefs {
    param(
        [Parameter(Mandatory = $true)]
        [System.Reflection.Assembly]$Assembly
    )

    if (-not [string]::IsNullOrWhiteSpace($Assembly.Location)) {
        [void]$hotkeyRefSet.Add($Assembly.Location)
    }

    foreach ($refName in $Assembly.GetReferencedAssemblies()) {
        try {
            $refAsm = [System.Reflection.Assembly]::Load($refName)
            if ($refAsm -and -not [string]::IsNullOrWhiteSpace($refAsm.Location)) {
                [void]$hotkeyRefSet.Add($refAsm.Location)
            }
        } catch {
        }
    }
}

Add-HotkeyAssemblyRefs -Assembly ([System.Windows.Forms.Form].Assembly)
Add-HotkeyAssemblyRefs -Assembly ([System.Object].Assembly)
Add-HotkeyAssemblyRefs -Assembly ([System.Runtime.InteropServices.DllImportAttribute].Assembly)
$hotkeyRefs = @($hotkeyRefSet)

if (-not ("RofiHotkeyWindow" -as [type])) {
    if ($hotkeyRefs.Count -gt 0) {
        Add-Type -TypeDefinition $hotkeyTypeSource -Language CSharp -ReferencedAssemblies $hotkeyRefs
    } else {
        Add-Type -TypeDefinition $hotkeyTypeSource -Language CSharp
    }
}

$audioSessionTypeSource = @"
using System;
using System.Runtime.InteropServices;

public enum EDataFlow
{
    eRender = 0,
    eCapture = 1,
    eAll = 2
}

public enum ERole
{
    eConsole = 0,
    eMultimedia = 1,
    eCommunications = 2
}

[Flags]
public enum CLSCTX : uint
{
    INPROC_SERVER = 0x1,
    INPROC_HANDLER = 0x2,
    LOCAL_SERVER = 0x4,
    REMOTE_SERVER = 0x10,
    ALL = INPROC_SERVER | INPROC_HANDLER | LOCAL_SERVER | REMOTE_SERVER
}

[ComImport]
[Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
public class MMDeviceEnumeratorComObject
{
}

[ComImport]
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDeviceEnumerator
{
    int EnumAudioEndpoints(EDataFlow dataFlow, int dwStateMask, out object ppDevices);
    [PreserveSig]
    int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
    int GetDevice(string pwstrId, out IMMDevice ppDevice);
    int RegisterEndpointNotificationCallback(IntPtr pClient);
    int UnregisterEndpointNotificationCallback(IntPtr pClient);
}

[ComImport]
[Guid("D666063F-1587-4E43-81F1-B948E807363F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDevice
{
    [PreserveSig]
    int Activate(ref Guid iid, CLSCTX dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
}

[ComImport]
[Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioSessionManager2
{
    int GetAudioSessionControl(IntPtr AudioSessionGuid, uint StreamFlags, out IntPtr SessionControl);
    int GetSimpleAudioVolume(IntPtr AudioSessionGuid, uint StreamFlags, out IntPtr AudioVolume);
    [PreserveSig]
    int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);
}

[ComImport]
[Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioSessionEnumerator
{
    [PreserveSig]
    int GetCount(out int SessionCount);
    [PreserveSig]
    int GetSession(int SessionCount, out IAudioSessionControl Session);
}

[ComImport]
[Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioSessionControl
{
    int GetState(out int pRetVal);
    int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
    int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
    int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
    int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
    int GetGroupingParam(out Guid pRetVal);
    int SetGroupingParam(ref Guid Override, ref Guid EventContext);
    int RegisterAudioSessionNotification(IntPtr NewNotifications);
    int UnregisterAudioSessionNotification(IntPtr NewNotifications);
}

[ComImport]
[Guid("BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAudioSessionControl2 : IAudioSessionControl
{
    new int GetState(out int pRetVal);
    new int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
    new int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
    new int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
    new int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
    new int GetGroupingParam(out Guid pRetVal);
    new int SetGroupingParam(ref Guid Override, ref Guid EventContext);
    new int RegisterAudioSessionNotification(IntPtr NewNotifications);
    new int UnregisterAudioSessionNotification(IntPtr NewNotifications);
    int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
    int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
    [PreserveSig]
    int GetProcessId(out uint pRetVal);
    int IsSystemSoundsSession();
    int SetDuckingPreference(bool optOut);
}

[ComImport]
[Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISimpleAudioVolume
{
    [PreserveSig]
    int SetMasterVolume(float fLevel, ref Guid EventContext);
    [PreserveSig]
    int GetMasterVolume(out float pfLevel);
    [PreserveSig]
    int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid EventContext);
    [PreserveSig]
    int GetMute([MarshalAs(UnmanagedType.Bool)] out bool pbMute);
}

public static class AudioSessionHelper
{
    private delegate void SessionAction(ISimpleAudioVolume volume);

    public static bool TryGetSnapshot(int processId, out float volumeScalar, out bool muted)
    {
        return VisitSessions(processId, null, out volumeScalar, out muted);
    }

    public static bool TrySetState(int processId, float volumeScalar, bool muted, out float appliedVolumeScalar, out bool appliedMuted)
    {
        float targetVolume = ClampScalar(volumeScalar);
        return VisitSessions(
            processId,
            delegate (ISimpleAudioVolume volume)
            {
                Guid context = Guid.Empty;
                Marshal.ThrowExceptionForHR(volume.SetMasterVolume(targetVolume, ref context));
                Marshal.ThrowExceptionForHR(volume.SetMute(muted, ref context));
            },
            out appliedVolumeScalar,
            out appliedMuted);
    }

    private static bool VisitSessions(int processId, SessionAction action, out float volumeScalar, out bool muted)
    {
        volumeScalar = 1.0f;
        muted = false;
        if (processId <= 0)
        {
            return false;
        }

        IMMDeviceEnumerator deviceEnumerator = null;
        IMMDevice device = null;
        object managerObject = null;
        IAudioSessionManager2 sessionManager = null;
        IAudioSessionEnumerator sessionEnumerator = null;

        try
        {
            deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
            Marshal.ThrowExceptionForHR(deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device));

            Guid sessionManagerGuid = typeof(IAudioSessionManager2).GUID;
            Marshal.ThrowExceptionForHR(device.Activate(ref sessionManagerGuid, CLSCTX.ALL, IntPtr.Zero, out managerObject));
            sessionManager = (IAudioSessionManager2)managerObject;

            Marshal.ThrowExceptionForHR(sessionManager.GetSessionEnumerator(out sessionEnumerator));

            int sessionCount = 0;
            Marshal.ThrowExceptionForHR(sessionEnumerator.GetCount(out sessionCount));

            bool matched = false;
            for (int i = 0; i < sessionCount; i++)
            {
                IAudioSessionControl sessionControl = null;
                try
                {
                    Marshal.ThrowExceptionForHR(sessionEnumerator.GetSession(i, out sessionControl));
                    IAudioSessionControl2 sessionControl2 = sessionControl as IAudioSessionControl2;
                    if (sessionControl2 == null)
                    {
                        continue;
                    }

                    uint sessionProcessId = 0;
                    Marshal.ThrowExceptionForHR(sessionControl2.GetProcessId(out sessionProcessId));
                    if (sessionProcessId != processId)
                    {
                        continue;
                    }

                    ISimpleAudioVolume simpleVolume = sessionControl as ISimpleAudioVolume;
                    if (simpleVolume == null)
                    {
                        continue;
                    }

                    float currentVolume = 1.0f;
                    bool currentMuted = false;
                    Marshal.ThrowExceptionForHR(simpleVolume.GetMasterVolume(out currentVolume));
                    Marshal.ThrowExceptionForHR(simpleVolume.GetMute(out currentMuted));

                    volumeScalar = ClampScalar(currentVolume);
                    muted = currentMuted;

                    if (action != null)
                    {
                        action(simpleVolume);
                        Marshal.ThrowExceptionForHR(simpleVolume.GetMasterVolume(out currentVolume));
                        Marshal.ThrowExceptionForHR(simpleVolume.GetMute(out currentMuted));
                        volumeScalar = ClampScalar(currentVolume);
                        muted = currentMuted;
                    }

                    matched = true;
                }
                finally
                {
                    ReleaseComObject(sessionControl);
                }
            }

            return matched;
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(sessionEnumerator);
            ReleaseComObject(sessionManager);
            ReleaseComObject(device);
            ReleaseComObject(deviceEnumerator);
        }
    }

    private static float ClampScalar(float value)
    {
        if (value < 0.0f)
        {
            return 0.0f;
        }

        if (value > 1.0f)
        {
            return 1.0f;
        }

        return value;
    }

    private static void ReleaseComObject(object value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            try
            {
                Marshal.ReleaseComObject(value);
            }
            catch
            {
            }
        }
    }
}
"@

if (-not ("AudioSessionHelper" -as [type])) {
    Add-Type -TypeDefinition $audioSessionTypeSource -Language CSharp
}

$script:AppName = "Rofi Beats - Windows"
$script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:StationsPath = Join-Path $script:ScriptRoot "stations.json"
$script:DataRoot = Join-Path $env:APPDATA "RofiBeats"
$script:ProfilePath = Join-Path $script:DataRoot "profile.json"
$script:StatePath = Join-Path $script:DataRoot "state.json"

$script:allStations = @()
$script:profile = $null
$script:playerProcess = $null
$script:currentStation = $null
$script:notifyIcon = $null
$script:hotkeyWindow = $null
$script:isExiting = $false
$script:mpvPath = $null
$script:playerBackend = $null
$script:hotkeyId = 9742
$script:playerMarker = "rofi-beats-radio"

$script:MoodMap = [ordered]@{
    all       = "All moods"
    focus     = "Focus"
    chill     = "Chill"
    energy    = "Energy"
    nostalgia = "Nostalgia"
    deep      = "Deep"
    study     = "Study"
    social    = "Social"
}

$script:GenreMap = [ordered]@{
    all         = "All genres"
    lofi        = "Lofi"
    chillout    = "Chillout"
    jazz        = "Jazz"
    rock        = "Rock"
    metal       = "Metal"
    hiphop      = "Hip-Hop"
    house       = "House"
    trance      = "Trance"
    techno      = "Techno"
    reggae      = "Reggae"
    blues       = "Blues"
    ambient     = "Ambient"
    synthwave   = "Synthwave"
    pop         = "Pop"
    turkce      = "Turkish"
    folk        = "Folk/Turku"
    game        = "Game/VGM"
    classical   = "Classical"
    news        = "News/Talk"
    electronic  = "Electronic"
}

$script:GenreProfiles = [ordered]@{
    all        = @()
    lofi       = @("lofi", "chillhop", "chillout")
    chillout   = @("chillout", "ambient", "lounge", "electronic")
    jazz       = @("jazz", "smoothjazz", "soul", "blues")
    rock       = @("rock", "metal", "pop")
    metal      = @("metal", "rock")
    hiphop     = @("hiphop", "rap", "rnb", "chillhop")
    house      = @("house", "electronic", "techno", "trance")
    trance     = @("trance", "electronic", "house", "techno")
    techno     = @("techno", "electronic", "house", "trance")
    reggae     = @("reggae", "ska", "soul", "chillout")
    blues      = @("blues", "jazz", "soul", "smoothjazz")
    ambient    = @("ambient", "chillout", "sleep", "vaporwave")
    synthwave  = @("synthwave", "electronic", "vaporwave")
    pop        = @("pop", "turkce", "electronic", "rock")
    turkce     = @("turkce", "pop", "folk", "nostalji", "arabesk")
    folk       = @("folk", "turku", "turkce")
    game       = @("game", "vgm", "chiptune", "synthwave", "electronic")
    classical  = @("classical", "public-radio", "ambient", "jazz")
    news       = @("news", "talk", "public-radio")
    electronic = @("electronic", "house", "trance", "techno", "synthwave", "ambient", "chillout", "vaporwave")
}

$script:DiscoveryMap = [ordered]@{
    balanced = "Safe picks (recommended)"
    new      = "Discover new stations"
    surprise = "Surprise me"
}

$script:statusItem = $null
$script:currentItem = $null
$script:toggleItem = $null
$script:hotkeyHandler = $null
$script:songItem = $null
$script:bitrateItem = $null
$script:volumeItem = $null
$script:volumeUpItem = $null
$script:volumeDownItem = $null
$script:muteItem = $null
$script:trayIcon = $null
$script:metadataTimer = $null
$script:audioSessionTimer = $null
$script:streamInfo = [ordered]@{
    StreamTitle = $null
    BitrateKbps = $null
    UpdatedAt   = $null
}
$script:audioSessionState = [ordered]@{
    VolumePercent = 100
    IsMuted       = $false
    LastPid       = $null
    AppliedPid    = $null
    PendingVolumePercent = $null
    PendingMuted  = $null
    UpdatedAt     = $null
}

function Ensure-DataRoot {
    if (-not (Test-Path -Path $script:DataRoot)) {
        New-Item -Path $script:DataRoot -ItemType Directory -Force | Out-Null
    }
}

function Read-JsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        return $null
    }

    $raw = Get-Content -Path $Path -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    return $raw | ConvertFrom-Json
}

function Save-JsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [object]$Value
    )

    $Value | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
}

function New-DefaultProfile {
    return [ordered]@{
        version        = 1
        hasOnboarded   = $false
        preferredMood  = "focus"
        preferredGenre = "lofi"
        preferTurkish  = $false
        publicOnly     = $true
        discoveryMode  = "balanced"
        volume         = 35
        maxStartupVolume = 45
        sessionVolume  = 35
        sessionMuted   = $false
        hotkey         = "Ctrl+Alt+B"
        lastStationId  = $null
        preferredPlayer = "auto"
        customPlayerPath = $null
        customMpvPath  = $null
    }
}

function Save-Profile {
    Ensure-DataRoot
    Save-JsonFile -Path $script:ProfilePath -Value $script:profile
}

function Initialize-Profile {
    Ensure-DataRoot

    if ($ResetProfile -and (Test-Path -Path $script:ProfilePath)) {
        Remove-Item -Path $script:ProfilePath -Force
    }

    $defaultProfile = New-DefaultProfile
    $existing = Read-JsonFile -Path $script:ProfilePath
    $merged = [ordered]@{}

    foreach ($key in $defaultProfile.Keys) {
        $merged[$key] = $defaultProfile[$key]
    }

    $existingVersion = 1
    if ($null -ne $existing) {
        $existingVersion = [int](Get-ObjectPropertyValue -Object $existing -Name "version" -Default 1)
        foreach ($prop in $existing.PSObject.Properties) {
            $merged[$prop.Name] = $prop.Value
        }
    }

    if ($existingVersion -lt 2) {
        $merged["version"] = 2
        $preferredVolume = [Math]::Max([Math]::Min([int](Get-ObjectPropertyValue -Object ([pscustomobject]$merged) -Name "volume" -Default 35), 100), 0)
        $hasExplicitSessionVolume = ($null -ne $existing -and $existing.PSObject.Properties["sessionVolume"])
        $currentSessionVolume = [int](Get-ObjectPropertyValue -Object ([pscustomobject]$merged) -Name "sessionVolume" -Default 100)
        if (-not $hasExplicitSessionVolume -or $currentSessionVolume -eq 100) {
            $merged["sessionVolume"] = $preferredVolume
        }
    }

    $script:profile = [pscustomobject]$merged
    Save-Profile
    Reset-AudioSessionState
}

function To-Array {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return @($Value)
    }

    return @($Value)
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [object]$Default = $null
    )

    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($Name)) {
        return $Default
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) {
            return $Object[$Name]
        }
        return $Default
    }

    if ($Object.PSObject -and $Object.PSObject.Properties[$Name]) {
        return $Object.PSObject.Properties[$Name].Value
    }

    return $Default
}

function Get-GenreSelectionKeys {
    param(
        [string]$GenreKey = "all"
    )

    if ([string]::IsNullOrWhiteSpace($GenreKey) -or $GenreKey -eq "all") {
        return @()
    }

    if ($script:GenreProfiles.Contains($GenreKey)) {
        return @(To-Array -Value $script:GenreProfiles[$GenreKey])
    }

    return @($GenreKey)
}

function Test-StationMatchesGenreSelection {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Station,
        [string]$GenreKey = "all"
    )

    $selectionKeys = Get-GenreSelectionKeys -GenreKey $GenreKey
    if (@($selectionKeys).Count -eq 0) {
        return $true
    }

    $stationGenres = To-Array -Value $Station.genres
    foreach ($key in @($selectionKeys)) {
        if ($stationGenres -contains $key) {
            return $true
        }
    }

    return $false
}

function Test-IsSafeStreamUrl {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $false
    }

    try {
        $uri = [System.Uri]$Url
    } catch {
        return $false
    }

    if (-not $uri.IsAbsoluteUri) {
        return $false
    }

    return ($uri.Scheme -eq [System.Uri]::UriSchemeHttp -or $uri.Scheme -eq [System.Uri]::UriSchemeHttps)
}

function Get-BackendFromExecutable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $name = [System.IO.Path]::GetFileName($Path).ToLowerInvariant()
    switch -Regex ($name) {
        "^vlc\.exe$" { return "vlc" }
        "^ffplay\.exe$" { return "ffplay" }
        "^(mpv|mpvnet)\.exe$" { return "mpv" }
        default { return $null }
    }
}

function Resolve-BackendPath {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("mpv", "vlc", "ffplay")]
        [string]$Backend
    )

    $exe = "$Backend.exe"

    $fallbacks = @()
    switch ($Backend) {
        "mpv" {
            $fallbacks += @(
                (Join-Path $env:ProgramFiles "mpv\mpv.exe"),
                (Join-Path $env:ProgramFiles "mpv.net\mpv.exe"),
                (Join-Path $env:ProgramFiles "MPC-HC\mpv.exe"),
                (Join-Path $env:LOCALAPPDATA "Programs\mpv\mpv.exe"),
                (Join-Path $env:USERPROFILE "scoop\shims\mpv.exe"),
                (Join-Path $env:ProgramData "chocolatey\bin\mpv.exe"),
                (Join-Path (Split-Path -Parent $script:ScriptRoot) "mpv\mpv.exe"),
                (Join-Path $script:ScriptRoot "mpv.exe")
            )
            if ($env:ProgramW6432) {
                $fallbacks += (Join-Path $env:ProgramW6432 "mpv\mpv.exe")
                $fallbacks += (Join-Path $env:ProgramW6432 "mpv.net\mpv.exe")
            }
        }
        "vlc" {
            $fallbacks += @(
                (Join-Path $env:ProgramFiles "VideoLAN\VLC\vlc.exe"),
                (Join-Path $env:LOCALAPPDATA "Programs\VideoLAN\VLC\vlc.exe"),
                (Join-Path $env:USERPROFILE "scoop\apps\vlc\current\vlc.exe"),
                (Join-Path $env:ProgramData "chocolatey\bin\vlc.exe"),
                (Join-Path $script:ScriptRoot "vlc.exe")
            )
            if (${env:ProgramFiles(x86)}) {
                $fallbacks += (Join-Path ${env:ProgramFiles(x86)} "VideoLAN\VLC\vlc.exe")
            }
        }
        "ffplay" {
            $fallbacks += @(
                (Join-Path $env:ProgramFiles "ffmpeg\bin\ffplay.exe"),
                (Join-Path $env:LOCALAPPDATA "Programs\ffmpeg\bin\ffplay.exe"),
                (Join-Path $env:USERPROFILE "scoop\shims\ffplay.exe"),
                (Join-Path $env:ProgramData "chocolatey\bin\ffplay.exe"),
                (Join-Path $script:ScriptRoot "ffplay.exe")
            )
        }
    }

    foreach ($candidate in $fallbacks) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -Path $candidate)) {
            return $candidate
        }
    }

    $cmd = Get-Command $exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.CommandType -eq "Application" -and (Test-Path -Path $cmd.Source)) {
        return $cmd.Source
    }

    return $null
}

function Resolve-MpvPath {
    return Resolve-BackendPath -Backend "mpv"
}

function Resolve-PlayerPath {
    if ($script:profile) {
        if ($script:profile.PSObject.Properties["customPlayerPath"]) {
            $profilePath = [string]$script:profile.customPlayerPath
            if (-not [string]::IsNullOrWhiteSpace($profilePath) -and (Test-Path -Path $profilePath)) {
                $backend = Get-BackendFromExecutable -Path $profilePath
                if ($backend) {
                    return [pscustomobject]@{ Backend = $backend; Path = $profilePath }
                }
            }
        }

        if ($script:profile.PSObject.Properties["customMpvPath"]) {
            $legacyMpv = [string]$script:profile.customMpvPath
            if (-not [string]::IsNullOrWhiteSpace($legacyMpv) -and (Test-Path -Path $legacyMpv)) {
                return [pscustomobject]@{ Backend = "mpv"; Path = $legacyMpv }
            }
        }
    }

    $order = @("vlc", "ffplay", "mpv")
    if ($script:profile -and $script:profile.PSObject.Properties["preferredPlayer"]) {
        $preferred = [string]$script:profile.preferredPlayer
        if (-not [string]::IsNullOrWhiteSpace($preferred) -and $preferred -ne "auto" -and ($order -contains $preferred)) {
            $order = @($preferred) + @($order | Where-Object { $_ -ne $preferred })
        }
    }

    foreach ($backend in $order) {
        $path = Resolve-BackendPath -Backend $backend
        if ($path) {
            return [pscustomobject]@{ Backend = $backend; Path = $path }
        }
    }

    return $null
}

function Set-CustomPlayerPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        return $false
    }

    $backend = Get-BackendFromExecutable -Path $Path
    if (-not $backend) {
        return $false
    }

    if (-not $script:profile.PSObject.Properties["customPlayerPath"]) {
        $script:profile | Add-Member -NotePropertyName customPlayerPath -NotePropertyValue $null
    }
    if (-not $script:profile.PSObject.Properties["preferredPlayer"]) {
        $script:profile | Add-Member -NotePropertyName preferredPlayer -NotePropertyValue "auto"
    }
    if (-not $script:profile.PSObject.Properties["customMpvPath"]) {
        $script:profile | Add-Member -NotePropertyName customMpvPath -NotePropertyValue $null
    }

    $script:profile.customPlayerPath = $Path
    $script:profile.preferredPlayer = $backend
    if ($backend -eq "mpv") {
        $script:profile.customMpvPath = $Path
    }

    Save-Profile
    $script:playerBackend = $backend
    return $true
}

function Set-CustomMpvPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    return (Set-CustomPlayerPath -Path $Path)
}

function Select-MpvExecutable {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select player executable (mpv.exe, vlc.exe, ffplay.exe)"
    $dialog.Filter = "Supported players (mpv.exe;vlc.exe;ffplay.exe)|mpv.exe;vlc.exe;ffplay.exe|Executable files (*.exe)|*.exe"
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false

    $likely = @(
        (Join-Path $env:ProgramFiles "VideoLAN\VLC"),
        (Join-Path $env:ProgramFiles "mpv"),
        (Join-Path $env:ProgramFiles "ffmpeg\bin"),
        (Join-Path $env:LOCALAPPDATA "Programs")
    )

    foreach ($dir in $likely) {
        if (Test-Path -Path $dir) {
            $dialog.InitialDirectory = $dir
            break
        }
    }

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $picked = $dialog.FileName
    if ([string]::IsNullOrWhiteSpace($picked) -or -not (Test-Path -Path $picked)) {
        return $null
    }

    if (-not (Set-CustomPlayerPath -Path $picked)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Unsupported player. Please select mpv.exe, vlc.exe, or ffplay.exe.",
            $script:AppName,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $null
    }

    return $picked
}

function Try-InstallMpvWithWinget {
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        return $false
    }

    $candidateIds = @(
        "VideoLAN.VLC",
        "Mpv.net.Mpv",
        "shinchiro.mpv",
        "Gyan.FFmpeg"
    )

    foreach ($id in $candidateIds) {
        try {
            $args = @(
                "install",
                "--id", $id,
                "--exact",
                "--accept-source-agreements",
                "--accept-package-agreements",
                "--silent",
                "--scope", "user"
            )

            $proc = Start-Process -FilePath $winget.Source -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
            if ($proc.ExitCode -eq 0) {
                $resolved = Resolve-PlayerPath
                if ($resolved) {
                    return $true
                }
            }
        } catch {
        }
    }

    return $false
}

function Ensure-MpvAvailable {
    $resolved = Resolve-PlayerPath
    if ($resolved) {
        $script:mpvPath = [string]$resolved.Path
        $script:playerBackend = [string]$resolved.Backend
        return $true
    }

    $choice = [System.Windows.Forms.MessageBox]::Show(
        "No supported player was found (VLC/FFplay/mpv).`n`nYes: try automatic install with winget (recommended)`nNo: select player .exe manually`nCancel: exit",
        $script:AppName,
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    switch ($choice) {
        ([System.Windows.Forms.DialogResult]::Yes) {
            if (Try-InstallMpvWithWinget) {
                $resolved = Resolve-PlayerPath
                if ($resolved) {
                    $script:mpvPath = [string]$resolved.Path
                    $script:playerBackend = [string]$resolved.Backend
                    Show-Balloon -Title $script:AppName -Text "Player installed and ready."
                    return $true
                }
            }

            $openPage = [System.Windows.Forms.MessageBox]::Show(
                "Automatic install failed.`nOpen VLC download page?",
                $script:AppName,
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            if ($openPage -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process "https://www.videolan.org/vlc/" | Out-Null
            }

            $picked = Select-MpvExecutable
            if ($picked) {
                $script:mpvPath = $picked
                $script:playerBackend = Get-BackendFromExecutable -Path $picked
                return $true
            }

            return $false
        }
        ([System.Windows.Forms.DialogResult]::No) {
            $picked = Select-MpvExecutable
            if ($picked) {
                $script:mpvPath = $picked
                $script:playerBackend = Get-BackendFromExecutable -Path $picked
                return $true
            }
            return $false
        }
        default {
            return $false
        }
    }
}

function Load-Stations {
    if (-not (Test-Path -Path $script:StationsPath)) {
        throw "stations.json not found: $script:StationsPath"
    }

    $stations = Read-JsonFile -Path $script:StationsPath
    if ($null -eq $stations) {
        throw "stations.json is empty or invalid."
    }

    $validated = New-Object "System.Collections.Generic.List[object]"
    $seenIds = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($station in @($stations)) {
        $stationId = [string](Get-ObjectPropertyValue -Object $station -Name "id")
        $stationName = [string](Get-ObjectPropertyValue -Object $station -Name "name")
        $stationUrl = [string](Get-ObjectPropertyValue -Object $station -Name "url")
        if ([string]::IsNullOrWhiteSpace($stationId) -or [string]::IsNullOrWhiteSpace($stationName)) {
            continue
        }
        if (-not (Test-IsSafeStreamUrl -Url $stationUrl)) {
            continue
        }
        if (-not $seenIds.Add($stationId)) {
            continue
        }
        [void]$validated.Add($station)
    }

    if ($validated.Count -eq 0) {
        throw "No valid stations were found in stations.json."
    }

    $script:allStations = $validated.ToArray()
}

function Get-StationById {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    return $script:allStations | Where-Object { $_.id -eq $Id } | Select-Object -First 1
}

function Test-IsSupportedPlayerProcessName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName
    )

    if ([string]::IsNullOrWhiteSpace($ProcessName)) {
        return $false
    }

    $name = $ProcessName.ToLowerInvariant()
    return ($name -like "mpv*" -or $name -like "vlc*" -or $name -like "ffplay*")
}

function Test-MatchesBackendProcessName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName,
        [string]$Backend
    )

    if ([string]::IsNullOrWhiteSpace($Backend)) {
        return (Test-IsSupportedPlayerProcessName -ProcessName $ProcessName)
    }

    switch ($Backend.ToLowerInvariant()) {
        "vlc" { return ($ProcessName -like "vlc*") }
        "ffplay" { return ($ProcessName -like "ffplay*") }
        default { return ($ProcessName -like "mpv*") }
    }
}

function Find-OwnedPlayerPids {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]$State,
        [string]$StationUrlHint = $null
    )

    $ids = New-Object "System.Collections.Generic.HashSet[int]"

    if ($script:playerProcess) {
        try {
            [void]$ids.Add([int]$script:playerProcess.Id)
        } catch {
        }
    }

    $statePid = Get-ObjectPropertyValue -Object $State -Name "pid"
    if ($statePid) {
        try {
            [void]$ids.Add([int]$statePid)
        } catch {
        }
    }

    $urlHints = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    if ($script:currentStation -and $script:currentStation.url) {
        [void]$urlHints.Add([string]$script:currentStation.url)
    }
    $stateStationUrl = [string](Get-ObjectPropertyValue -Object $State -Name "stationUrl")
    if (-not [string]::IsNullOrWhiteSpace($stateStationUrl)) {
        [void]$urlHints.Add($stateStationUrl)
    }
    if (-not [string]::IsNullOrWhiteSpace($StationUrlHint)) {
        [void]$urlHints.Add($StationUrlHint)
    }
    $markerHint = [string](Get-ObjectPropertyValue -Object $State -Name "marker" -Default $script:playerMarker)
    if ([string]::IsNullOrWhiteSpace($markerHint)) {
        $markerHint = $script:playerMarker
    }

    try {
        $procs = Get-CimInstance Win32_Process -Filter "Name='mpv.exe' OR Name='vlc.exe' OR Name='ffplay.exe'" -ErrorAction Stop
        foreach ($proc in $procs) {
            $procId = [int]$proc.ProcessId
            $isOwned = $ids.Contains($procId)
            $cmd = [string]$proc.CommandLine

            if (-not $isOwned -and -not [string]::IsNullOrWhiteSpace($cmd)) {
                if ($cmd.IndexOf($markerHint, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $isOwned = $true
                }
            }

            if (-not $isOwned -and -not [string]::IsNullOrWhiteSpace($cmd)) {
                foreach ($url in $urlHints) {
                    if (-not [string]::IsNullOrWhiteSpace($url) -and $cmd.IndexOf($url, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $isOwned = $true
                        break
                    }
                }
            }

            if ($isOwned) {
                try {
                    [void]$ids.Add([int]$proc.ProcessId)
                } catch {
                }
            }
        }
    } catch {
    }

    return @($ids)
}

function Find-OwnedPlayerProcess {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]$State,
        [string]$StationUrlHint = $null
    )

    foreach ($procId in (Find-OwnedPlayerPids -State $State -StationUrlHint $StationUrlHint)) {
        try {
            $proc = Get-Process -Id ([int]$procId) -ErrorAction Stop
            if (Test-IsSupportedPlayerProcessName -ProcessName ([string]$proc.ProcessName)) {
                return $proc
            }
        } catch {
        }
    }

    return $null
}

function Save-State {
    Ensure-DataRoot
    $state = [ordered]@{
        stationId   = if ($script:currentStation) { $script:currentStation.id } else { $null }
        stationName = if ($script:currentStation) { $script:currentStation.name } else { $null }
        stationUrl  = if ($script:currentStation) { $script:currentStation.url } else { $null }
        pid         = if ($script:playerProcess) { $script:playerProcess.Id } else { $null }
        backend     = if ($script:playerBackend) { $script:playerBackend } else { $null }
        playerPath  = if ($script:mpvPath) { $script:mpvPath } else { $null }
        marker      = $script:playerMarker
        updatedAt   = (Get-Date).ToString("o")
    }
    Save-JsonFile -Path $script:StatePath -Value $state
}

function Try-RestoreState {
    $state = Read-JsonFile -Path $script:StatePath
    if ($null -eq $state) {
        return
    }

    $stateStationId = [string](Get-ObjectPropertyValue -Object $state -Name "stationId")
    if (-not [string]::IsNullOrWhiteSpace($stateStationId)) {
        $script:currentStation = Get-StationById -Id $stateStationId
    }

    $statePlayerPath = [string](Get-ObjectPropertyValue -Object $state -Name "playerPath")
    if (-not [string]::IsNullOrWhiteSpace($statePlayerPath) -and (Test-Path -Path $statePlayerPath)) {
        $script:mpvPath = $statePlayerPath
    }

    $stateBackend = [string](Get-ObjectPropertyValue -Object $state -Name "backend")
    if (-not [string]::IsNullOrWhiteSpace($stateBackend)) {
        $script:playerBackend = $stateBackend
    }

    $statePid = Get-ObjectPropertyValue -Object $state -Name "pid"
    if ($statePid) {
        try {
            $proc = Get-Process -Id ([int]$statePid) -ErrorAction Stop
            if (Test-MatchesBackendProcessName -ProcessName ([string]$proc.ProcessName) -Backend $stateBackend) {
                $script:playerProcess = $proc
            }
        } catch {
            $script:playerProcess = $null
        }
    }

    if (-not $script:playerProcess) {
        $stateStationUrl = [string](Get-ObjectPropertyValue -Object $state -Name "stationUrl")
        $restored = Find-OwnedPlayerProcess -State $state -StationUrlHint $stateStationUrl
        if ($restored) {
            $script:playerProcess = $restored
        }
    }
}

function Get-ProfileSessionVolume {
    $fallbackVolume = if ($script:profile -and $script:profile.PSObject.Properties["volume"]) { [int]$script:profile.volume } else { 35 }
    $volume = if ($script:profile -and $script:profile.PSObject.Properties["sessionVolume"]) { [int]$script:profile.sessionVolume } else { $fallbackVolume }
    return [Math]::Max([Math]::Min($volume, 100), 0)
}

function Get-ProfileSessionMuted {
    if ($script:profile -and $script:profile.PSObject.Properties["sessionMuted"]) {
        return [bool]$script:profile.sessionMuted
    }

    return $false
}

function Set-ProfileSessionPreferences {
    param(
        [int]$VolumePercent = -1,
        [object]$Muted = $null,
        [switch]$SkipSave
    )

    if (-not $script:profile.PSObject.Properties["sessionVolume"]) {
        $defaultSessionVolume = if ($script:profile -and $script:profile.PSObject.Properties["volume"]) { [int]$script:profile.volume } else { 35 }
        $script:profile | Add-Member -NotePropertyName sessionVolume -NotePropertyValue $defaultSessionVolume
    }
    if (-not $script:profile.PSObject.Properties["sessionMuted"]) {
        $script:profile | Add-Member -NotePropertyName sessionMuted -NotePropertyValue $false
    }

    $changed = $false
    if ($VolumePercent -ge 0) {
        $clamped = [Math]::Max([Math]::Min([int]$VolumePercent, 100), 0)
        if ([int]$script:profile.sessionVolume -ne $clamped) {
            $script:profile.sessionVolume = $clamped
            $changed = $true
        }
    }

    if ($null -ne $Muted) {
        $mutedValue = [bool]$Muted
        if ([bool]$script:profile.sessionMuted -ne $mutedValue) {
            $script:profile.sessionMuted = $mutedValue
            $changed = $true
        }
    }

    if ($changed -and -not $SkipSave) {
        Save-Profile
    }

    return $changed
}

function Get-CurrentPlayerProcessId {
    if ($script:playerProcess) {
        try {
            $script:playerProcess.Refresh()
            if (-not $script:playerProcess.HasExited) {
                return [int]$script:playerProcess.Id
            }
        } catch {
            $script:playerProcess = $null
        }
    }

    return 0
}

function Reset-AudioSessionState {
    $script:audioSessionState.VolumePercent = Get-ProfileSessionVolume
    $script:audioSessionState.IsMuted = Get-ProfileSessionMuted
    $script:audioSessionState.LastPid = $null
    $script:audioSessionState.AppliedPid = $null
    $script:audioSessionState.PendingVolumePercent = $null
    $script:audioSessionState.PendingMuted = $null
    $script:audioSessionState.UpdatedAt = (Get-Date).ToString("o")
}

function Update-VolumeMenuState {
    $volume = if ($script:audioSessionState -and $null -ne $script:audioSessionState.VolumePercent) { [int]$script:audioSessionState.VolumePercent } else { Get-ProfileSessionVolume }
    $muted = if ($script:audioSessionState) { [bool]$script:audioSessionState.IsMuted } else { Get-ProfileSessionMuted }
    $isPlaying = Get-IsPlaying

    if ($script:volumeItem) {
        if ($muted) {
            $script:volumeItem.Text = "Volume: Muted ($volume%)"
        } else {
            $script:volumeItem.Text = "Volume: $volume%"
        }
    }

    if ($script:volumeUpItem) {
        $script:volumeUpItem.Enabled = ($isPlaying -and $volume -lt 100)
    }

    if ($script:volumeDownItem) {
        $script:volumeDownItem.Enabled = ($isPlaying -and $volume -gt 0)
    }

    if ($script:muteItem) {
        $script:muteItem.Text = if ($muted) { "Unmute" } else { "Mute" }
        $script:muteItem.Enabled = $isPlaying
    }
}

function Set-AudioSessionSnapshot {
    param(
        [int]$ProcessId = 0,
        [int]$VolumePercent = 100,
        [bool]$IsMuted = $false,
        [switch]$PersistToProfile
    )

    $clamped = [Math]::Max([Math]::Min([int]$VolumePercent, 100), 0)
    $script:audioSessionState.VolumePercent = $clamped
    $script:audioSessionState.IsMuted = [bool]$IsMuted
    $script:audioSessionState.LastPid = if ($ProcessId -gt 0) { [int]$ProcessId } else { $null }
    $script:audioSessionState.UpdatedAt = (Get-Date).ToString("o")

    if ($PersistToProfile) {
        Set-ProfileSessionPreferences -VolumePercent $clamped -Muted $IsMuted | Out-Null
    }

    Update-VolumeMenuState
}

function Set-PendingAudioSessionApply {
    param(
        [Parameter(Mandatory = $true)]
        [int]$VolumePercent,
        [bool]$IsMuted = $false
    )

    $script:audioSessionState.PendingVolumePercent = [Math]::Max([Math]::Min([int]$VolumePercent, 100), 0)
    $script:audioSessionState.PendingMuted = [bool]$IsMuted
}

function Clear-PendingAudioSessionApply {
    $script:audioSessionState.PendingVolumePercent = $null
    $script:audioSessionState.PendingMuted = $null
}

function Get-PlayerAudioSessionSnapshot {
    param(
        [int]$ProcessId = 0
    )

    $playerProcessId = if ($ProcessId -gt 0) { [int]$ProcessId } else { Get-CurrentPlayerProcessId }
    if ($playerProcessId -le 0) {
        return $null
    }

    [single]$volumeScalar = 1.0
    [bool]$muted = $false
    $found = $false

    try {
        $found = [AudioSessionHelper]::TryGetSnapshot($playerProcessId, [ref]$volumeScalar, [ref]$muted)
    } catch {
        return $null
    }

    if (-not $found) {
        return $null
    }

    $percent = [int][Math]::Round([Math]::Max([Math]::Min([double]$volumeScalar, 1.0), 0.0) * 100)
    return [pscustomobject]@{
        ProcessId      = $playerProcessId
        VolumePercent  = $percent
        IsMuted        = [bool]$muted
    }
}

function Refresh-PlayerAudioSessionState {
    param(
        [int]$ProcessId = 0,
        [switch]$SkipPersist
    )

    $snapshot = Get-PlayerAudioSessionSnapshot -ProcessId $ProcessId
    if (-not $snapshot) {
        return $false
    }

    Set-AudioSessionSnapshot -ProcessId ([int]$snapshot.ProcessId) -VolumePercent ([int]$snapshot.VolumePercent) -IsMuted ([bool]$snapshot.IsMuted) -PersistToProfile:(-not $SkipPersist)
    return $true
}

function Apply-PlayerAudioSessionState {
    param(
        [int]$ProcessId = 0,
        [Parameter(Mandatory = $true)]
        [int]$VolumePercent,
        [bool]$Muted = $false,
        [switch]$PersistToProfile
    )

    $playerProcessId = if ($ProcessId -gt 0) { [int]$ProcessId } else { Get-CurrentPlayerProcessId }
    if ($playerProcessId -le 0) {
        return $false
    }

    $clampedVolume = [Math]::Max([Math]::Min([int]$VolumePercent, 100), 0)
    [single]$targetScalar = [single]($clampedVolume / 100.0)
    [single]$appliedScalar = 1.0
    [bool]$appliedMuted = $false

    $applied = $false
    try {
        $applied = [AudioSessionHelper]::TrySetState($playerProcessId, $targetScalar, $Muted, [ref]$appliedScalar, [ref]$appliedMuted)
    } catch {
        return $false
    }

    if (-not $applied) {
        return $false
    }

    if ($PersistToProfile) {
        Set-ProfileSessionPreferences -VolumePercent $clampedVolume -Muted $Muted | Out-Null
    }

    $script:audioSessionState.AppliedPid = $playerProcessId
    Clear-PendingAudioSessionApply
    $appliedPercent = [int][Math]::Round([Math]::Max([Math]::Min([double]$appliedScalar, 1.0), 0.0) * 100)
    Set-AudioSessionSnapshot -ProcessId $playerProcessId -VolumePercent $appliedPercent -IsMuted ([bool]$appliedMuted)
    return $true
}

function Apply-PlayerAudioSessionProfile {
    param(
        [int]$ProcessId = 0
    )

    return (Apply-PlayerAudioSessionState -ProcessId $ProcessId -VolumePercent (Get-ProfileSessionVolume) -Muted (Get-ProfileSessionMuted))
}

function Set-PlayerSessionVolume {
    param(
        [Parameter(Mandatory = $true)]
        [int]$VolumePercent
    )

    $clamped = [Math]::Max([Math]::Min([int]$VolumePercent, 100), 0)

    $playerProcessId = Get-CurrentPlayerProcessId
    if ($playerProcessId -gt 0) {
        Set-PendingAudioSessionApply -VolumePercent $clamped -IsMuted $false
        $applied = Apply-PlayerAudioSessionState -ProcessId $playerProcessId -VolumePercent $clamped -Muted $false -PersistToProfile
        if (-not $applied) {
            Set-ProfileSessionPreferences -VolumePercent $clamped -Muted $false | Out-Null
            Set-AudioSessionSnapshot -ProcessId $playerProcessId -VolumePercent $clamped -IsMuted $false
        }
    } else {
        Set-ProfileSessionPreferences -VolumePercent $clamped -Muted $false | Out-Null
        Set-AudioSessionSnapshot -VolumePercent $clamped -IsMuted $false
    }

    Update-TrayStatus
}

function Adjust-PlayerSessionVolume {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Delta
    )

    $current = if ($script:audioSessionState -and $null -ne $script:audioSessionState.VolumePercent) { [int]$script:audioSessionState.VolumePercent } else { Get-ProfileSessionVolume }
    Set-PlayerSessionVolume -VolumePercent ($current + $Delta)
}

function Set-PlayerSessionMute {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Muted
    )

    $targetVolume = if ($script:audioSessionState -and $null -ne $script:audioSessionState.VolumePercent) { [int]$script:audioSessionState.VolumePercent } else { Get-ProfileSessionVolume }

    $playerProcessId = Get-CurrentPlayerProcessId
    if ($playerProcessId -gt 0) {
        Set-PendingAudioSessionApply -VolumePercent $targetVolume -IsMuted $Muted
        $applied = Apply-PlayerAudioSessionState -ProcessId $playerProcessId -VolumePercent $targetVolume -Muted $Muted -PersistToProfile
        if (-not $applied) {
            Set-ProfileSessionPreferences -Muted $Muted | Out-Null
            Set-AudioSessionSnapshot -ProcessId $playerProcessId -VolumePercent $targetVolume -IsMuted $Muted
        }
    } else {
        Set-ProfileSessionPreferences -Muted $Muted | Out-Null
        Set-AudioSessionSnapshot -VolumePercent $targetVolume -IsMuted $Muted
    }

    Update-TrayStatus
}

function Toggle-PlayerSessionMute {
    $currentMuted = if ($script:audioSessionState) { [bool]$script:audioSessionState.IsMuted } else { Get-ProfileSessionMuted }
    Set-PlayerSessionMute -Muted (-not $currentMuted)
}

function Start-AudioSessionTimer {
    param(
        [int]$InitialIntervalMs = 500
    )

    if ($script:audioSessionTimer) {
        try {
            $script:audioSessionTimer.Stop()
            $script:audioSessionTimer.Interval = [Math]::Max($InitialIntervalMs, 250)
            $script:audioSessionTimer.Start()
        } catch {
        }
        return
    }

    $script:audioSessionTimer = New-Object System.Windows.Forms.Timer
    $script:audioSessionTimer.Interval = [Math]::Max($InitialIntervalMs, 250)
    $script:audioSessionTimer.Add_Tick({
            try {
                if (-not (Get-IsPlaying)) {
                    Stop-AudioSessionTimer
                    return
                }

                $playerProcessId = Get-CurrentPlayerProcessId
                if ($playerProcessId -le 0) {
                    return
                }

                $pendingVolume = Get-ObjectPropertyValue -Object $script:audioSessionState -Name "PendingVolumePercent"
                $pendingMuted = [bool](Get-ObjectPropertyValue -Object $script:audioSessionState -Name "PendingMuted" -Default $false)
                if ($null -ne $pendingVolume) {
                    if (-not (Apply-PlayerAudioSessionState -ProcessId $playerProcessId -VolumePercent ([int]$pendingVolume) -Muted $pendingMuted)) {
                        return
                    }
                } elseif ($script:audioSessionState.AppliedPid -ne $playerProcessId) {
                    Refresh-PlayerAudioSessionState -ProcessId $playerProcessId | Out-Null
                } else {
                    Refresh-PlayerAudioSessionState -ProcessId $playerProcessId | Out-Null
                }

                if ($script:audioSessionTimer -and $script:audioSessionTimer.Interval -lt 2000) {
                    $script:audioSessionTimer.Interval = 2000
                }
            } catch {
            }
        })
    $script:audioSessionTimer.Start()
}

function Stop-AudioSessionTimer {
    if (-not $script:audioSessionTimer) {
        return
    }

    try {
        $script:audioSessionTimer.Stop()
        $script:audioSessionTimer.Dispose()
    } catch {
    }

    $script:audioSessionTimer = $null
}

function Set-NotifyText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    if ($null -eq $script:notifyIcon) {
        return
    }

    $safeText = $Text
    if ($safeText.Length -gt 63) {
        $safeText = $safeText.Substring(0, 63)
    }
    $script:notifyIcon.Text = $safeText
}

function Show-Balloon {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [int]$TimeoutMs = 2500
    )

    if ($null -eq $script:notifyIcon) {
        return
    }

    $script:notifyIcon.BalloonTipTitle = $Title
    $script:notifyIcon.BalloonTipText = $Text
    $script:notifyIcon.ShowBalloonTip($TimeoutMs)
}

function Reset-StreamInfo {
    $script:streamInfo.StreamTitle = $null
    $script:streamInfo.BitrateKbps = $null
    $script:streamInfo.UpdatedAt = $null
}

function Get-HeaderValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.WebHeaderCollection]$Headers,
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }
        $value = $Headers[$name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return [string]$value
        }
    }

    return $null
}

function Convert-ToIntOrNull {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $match = [regex]::Match($Text, "\d+")
    if (-not $match.Success) {
        return $null
    }

    try {
        return [int]$match.Value
    } catch {
        return $null
    }
}

function Get-StreamMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [int]$TimeoutMs = 2500
    )

    if (-not (Test-IsSafeStreamUrl -Url $Url)) {
        return $null
    }

    $request = [System.Net.HttpWebRequest]::Create($Url)
    $request.Method = "GET"
    $request.UserAgent = "$script:AppName/1.0"
    $request.AllowAutoRedirect = $true
    $request.KeepAlive = $false
    $request.Timeout = $TimeoutMs
    $request.ReadWriteTimeout = $TimeoutMs
    $request.Headers["Icy-MetaData"] = "1"

    $response = $null
    try {
        $response = $request.GetResponse()
        $headers = $response.Headers

        $bitrateRaw = Get-HeaderValue -Headers $headers -Names @("icy-br", "ice-bitrate", "x-audiocast-bitrate")
        $bitrate = Convert-ToIntOrNull -Text $bitrateRaw

        $songTitle = $null
        $metaIntRaw = Get-HeaderValue -Headers $headers -Names @("icy-metaint")
        $metaInt = Convert-ToIntOrNull -Text $metaIntRaw

        if ($metaInt -and $metaInt -gt 0) {
            $stream = $response.GetResponseStream()
            if ($stream) {
                $audioBuffer = New-Object byte[] $metaInt
                $audioRead = 0
                while ($audioRead -lt $metaInt) {
                    $chunk = $stream.Read($audioBuffer, $audioRead, $metaInt - $audioRead)
                    if ($chunk -le 0) {
                        break
                    }
                    $audioRead += $chunk
                }

                if ($audioRead -eq $metaInt) {
                    $metaLengthFlag = $stream.ReadByte()
                    if ($metaLengthFlag -gt 0) {
                        $metaBytesLen = $metaLengthFlag * 16
                        $metaBytes = New-Object byte[] $metaBytesLen
                        $metaRead = 0
                        while ($metaRead -lt $metaBytesLen) {
                            $chunk = $stream.Read($metaBytes, $metaRead, $metaBytesLen - $metaRead)
                            if ($chunk -le 0) {
                                break
                            }
                            $metaRead += $chunk
                        }

                        if ($metaRead -gt 0) {
                            $metaText = [System.Text.Encoding]::UTF8.GetString($metaBytes, 0, $metaRead)
                            if ([string]::IsNullOrWhiteSpace($metaText)) {
                                $metaText = [System.Text.Encoding]::ASCII.GetString($metaBytes, 0, $metaRead)
                            }
                            $titleMatch = [regex]::Match($metaText, "StreamTitle='([^']*)';", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                            if ($titleMatch.Success) {
                                $candidate = [string]$titleMatch.Groups[1].Value
                                if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                                    $songTitle = $candidate.Trim()
                                }
                            }
                        }
                    }
                }
            }
        }

        return [pscustomobject]@{
            BitrateKbps = $bitrate
            StreamTitle = $songTitle
        }
    } catch {
        return $null
    } finally {
        if ($response) {
            try { $response.Close() } catch { }
        }
    }
}

function Refresh-StreamInfo {
    param(
        [int]$TimeoutMs = 2500
    )

    if (-not $script:currentStation -or -not $script:currentStation.url) {
        return
    }

    if (-not (Get-IsPlaying)) {
        return
    }

    $metadata = Get-StreamMetadata -Url ([string]$script:currentStation.url) -TimeoutMs $TimeoutMs
    if ($metadata) {
        if ($metadata.PSObject.Properties["BitrateKbps"] -and $metadata.BitrateKbps) {
            $script:streamInfo.BitrateKbps = [int]$metadata.BitrateKbps
        }
        if ($metadata.PSObject.Properties["StreamTitle"] -and -not [string]::IsNullOrWhiteSpace([string]$metadata.StreamTitle)) {
            $script:streamInfo.StreamTitle = [string]$metadata.StreamTitle
        }
        $script:streamInfo.UpdatedAt = (Get-Date).ToString("o")
    }

    Update-TrayStatus
}

function Start-MetadataTimer {
    param(
        [int]$InitialIntervalMs = 4000
    )

    if ($script:metadataTimer) {
        try {
            $script:metadataTimer.Stop()
            $script:metadataTimer.Interval = [Math]::Max($InitialIntervalMs, 1500)
            $script:metadataTimer.Start()
        } catch {
        }
        return
    }

    $script:metadataTimer = New-Object System.Windows.Forms.Timer
    $script:metadataTimer.Interval = [Math]::Max($InitialIntervalMs, 1500)
    $script:metadataTimer.Add_Tick({
            try {
                if (Get-IsPlaying) {
                    Refresh-StreamInfo -TimeoutMs 1500
                    if ($script:metadataTimer -and $script:metadataTimer.Interval -lt 12000) {
                        $script:metadataTimer.Interval = 12000
                    }
                } else {
                    Stop-MetadataTimer
                }
            } catch {
            }
        })
    $script:metadataTimer.Start()
}

function Stop-MetadataTimer {
    if (-not $script:metadataTimer) {
        return
    }

    try {
        $script:metadataTimer.Stop()
        $script:metadataTimer.Dispose()
    } catch {
    }
    $script:metadataTimer = $null
}

function New-MusicTrayIcon {
    $size = 32
    $bitmap = New-Object System.Drawing.Bitmap($size, $size)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $bgBrush = $null
    $fgBrush = $null
    $font = $null
    $format = $null

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $bgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(21, 101, 192))
        $graphics.FillEllipse($bgBrush, 0, 0, $size - 1, $size - 1)

        $font = New-Object System.Drawing.Font("Segoe UI Symbol", 18, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $fgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $format = New-Object System.Drawing.StringFormat
        $format.Alignment = [System.Drawing.StringAlignment]::Center
        $format.LineAlignment = [System.Drawing.StringAlignment]::Center
        $graphics.DrawString("♪", $font, $fgBrush, (New-Object System.Drawing.RectangleF(0, 0, $size, $size)), $format)

        $iconHandle = $bitmap.GetHicon()
        try {
            return [System.Drawing.Icon]::FromHandle($iconHandle).Clone()
        } finally {
            [RofiHotkeyWindow]::DestroyIcon($iconHandle) | Out-Null
        }
    } catch {
        return [System.Drawing.SystemIcons]::Information
    } finally {
        if ($format) { $format.Dispose() }
        if ($font) { $font.Dispose() }
        if ($fgBrush) { $fgBrush.Dispose() }
        if ($bgBrush) { $bgBrush.Dispose() }
        if ($graphics) { $graphics.Dispose() }
        if ($bitmap) { $bitmap.Dispose() }
    }
}

function Get-IsPlaying {
    if ($script:playerProcess) {
        try {
            $script:playerProcess.Refresh()
            if (-not $script:playerProcess.HasExited) {
                return $true
            }
        } catch {
            $script:playerProcess = $null
        }
    }

    $state = Read-JsonFile -Path $script:StatePath
    $statePid = Get-ObjectPropertyValue -Object $state -Name "pid"
    $stateBackend = [string](Get-ObjectPropertyValue -Object $state -Name "backend")
    if ($state -and $statePid) {
        try {
            $proc = Get-Process -Id ([int]$statePid) -ErrorAction Stop
            if (Test-MatchesBackendProcessName -ProcessName ([string]$proc.ProcessName) -Backend $stateBackend) {
                $script:playerProcess = $proc
                return $true
            }
        } catch {
        }
    }

    if ($state) {
        $stateStationUrl = [string](Get-ObjectPropertyValue -Object $state -Name "stationUrl")
        $owned = Find-OwnedPlayerProcess -State $state -StationUrlHint $stateStationUrl
        if ($owned) {
            $script:playerProcess = $owned
            return $true
        }
    }

    return $false
}

function Update-TrayStatus {
    if ($null -eq $script:notifyIcon) {
        return
    }

    Update-VolumeMenuState

    if (Get-IsPlaying) {
        $currentName = if ($script:currentStation) { [string]$script:currentStation.name } else { "Unknown station" }
        $songText = if ($script:streamInfo -and -not [string]::IsNullOrWhiteSpace([string]$script:streamInfo.StreamTitle)) { [string]$script:streamInfo.StreamTitle } else { "Unknown" }
        $bitrateText = if ($script:streamInfo -and $script:streamInfo.BitrateKbps) { "{0} kbps" -f [int]$script:streamInfo.BitrateKbps } else { "Unknown" }
        $script:statusItem.Text = "Status: Playing"
        $script:currentItem.Text = "Now: $currentName"
        if ($script:songItem) {
            $script:songItem.Text = "Song: $songText"
        }
        if ($script:bitrateItem) {
            $script:bitrateItem.Text = "Bitrate: $bitrateText"
        }
        $script:toggleItem.Text = "Stop (Ctrl+Alt+B)"
        if (-not $script:metadataTimer) {
            Start-MetadataTimer
        }
        if (-not $script:audioSessionTimer) {
            Start-AudioSessionTimer
        }
        if ($songText -ne "Unknown") {
            Set-NotifyText -Text "$script:AppName - $songText"
        } else {
            Set-NotifyText -Text "$script:AppName - $currentName"
        }
    } else {
        $script:statusItem.Text = "Status: Idle"
        $script:currentItem.Text = "Now: Off"
        if ($script:songItem) {
            $script:songItem.Text = "Song: -"
        }
        if ($script:bitrateItem) {
            $script:bitrateItem.Text = "Bitrate: -"
        }
        $script:toggleItem.Text = "Play (Ctrl+Alt+B)"
        Stop-MetadataTimer
        Stop-AudioSessionTimer
        Set-NotifyText -Text $script:AppName
    }
}

function Stop-Playback {
    param(
        [switch]$Silent
    )

    $stopped = $false
    $state = Read-JsonFile -Path $script:StatePath
    $stateStationUrl = [string](Get-ObjectPropertyValue -Object $state -Name "stationUrl")
    $statePid = Get-ObjectPropertyValue -Object $state -Name "pid"
    $urlHint = if ($script:currentStation -and $script:currentStation.url) { [string]$script:currentStation.url } elseif (-not [string]::IsNullOrWhiteSpace($stateStationUrl)) { $stateStationUrl } else { $null }

    if ($script:playerProcess) {
        try {
            $script:playerProcess.Refresh()
            if (-not $script:playerProcess.HasExited) {
                Stop-Process -Id $script:playerProcess.Id -Force -ErrorAction SilentlyContinue
                $stopped = $true
            }
        } catch {
        }
    }

    if ($state -and $statePid) {
        try {
            Stop-Process -Id ([int]$statePid) -Force -ErrorAction SilentlyContinue
            $stopped = $true
        } catch {
        }
    }

    foreach ($procId in (Find-OwnedPlayerPids -State $state -StationUrlHint $urlHint)) {
        try {
            $proc = Get-Process -Id ([int]$procId) -ErrorAction Stop
            Stop-Process -Id ([int]$proc.Id) -Force -ErrorAction SilentlyContinue
            $stopped = $true
        } catch {
        }
    }

    $script:playerProcess = $null
    $script:currentStation = $null
    Stop-MetadataTimer
    Stop-AudioSessionTimer
    Reset-StreamInfo
    Reset-AudioSessionState
    Save-State
    Update-TrayStatus

    if (-not $Silent -and $stopped) {
        Show-Balloon -Title $script:AppName -Text "Playback stopped."
    }
}

function Start-Station {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Station
    )

    if (-not $script:mpvPath -or -not (Test-Path -Path $script:mpvPath)) {
        if (-not (Ensure-MpvAvailable)) {
            throw "No supported player found."
        }
    }

    if (-not $script:playerBackend) {
        $script:playerBackend = Get-BackendFromExecutable -Path $script:mpvPath
    }

    if (-not (Test-IsSafeStreamUrl -Url ([string]$Station.url))) {
        throw "Station URL is invalid or unsupported."
    }

    Stop-Playback -Silent

    $volume = [Math]::Max([Math]::Min([int]$script:profile.volume, 100), 0)
    $maxStartupVolume = if ($script:profile.PSObject.Properties["maxStartupVolume"] -and $null -ne $script:profile.maxStartupVolume) { [int]$script:profile.maxStartupVolume } else { 45 }
    $maxStartupVolume = [Math]::Max([Math]::Min($maxStartupVolume, 100), 0)
    $sessionVolume = Get-ProfileSessionVolume
    $startupSessionVolume = [Math]::Min($sessionVolume, $maxStartupVolume)
    switch ($script:playerBackend) {
        "vlc" {
            $vlcVolume = 256
            $args = @(
                "--intf", "dummy",
                "--dummy-quiet",
                "--no-video",
                "--no-video-title-show",
                "--network-caching=1000",
                "--meta-title=$script:playerMarker",
                "--volume=$vlcVolume",
                [string]$Station.url
            )
        }
        "ffplay" {
            $args = @(
                "-nodisp",
                "-autoexit",
                "-loglevel", "quiet",
                "-window_title", $script:playerMarker,
                "-volume", 100,
                [string]$Station.url
            )
        }
        default {
            $args = @(
                "--no-video",
                "--force-window=no",
                "--title=$script:playerMarker",
                "--volume=100",
                "--quiet",
                [string]$Station.url
            )
            $script:playerBackend = "mpv"
        }
    }

    $proc = Start-Process -FilePath $script:mpvPath -ArgumentList $args -PassThru -WindowStyle Hidden

    $script:playerProcess = $proc
    $script:currentStation = $Station
    Reset-AudioSessionState
    $script:profile.lastStationId = [string]$Station.id
    Save-Profile
    Reset-StreamInfo
    Save-State
    Update-TrayStatus
    Set-PendingAudioSessionApply -VolumePercent $startupSessionVolume -IsMuted (Get-ProfileSessionMuted)
    Start-AudioSessionTimer -InitialIntervalMs 350
    Start-MetadataTimer
    Show-Balloon -Title $script:AppName -Text ("Playing: {0}" -f [string]$Station.name)
    try {
        $deadline = (Get-Date).AddMilliseconds(2500)
        while ((Get-Date) -lt $deadline) {
            if (Apply-PlayerAudioSessionState -ProcessId ([int]$proc.Id) -VolumePercent $startupSessionVolume -Muted (Get-ProfileSessionMuted)) {
                break
            }
            Start-Sleep -Milliseconds 100
        }
        Refresh-StreamInfo -TimeoutMs 1500
    } catch {
    }
}

function Get-RecommendationScore {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Station,
        [string]$MoodKey = "all",
        [string]$GenreKey = "all",
        [bool]$PreferTurkish = $false,
        [string]$DiscoveryMode = "balanced"
    )

    $score = 0
    $stationMoods = To-Array -Value $Station.moods
    $stationGenres = To-Array -Value $Station.genres

    if ($MoodKey -ne "all") {
        if ($stationMoods -contains $MoodKey) {
            $score += 4
        } else {
            $score -= 1
        }
    }

    if ($GenreKey -ne "all") {
        if (Test-StationMatchesGenreSelection -Station $Station -GenreKey $GenreKey) {
            $score += 4
        } else {
            $score -= 1
        }
    }

    if ($PreferTurkish) {
        if ([bool]$Station.turkish) {
            $score += 2
        } else {
            $score -= 1
        }
    } elseif (-not [bool]$Station.turkish) {
        $score += 1
    }

    if ([bool]$Station.public) {
        $score += 1
    }

    $isLastStation = ($script:profile.lastStationId -and $script:profile.lastStationId -eq $Station.id)
    $isCurrentStation = ($script:currentStation -and $script:currentStation.id -eq $Station.id)

    if ($isLastStation -or $isCurrentStation) {
        switch ($DiscoveryMode) {
            "new" {
                $score -= 3
            }
            "surprise" {
                $score -= 4
            }
            default {
                $score += 1
            }
        }
    }

    switch ($DiscoveryMode) {
        "new" {
            $score += Get-Random -Minimum 0 -Maximum 3
        }
        "surprise" {
            $score += Get-Random -Minimum -2 -Maximum 5
        }
        default {
            $score += Get-Random -Minimum 0 -Maximum 2
        }
    }

    return $score
}

function Get-RankedStations {
    param(
        [string]$SearchText = "",
        [string]$MoodKey = "all",
        [string]$GenreKey = "all",
        [bool]$TurkishOnly = $false,
        [bool]$PublicOnly = $true,
        [string]$DiscoveryMode = "balanced"
    )

    $searchNormalized = $SearchText.Trim().ToLowerInvariant()
    $scored = foreach ($station in $script:allStations) {
        if ($TurkishOnly -and -not [bool]$station.turkish) {
            continue
        }

        if ($PublicOnly -and -not [bool]$station.public) {
            continue
        }

        $stationMoods = To-Array -Value $station.moods
        $stationGenres = To-Array -Value $station.genres

        if ($MoodKey -ne "all" -and -not ($stationMoods -contains $MoodKey)) {
            continue
        }

        if ($GenreKey -ne "all" -and -not (Test-StationMatchesGenreSelection -Station $station -GenreKey $GenreKey)) {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($searchNormalized)) {
            $blob = ("{0} {1} {2} {3}" -f $station.name, $station.country, ($stationGenres -join " "), ($stationMoods -join " ")).ToLowerInvariant()
            if (-not $blob.Contains($searchNormalized)) {
                continue
            }
        }

        [pscustomobject]@{
            station = $station
            score   = Get-RecommendationScore -Station $station -MoodKey $MoodKey -GenreKey $GenreKey -PreferTurkish $TurkishOnly -DiscoveryMode $DiscoveryMode
        }
    }

    $ranked = $scored |
        Sort-Object -Property @{
            Expression = {
                if ($_.PSObject.Properties["score"]) { [int]$_.score } else { 0 }
            }; Descending = $true
        }, @{
            Expression = {
                if ($_.PSObject.Properties["station"] -and $_.station -and $_.station.PSObject.Properties["name"]) {
                    [string]$_.station.name
                } elseif ($_.PSObject.Properties["name"]) {
                    [string]$_.name
                } else {
                    ""
                }
            }; Descending = $false
        }

    $result = New-Object "System.Collections.Generic.List[object]"
    foreach ($entry in @($ranked)) {
        if ($null -eq $entry) {
            continue
        }

        if ($entry.PSObject.Properties["station"] -and $entry.station) {
            [void]$result.Add($entry.station)
            continue
        }

        if ($entry.PSObject.Properties["url"]) {
            [void]$result.Add($entry)
        }
    }

    return $result.ToArray()
}

function Get-SurpriseStation {
    param(
        [string]$MoodKey = "all",
        [string]$GenreKey = "all",
        [bool]$TurkishOnly = $false,
        [bool]$PublicOnly = $true,
        [string]$DiscoveryMode = "surprise"
    )

    $ranked = Get-RankedStations -MoodKey $MoodKey -GenreKey $GenreKey -TurkishOnly $TurkishOnly -PublicOnly $PublicOnly -DiscoveryMode $DiscoveryMode
    if (@($ranked).Count -eq 0) {
        $ranked = Get-RankedStations -MoodKey "all" -GenreKey "all" -TurkishOnly $false -PublicOnly $PublicOnly -DiscoveryMode $DiscoveryMode
    }

    if (@($ranked).Count -eq 0) {
        return $null
    }

    $excludedIds = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    if ($script:currentStation -and $script:currentStation.id) {
        [void]$excludedIds.Add([string]$script:currentStation.id)
    }
    if ($script:profile.lastStationId) {
        [void]$excludedIds.Add([string]$script:profile.lastStationId)
    }

    $poolSource = $ranked
    $freshPool = $ranked | Where-Object { -not $excludedIds.Contains([string]$_.id) }
    if (@($freshPool).Count -gt 0) {
        $poolSource = @($freshPool)
    }

    $topN = [Math]::Min(12, @($poolSource).Count)
    $pool = $poolSource | Select-Object -First $topN
    return Get-Random -InputObject $pool
}

function Set-ComboSource {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ComboBox]$Combo,
        [Parameter(Mandatory = $true)]
        [hashtable]$Map,
        [string]$SelectedKey = "all"
    )

    $items = New-Object "System.Collections.Generic.List[object]"
    foreach ($key in $Map.Keys) {
        $items.Add([pscustomobject]@{
                Key   = $key
                Label = $Map[$key]
            })
    }

    $Combo.DisplayMember = "Label"
    $Combo.ValueMember = "Key"
    $Combo.DataSource = $items

    if ($Map.ContainsKey($SelectedKey)) {
        $Combo.SelectedValue = $SelectedKey
    } else {
        $Combo.SelectedIndex = 0
    }
}

function Get-ComboKey {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ComboBox]$Combo
    )

    if ($Combo.SelectedItem -and $Combo.SelectedItem.PSObject.Properties["Key"]) {
        return [string]$Combo.SelectedItem.Key
    }

    return "all"
}

function Join-Preview {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items,
        [int]$MaxItems = 2
    )

    if ($Items.Count -eq 0) {
        return "-"
    }

    $slice = $Items | Select-Object -First $MaxItems
    $text = ($slice -join ", ")
    if ($Items.Count -gt $MaxItems) {
        $text += "..."
    }
    return $text
}

function Show-OnboardingWizard {
    param(
        [switch]$Force
    )

    if (-not $Force -and [bool]$script:profile.hasOnboarded) {
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$script:AppName - Quick Setup"
    $form.Size = New-Object System.Drawing.Size(560, 460)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Let's personalize this (30 seconds)"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Pick mood + genre, and we'll recommend stations."
    $subtitle.AutoSize = $true
    $subtitle.Location = New-Object System.Drawing.Point(22, 55)
    $form.Controls.Add($subtitle)

    $lblMood = New-Object System.Windows.Forms.Label
    $lblMood.Text = "1) What's your current mood?"
    $lblMood.AutoSize = $true
    $lblMood.Location = New-Object System.Drawing.Point(22, 95)
    $form.Controls.Add($lblMood)

    $cmbMood = New-Object System.Windows.Forms.ComboBox
    $cmbMood.DropDownStyle = "DropDownList"
    $cmbMood.Location = New-Object System.Drawing.Point(25, 118)
    $cmbMood.Size = New-Object System.Drawing.Size(240, 28)
    Set-ComboSource -Combo $cmbMood -Map $script:MoodMap -SelectedKey ([string]$script:profile.preferredMood)
    $form.Controls.Add($cmbMood)

    $lblGenre = New-Object System.Windows.Forms.Label
    $lblGenre.Text = "2) Which genre do you want?"
    $lblGenre.AutoSize = $true
    $lblGenre.Location = New-Object System.Drawing.Point(22, 160)
    $form.Controls.Add($lblGenre)

    $cmbGenre = New-Object System.Windows.Forms.ComboBox
    $cmbGenre.DropDownStyle = "DropDownList"
    $cmbGenre.Location = New-Object System.Drawing.Point(25, 183)
    $cmbGenre.Size = New-Object System.Drawing.Size(240, 28)
    Set-ComboSource -Combo $cmbGenre -Map $script:GenreMap -SelectedKey ([string]$script:profile.preferredGenre)
    $form.Controls.Add($cmbGenre)

    $lblDiscovery = New-Object System.Windows.Forms.Label
    $lblDiscovery.Text = "3) Discovery style?"
    $lblDiscovery.AutoSize = $true
    $lblDiscovery.Location = New-Object System.Drawing.Point(22, 225)
    $form.Controls.Add($lblDiscovery)

    $cmbDiscovery = New-Object System.Windows.Forms.ComboBox
    $cmbDiscovery.DropDownStyle = "DropDownList"
    $cmbDiscovery.Location = New-Object System.Drawing.Point(25, 248)
    $cmbDiscovery.Size = New-Object System.Drawing.Size(240, 28)
    Set-ComboSource -Combo $cmbDiscovery -Map $script:DiscoveryMap -SelectedKey ([string]$script:profile.discoveryMode)
    $form.Controls.Add($cmbDiscovery)

    $chkTurkish = New-Object System.Windows.Forms.CheckBox
    $chkTurkish.Text = "Prefer Turkish stations"
    $chkTurkish.Checked = [bool]$script:profile.preferTurkish
    $chkTurkish.AutoSize = $true
    $chkTurkish.Location = New-Object System.Drawing.Point(310, 118)
    $form.Controls.Add($chkTurkish)

    $chkPublic = New-Object System.Windows.Forms.CheckBox
    $chkPublic.Text = "Public stations only"
    $chkPublic.Checked = [bool]$script:profile.publicOnly
    $chkPublic.AutoSize = $true
    $chkPublic.Location = New-Object System.Drawing.Point(310, 148)
    $form.Controls.Add($chkPublic)

    $lblVol = New-Object System.Windows.Forms.Label
    $lblVol.Text = "Startup volume"
    $lblVol.AutoSize = $true
    $lblVol.Location = New-Object System.Drawing.Point(310, 190)
    $form.Controls.Add($lblVol)

    $trackVol = New-Object System.Windows.Forms.TrackBar
    $trackVol.Location = New-Object System.Drawing.Point(310, 210)
    $trackVol.Size = New-Object System.Drawing.Size(220, 45)
    $trackVol.Minimum = 10
    $trackVol.Maximum = 100
    $trackVol.TickFrequency = 10
    $trackVol.Value = [Math]::Max([Math]::Min([int]$script:profile.volume, 100), 10)
    $form.Controls.Add($trackVol)

    $volValue = New-Object System.Windows.Forms.Label
    $volValue.Text = "$($trackVol.Value)%"
    $volValue.AutoSize = $true
    $volValue.Location = New-Object System.Drawing.Point(500, 190)
    $form.Controls.Add($volValue)
    $trackVol.Add_ValueChanged({
            $volValue.Text = "$($trackVol.Value)%"
        })

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Tip: Double-click tray icon for station list, or use Ctrl+Alt+B to quickly toggle playback."
    $hint.Size = New-Object System.Drawing.Size(510, 45)
    $hint.Location = New-Object System.Drawing.Point(22, 300)
    $form.Controls.Add($hint)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Start"
    $btnSave.Location = New-Object System.Drawing.Point(330, 360)
    $btnSave.Size = New-Object System.Drawing.Size(95, 32)
    $btnSave.Add_Click({
            $script:profile.preferredMood = Get-ComboKey -Combo $cmbMood
            $script:profile.preferredGenre = Get-ComboKey -Combo $cmbGenre
            $script:profile.discoveryMode = Get-ComboKey -Combo $cmbDiscovery
            $script:profile.preferTurkish = $chkTurkish.Checked
            $script:profile.publicOnly = $chkPublic.Checked
            $script:profile.volume = [int]$trackVol.Value
            if ($script:profile.PSObject.Properties["sessionVolume"]) {
                $script:profile.sessionVolume = [int]$trackVol.Value
            } else {
                $script:profile | Add-Member -NotePropertyName sessionVolume -NotePropertyValue ([int]$trackVol.Value)
            }
            $script:profile.hasOnboarded = $true
            Save-Profile

            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        })
    $form.Controls.Add($btnSave)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text = "Skip"
    $btnSkip.Location = New-Object System.Drawing.Point(435, 360)
    $btnSkip.Size = New-Object System.Drawing.Size(95, 32)
    $btnSkip.Add_Click({
            $script:profile.hasOnboarded = $true
            Save-Profile
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        })
    $form.Controls.Add($btnSkip)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            if (-not (Start-RecommendationFromProfile -PreferDifferentStation)) {
                Show-StationPicker
            }
        } catch {
            Show-Balloon -Title $script:AppName -Text ("Setup action failed: {0}" -f $_.Exception.Message)
        }
    }
}

function Start-RecommendationFromProfile {
    param(
        [switch]$PreferDifferentStation
    )

    $mood = [string]$script:profile.preferredMood
    $genre = [string]$script:profile.preferredGenre
    $turkish = [bool]$script:profile.preferTurkish
    $publicOnly = [bool]$script:profile.publicOnly
    $mode = [string]$script:profile.discoveryMode

    try {
        $ranked = Get-RankedStations -MoodKey $mood -GenreKey $genre -TurkishOnly $turkish -PublicOnly $publicOnly -DiscoveryMode $mode
        if (@($ranked).Count -eq 0 -and $genre -ne "all") {
            $ranked = Get-RankedStations -MoodKey "all" -GenreKey $genre -TurkishOnly $turkish -PublicOnly $publicOnly -DiscoveryMode $mode
        }
        if (@($ranked).Count -eq 0 -and $mood -ne "all") {
            $ranked = Get-RankedStations -MoodKey $mood -GenreKey "all" -TurkishOnly $turkish -PublicOnly $publicOnly -DiscoveryMode $mode
        }
        if (@($ranked).Count -eq 0) {
            $ranked = Get-RankedStations -MoodKey "all" -GenreKey "all" -TurkishOnly $turkish -PublicOnly $publicOnly -DiscoveryMode $mode
        }

        if (@($ranked).Count -eq 0) {
            return $false
        }

        $pick = $ranked[0]
        if ($mode -eq "new") {
            $options = $ranked | Where-Object { $_.id -ne $script:profile.lastStationId } | Select-Object -First 6
            if (@($options).Count -gt 0) {
                $pick = Get-Random -InputObject @($options)
            }
        } elseif ($mode -eq "surprise") {
            $pick = Get-SurpriseStation -MoodKey $mood -GenreKey $genre -TurkishOnly $turkish -PublicOnly $publicOnly -DiscoveryMode $mode
        }

        if ($PreferDifferentStation -and $script:currentStation -and @($ranked).Count -gt 1) {
            $alternate = $ranked | Where-Object { $_.id -ne $script:currentStation.id } | Select-Object -First 1
            if ($alternate) {
                $pick = $alternate
            }
        }

        if ($null -eq $pick) {
            return $false
        }

        Start-Station -Station $pick
        return $true
    } catch {
        Show-Balloon -Title $script:AppName -Text ("Recommendation failed: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Toggle-Playback {
    try {
        if (Get-IsPlaying) {
            Stop-Playback
            return
        }

        if ($script:profile.lastStationId) {
            $last = Get-StationById -Id ([string]$script:profile.lastStationId)
            if ($last) {
                Start-Station -Station $last
                return
            }
        }

        if (-not (Start-RecommendationFromProfile)) {
            Show-StationPicker
        }
    } catch {
        Show-Balloon -Title $script:AppName -Text ("Playback action failed: {0}" -f $_.Exception.Message)
    }
}

function Show-StationPicker {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$script:AppName - Station Picker"
    $form.Size = New-Object System.Drawing.Size(860, 560)
    $form.StartPosition = "CenterScreen"

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search:"
    $lblSearch.AutoSize = $true
    $lblSearch.Location = New-Object System.Drawing.Point(20, 18)
    $form.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(55, 14)
    $txtSearch.Size = New-Object System.Drawing.Size(250, 27)
    $form.Controls.Add($txtSearch)

    $cmbMood = New-Object System.Windows.Forms.ComboBox
    $cmbMood.DropDownStyle = "DropDownList"
    $cmbMood.Location = New-Object System.Drawing.Point(320, 14)
    $cmbMood.Size = New-Object System.Drawing.Size(160, 27)
    Set-ComboSource -Combo $cmbMood -Map $script:MoodMap -SelectedKey ([string]$script:profile.preferredMood)
    $form.Controls.Add($cmbMood)

    $cmbGenre = New-Object System.Windows.Forms.ComboBox
    $cmbGenre.DropDownStyle = "DropDownList"
    $cmbGenre.Location = New-Object System.Drawing.Point(490, 14)
    $cmbGenre.Size = New-Object System.Drawing.Size(160, 27)
    Set-ComboSource -Combo $cmbGenre -Map $script:GenreMap -SelectedKey ([string]$script:profile.preferredGenre)
    $form.Controls.Add($cmbGenre)

    $chkTurkish = New-Object System.Windows.Forms.CheckBox
    $chkTurkish.Text = "Turkish"
    $chkTurkish.Checked = [bool]$script:profile.preferTurkish
    $chkTurkish.AutoSize = $true
    $chkTurkish.Location = New-Object System.Drawing.Point(665, 17)
    $form.Controls.Add($chkTurkish)

    $chkPublic = New-Object System.Windows.Forms.CheckBox
    $chkPublic.Text = "Public"
    $chkPublic.Checked = [bool]$script:profile.publicOnly
    $chkPublic.AutoSize = $true
    $chkPublic.Location = New-Object System.Drawing.Point(740, 17)
    $form.Controls.Add($chkPublic)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20, 55)
    $listView.Size = New-Object System.Drawing.Size(810, 410)
    $listView.View = "Details"
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.MultiSelect = $false
    $listView.HideSelection = $false
    [void]$listView.Columns.Add("Station", 250)
    [void]$listView.Columns.Add("Genre", 210)
    [void]$listView.Columns.Add("Mood", 180)
    [void]$listView.Columns.Add("Country", 80)
    [void]$listView.Columns.Add("Type", 70)
    $form.Controls.Add($listView)

    $btnPlay = New-Object System.Windows.Forms.Button
    $btnPlay.Text = "Play"
    $btnPlay.Location = New-Object System.Drawing.Point(520, 480)
    $btnPlay.Size = New-Object System.Drawing.Size(75, 32)
    $form.Controls.Add($btnPlay)

    $btnSurprise = New-Object System.Windows.Forms.Button
    $btnSurprise.Text = "Surprise"
    $btnSurprise.Location = New-Object System.Drawing.Point(605, 480)
    $btnSurprise.Size = New-Object System.Drawing.Size(75, 32)
    $form.Controls.Add($btnSurprise)

    $btnStop = New-Object System.Windows.Forms.Button
    $btnStop.Text = "Stop"
    $btnStop.Location = New-Object System.Drawing.Point(690, 480)
    $btnStop.Size = New-Object System.Drawing.Size(65, 32)
    $btnStop.Add_Click({ Stop-Playback })
    $form.Controls.Add($btnStop)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(765, 480)
    $btnClose.Size = New-Object System.Drawing.Size(65, 32)
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    $selection = @{
        Station = $null
    }

    $refreshList = {
        $moodKey = Get-ComboKey -Combo $cmbMood
        $genreKey = Get-ComboKey -Combo $cmbGenre
        $stations = Get-RankedStations `
            -SearchText $txtSearch.Text `
            -MoodKey $moodKey `
            -GenreKey $genreKey `
            -TurkishOnly $chkTurkish.Checked `
            -PublicOnly $chkPublic.Checked `
            -DiscoveryMode ([string]$script:profile.discoveryMode)

        $listView.BeginUpdate()
        $listView.Items.Clear()

        foreach ($station in $stations) {
            $genres = Join-Preview -Items (To-Array -Value $station.genres) -MaxItems 3
            $moods = Join-Preview -Items (To-Array -Value $station.moods) -MaxItems 2
            $typeLabel = if ([bool]$station.turkish) { "TR" } else { "Global" }

            $item = New-Object System.Windows.Forms.ListViewItem([string]$station.name)
            [void]$item.SubItems.Add($genres)
            [void]$item.SubItems.Add($moods)
            [void]$item.SubItems.Add([string]$station.country)
            [void]$item.SubItems.Add($typeLabel)
            $item.Tag = $station
            [void]$listView.Items.Add($item)
        }

        $listView.EndUpdate()
        if ($listView.Items.Count -gt 0) {
            $listView.Items[0].Selected = $true
        }
    }

    $playSelected = {
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a station.", $script:AppName, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }

        $selection.Station = $listView.SelectedItems[0].Tag
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    $txtSearch.Add_TextChanged($refreshList)
    $cmbMood.Add_SelectedIndexChanged($refreshList)
    $cmbGenre.Add_SelectedIndexChanged($refreshList)
    $chkTurkish.Add_CheckedChanged($refreshList)
    $chkPublic.Add_CheckedChanged($refreshList)
    $btnPlay.Add_Click($playSelected)
    $listView.Add_DoubleClick($playSelected)

    $btnSurprise.Add_Click({
            $moodKey = Get-ComboKey -Combo $cmbMood
            $genreKey = Get-ComboKey -Combo $cmbGenre
            $selection.Station = Get-SurpriseStation `
                -MoodKey $moodKey `
                -GenreKey $genreKey `
                -TurkishOnly $chkTurkish.Checked `
                -PublicOnly $chkPublic.Checked `
                -DiscoveryMode "surprise"

            if ($null -eq $selection.Station) {
                [System.Windows.Forms.MessageBox]::Show("No station found with current filters.", $script:AppName, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        })

    & $refreshList
    $result = $form.ShowDialog()

    $script:profile.preferredMood = Get-ComboKey -Combo $cmbMood
    $script:profile.preferredGenre = Get-ComboKey -Combo $cmbGenre
    $script:profile.preferTurkish = $chkTurkish.Checked
    $script:profile.publicOnly = $chkPublic.Checked
    Save-Profile

    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $selection.Station) {
        Start-Station -Station $selection.Station
    }
}

function Register-GlobalHotkey {
    $script:hotkeyWindow = New-Object RofiHotkeyWindow
    $script:hotkeyWindow.ShowInTaskbar = $false
    $script:hotkeyWindow.FormBorderStyle = "FixedToolWindow"
    $script:hotkeyWindow.StartPosition = "Manual"
    $script:hotkeyWindow.Location = New-Object System.Drawing.Point(-32000, -32000)
    $script:hotkeyWindow.Size = New-Object System.Drawing.Size(1, 1)
    $script:hotkeyWindow.Opacity = 0

    $null = $script:hotkeyWindow.Show()
    $script:hotkeyWindow.Hide()

    $mods = 0x0001 -bor 0x0002
    $registered = [RofiHotkeyWindow]::RegisterHotKey($script:hotkeyWindow.Handle, $script:hotkeyId, $mods, 0x42)
    if (-not $registered) {
        Show-Balloon -Title $script:AppName -Text "Ctrl+Alt+B registration failed. Shortcut may be used by another app."
    }

    $script:hotkeyHandler = [System.EventHandler]{
        param($sender, $eventArgs)
        Toggle-Playback
    }
    $script:hotkeyWindow.add_HotkeyPressed($script:hotkeyHandler)
}

function New-MenuItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [scriptblock]$OnClick
    )

    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = $Text
    if ($OnClick) {
        $item.Add_Click($OnClick)
    }
    return $item
}

function Exit-App {
    if ($script:isExiting) {
        return
    }

    $script:isExiting = $true

    try {
        Stop-Playback -Silent
    } catch {
    }

    try {
        Stop-MetadataTimer
    } catch {
    }

    try {
        Stop-AudioSessionTimer
    } catch {
    }

    try {
        if ($script:hotkeyWindow) {
            if ($script:hotkeyHandler) {
                try {
                    $script:hotkeyWindow.remove_HotkeyPressed($script:hotkeyHandler)
                } catch {
                }
                $script:hotkeyHandler = $null
            }
            [RofiHotkeyWindow]::UnregisterHotKey($script:hotkeyWindow.Handle, $script:hotkeyId) | Out-Null
            $script:hotkeyWindow.Close()
            $script:hotkeyWindow.Dispose()
        }
    } catch {
    }

    try {
        if ($script:notifyIcon) {
            $script:notifyIcon.Visible = $false
            $script:notifyIcon.Dispose()
        }
    } catch {
    }

    try {
        if ($script:trayIcon) {
            $script:trayIcon.Dispose()
            $script:trayIcon = $null
        }
    } catch {
    }

    [System.Windows.Forms.Application]::ExitThread()
}

function Initialize-Tray {
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

    $script:statusItem = New-MenuItem -Text "Status: Idle"
    $script:statusItem.Enabled = $false
    $script:currentItem = New-MenuItem -Text "Now: Off"
    $script:currentItem.Enabled = $false
    $script:songItem = New-MenuItem -Text "Song: -"
    $script:songItem.Enabled = $false
    $script:bitrateItem = New-MenuItem -Text "Bitrate: -"
    $script:bitrateItem.Enabled = $false
    $script:volumeItem = New-MenuItem -Text "Volume: 100%"
    $script:volumeItem.Enabled = $false
    $script:toggleItem = New-MenuItem -Text "Play (Ctrl+Alt+B)" -OnClick { Toggle-Playback }
    $pickItem = New-MenuItem -Text "Choose station..." -OnClick { Show-StationPicker }
    $surpriseItem = New-MenuItem -Text "Surprise station" -OnClick {
        try {
            $pick = Get-SurpriseStation -MoodKey ([string]$script:profile.preferredMood) -GenreKey ([string]$script:profile.preferredGenre) -TurkishOnly ([bool]$script:profile.preferTurkish) -PublicOnly ([bool]$script:profile.publicOnly) -DiscoveryMode "surprise"
            if ($pick) {
                Start-Station -Station $pick
            }
        } catch {
            Show-Balloon -Title $script:AppName -Text ("Surprise action failed: {0}" -f $_.Exception.Message)
        }
    }
    $onboardingItem = New-MenuItem -Text "Open setup wizard" -OnClick { Show-OnboardingWizard -Force }

    $moodMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $moodMenu.Text = "Quick mood"
    foreach ($moodKey in $script:MoodMap.Keys) {
        if ($moodKey -eq "all") {
            continue
        }

        $moodItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $moodItem.Text = [string]$script:MoodMap[$moodKey]
        $moodItem.Tag = $moodKey
        $moodItem.Add_Click({
                param($sender, $eventArgs)
                try {
                    $script:profile.preferredMood = [string]$sender.Tag
                    Save-Profile
                    Start-RecommendationFromProfile | Out-Null
                } catch {
                    Show-Balloon -Title $script:AppName -Text ("Quick mood failed: {0}" -f $_.Exception.Message)
                }
            })
        [void]$moodMenu.DropDownItems.Add($moodItem)
    }

    $genreMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $genreMenu.Text = "Quick genre"
    foreach ($genreKey in $script:GenreMap.Keys) {
        if ($genreKey -eq "all") {
            continue
        }

        $genreItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $genreItem.Text = [string]$script:GenreMap[$genreKey]
        $genreItem.Tag = $genreKey
        $genreItem.Add_Click({
                param($sender, $eventArgs)
                try {
                    $script:profile.preferredGenre = [string]$sender.Tag
                    Save-Profile
                    Start-RecommendationFromProfile | Out-Null
                } catch {
                    Show-Balloon -Title $script:AppName -Text ("Quick genre failed: {0}" -f $_.Exception.Message)
                }
            })
        [void]$genreMenu.DropDownItems.Add($genreItem)
    }

    $script:volumeUpItem = New-MenuItem -Text "Volume +" -OnClick { Adjust-PlayerSessionVolume -Delta 5 }
    $script:volumeDownItem = New-MenuItem -Text "Volume -" -OnClick { Adjust-PlayerSessionVolume -Delta -5 }
    $script:muteItem = New-MenuItem -Text "Mute" -OnClick { Toggle-PlayerSessionMute }
    $stopItem = New-MenuItem -Text "Stop playback" -OnClick { Stop-Playback }
    $exitItem = New-MenuItem -Text "Exit (stops music)" -OnClick { Exit-App }

    [void]$contextMenu.Items.Add($script:statusItem)
    [void]$contextMenu.Items.Add($script:currentItem)
    [void]$contextMenu.Items.Add($script:songItem)
    [void]$contextMenu.Items.Add($script:bitrateItem)
    [void]$contextMenu.Items.Add($script:volumeItem)
    [void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$contextMenu.Items.Add($script:toggleItem)
    [void]$contextMenu.Items.Add($pickItem)
    [void]$contextMenu.Items.Add($surpriseItem)
    [void]$contextMenu.Items.Add($script:volumeUpItem)
    [void]$contextMenu.Items.Add($script:volumeDownItem)
    [void]$contextMenu.Items.Add($script:muteItem)
    [void]$contextMenu.Items.Add($moodMenu)
    [void]$contextMenu.Items.Add($genreMenu)
    [void]$contextMenu.Items.Add($onboardingItem)
    [void]$contextMenu.Items.Add($stopItem)
    [void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$contextMenu.Items.Add($exitItem)

    $script:notifyIcon = New-Object System.Windows.Forms.NotifyIcon

    $trayCandidate = New-MusicTrayIcon
    $resolvedTrayIcon = $null
    if ($trayCandidate -is [System.Array]) {
        $resolvedTrayIcon = $trayCandidate | Where-Object { $_ -is [System.Drawing.Icon] } | Select-Object -First 1
    } elseif ($trayCandidate -is [System.Drawing.Icon]) {
        $resolvedTrayIcon = $trayCandidate
    }

    if (-not $resolvedTrayIcon) {
        $resolvedTrayIcon = [System.Drawing.SystemIcons]::Information
    }

    $script:trayIcon = $resolvedTrayIcon
    $script:notifyIcon.Icon = $resolvedTrayIcon
    $script:notifyIcon.Visible = $true
    $script:notifyIcon.ContextMenuStrip = $contextMenu
    Set-NotifyText -Text $script:AppName
    $script:notifyIcon.Add_DoubleClick({
            Show-StationPicker
        })
}

try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    try {
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
    } catch {
        # Some hosts initialize WinForms internals early; continue with default rendering settings.
    }

    Load-Stations
    Initialize-Profile
    if (-not (Ensure-MpvAvailable)) {
        return
    }

    Try-RestoreState
    Initialize-Tray
    Register-GlobalHotkey
    Update-TrayStatus

    if (-not [bool]$script:profile.hasOnboarded) {
        Show-OnboardingWizard
    }

    Show-Balloon -Title $script:AppName -Text "Ready. Use Ctrl+Alt+B to toggle, double-click tray icon for station list."
    [System.Windows.Forms.Application]::Run()
} catch {
    if (-not $script:isExiting) {
        $lineInfo = ""
        if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) {
            $lineInfo = " (line $($_.InvocationInfo.ScriptLineNumber))"
        }
        [System.Windows.Forms.MessageBox]::Show(
            "$script:AppName failed to start:`n$($_.Exception.GetType().Name): $($_.Exception.Message)$lineInfo",
            $script:AppName,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
} finally {
    Exit-App
}
