<#
.SYNOPSIS
    Robust Hybrid File Manager 2.0 - a menu-driven file manager for SharePoint Online,
    OneDrive and local / network folders, built for Windows PowerShell 5.1.

.DESCRIPTION
    Copy, move, download, upload, verify, search, wipe and archive files between
    SharePoint / OneDrive and your PC or network shares - all from one friendly menu.

    Authentication is unchanged from earlier versions:
      * SharePoint / OneDrive : PnP.PowerShell 1.12.0 with Connect-PnPOnline -UseWebLogin
      * Option [23] (licences) : Microsoft Graph REST with the device-code sign-in flow and a
                                  DPAPI-encrypted refresh-token cache (no Graph module needed)

    The PnP module is only loaded (and installed if missing) the first time you pick a
    cloud option, so local-only work starts instantly.

.PARAMETER WorkDir
    Folder for logs, favourites, settings and job history.

.PARAMETER MaxLogSizeMB
    The log file is rotated when it grows past this size.

.PARAMETER CloudPathLimit
    Paths longer than this many characters are skipped and listed in LongPaths_*.txt
    (SharePoint / OneDrive sync limit safety margin).

.PARAMETER Ascii
    Draw boxes with plain ASCII characters (+ - |) for very old consoles or fonts.
    You can also switch this permanently in [S] Settings.

.PARAMETER LibraryMode
    Load every function without starting the menu (for automation or testing).

.EXAMPLE
    .\Cross-Tenant-Cloud-FileManager.ps1

.EXAMPLE
    .\Cross-Tenant-Cloud-FileManager.ps1 -WorkDir D:\FileManagerWork -Ascii

.NOTES
    QUICK START
      1. Right-click the file > "Run with PowerShell", or from PowerShell:
           Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
           .\Cross-Tenant-Cloud-FileManager.ps1
      2. Pick an option by number (or type a word such as "download"), or press W for the
         guided mode that asks two simple questions and picks the right option for you.
      3. At any path prompt you can paste, drag a folder from Explorer into the window,
         press [B] to browse, [F] for Favourites or [R] for recently used paths.
      4. During a job press [P] to pause or [C] (or Ctrl+C) to cancel cleanly.

    REQUIREMENTS
      * Windows PowerShell 5.1 (built into Windows 10 / 11 / Server 2016+)
      * PnP.PowerShell 1.12.0 (installed automatically the first time a cloud option is used)
      * A Microsoft 365 account with access to the SharePoint sites / OneDrives involved

    FILES (in WorkDir)
      Logs\HybridJobLog_<user>.log      full audit log (rotates at MaxLogSizeMB)
      Favorites_<user>.json             saved favourites (same format as v80-v111)
      Settings_<user>.json              display / behaviour settings
      JobHistory_<user>.json            the last 500 completed jobs (for [J] history)
      MgTokenCache_<user>.dat           DPAPI-encrypted Graph sign-in for option [23]

    VERSION HISTORY
      v1 - v111  Incremental releases (see the git history of this file for the full,
                 original changelog). Highlights: PnP transfers, long-path kernel bypass,
                 batch menus, verification engines, favourites, UNC support, upload and
                 archive engines, leaver OneDrive archiving, licence-group archiving.

      2.0  COMPLETE OVERHAUL (option numbers 1-23 and authentication are unchanged)
        Look and feel
          * New display toolkit: every box, table and panel sizes itself to the console
            width, long paths wrap instead of breaking the borders, and the source code is
            pure ASCII so the script can never be corrupted by file-encoding changes.
          * Glyphs limited to characters the Windows PowerShell 5.1 console fonts can draw
            (no more empty squares instead of tick marks); optional ASCII theme.
          * Redesigned two-column main menu, guided mode [W], keyword search in the menu,
            job history [J], settings [S], session scoreboard, window-title progress,
            optional completion chime and friendlier, shorter explanations (can be hidden).
        Ease of use
          * Path prompts accept drag-and-drop, quoted paths, [B] Browse dialog, [F]
            Favourites and [R] Recent; invalid input re-prompts instead of failing later.
          * One shared item picker for every batch screen (3 | 2-5 | 1,3,7 | 1 3 7 | A).
          * Cloud connections are reused between operations on the same site, so you sign
            in far less often (same -UseWebLogin method).
          * Ctrl+C cancels a running job cleanly instead of killing the script.
          * Copy-to-clipboard and "show in Explorer" (no longer runs the file) in searches.
          * Local verify can optionally compare SHA-256 hashes; differences export to CSV.
        Reliability and safety fixes
          * Archive verification now checks the files on DISK against a fresh, strict cloud
            inventory. Folder-listing failures, name collisions, cancelled runs, missing or
            wrong-size files all block deletion (previously some of these could still
            delete cloud data). Single-file archives delete with Remove-PnPFile, library
            roots are never deleted (their contents are), and every delete asks once.
          * Cancel now stops the whole run (all items, libraries and leavers).
          * Leaver runs report real outcomes instead of always "ARCHIVED".
          * Archive resumes compare dates as well as sizes, and an archive never overwrites a
            different file already in the archive folder (an identical file left by an older
            version is recognised by checksum). Only items that were verified are recycled.
          * Moves and deletes check that two paths are not the same file or folder on disk
            (mapped drive vs UNC path, server aliases, junctions) before removing anything.
          * Library roots (also on sub-sites) are never deleted; a user folder that happens
            to be called "Forms" is no longer mistaken for the hidden system folder.
          * Leaver libraries other than Documents are saved as "<Library> (library)" so they
            can never mix with a Documents folder of the same name on a re-run.
          * Option [17] no longer crashes after the download step.
          * Option [23] now really removes users from the licence group (the old code
            built an invalid URL), checks membership before and after so a retried request
            can never leave someone licensed, retries the removal, cleans up when cancelled,
            handles slow_down, shows (and lets you switch) the signed-in account, falls back
            to a mail lookup and lets you pick among groups with the same name.
          * Leaver options only accept OneDrive (/personal/) addresses, so a team site can
            never be archived by mistake.
          * Favourites no longer merge into one corrupted entry on load in PS 5.1.
          * Free-space checks work on UNC paths in PS 5.1 (the [ulong] type does not exist
            there, so the old check always failed).
          * Move jobs never delete folders that still contain un-moved files; a move onto an
            existing different file copies first and deletes the source afterwards.
          * Folder enumeration skips unreadable folders and REPORTS them instead of silently
            missing files; no more duplicated items after a partial scan.
          * Cloud listing / download / upload calls retry with throttle-aware back-off,
            and pause / cancel works during the waits.
          * Downloads check the received size; resumed downloads keep cloud dates.
          * Upload skip checks use one folder listing instead of one request per file.
          * URL parsing handles /teams/ sites, folder names containing '&', view pages
            (Forms/AllItems.aspx), viewid= links and OneDrive page links.
          * Drive roots like C:\ are no longer turned into "C:" (current directory).
          * Search no longer breaks on [ ] in names; *.xls no longer matches .xlsx.
          * Many smaller fixes: duration rounding, stray "+" on screen, literal \u2014
            text, misaligned boxes, wrong option cross-references, TLS 1.2 for Graph.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$WorkDir        = "C:\Users\Public\Documents\RobustCloudWork",
    [int]   $MaxLogSizeMB   = 5,
    [int]   $CloudPathLimit = 400,
    [switch]$Ascii,
    [switch]$LibraryMode
)

# ======================================================================================
# SECTION 1 - START-UP
# ======================================================================================

Set-StrictMode -Off
$script:AppName    = 'Robust Hybrid File Manager'
$script:AppVersion = '2.0'
$script:IsWin      = ($PSVersionTable.PSVersion.Major -le 5) -or ($env:OS -eq 'Windows_NT')
$env:PNPPOWERSHELL_UPDATECHECK = 'Off'

# Microsoft sign-in and Graph require TLS 1.2. Older .NET Framework defaults may not
# include it, which makes Invoke-RestMethod fail with "underlying connection was closed".
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

$ActiveUser = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { 'user' }

# Make sure the work folder exists. If the shared default is not writable (locked-down
# PCs), fall back to the user's own profile so the script still starts.
function Initialize-WorkFolder {
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path)) { $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop }
        $probe = Join-Path $Path ("write_probe_{0}.tmp" -f [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, 'ok')
        [System.IO.File]::Delete($probe)
        return $Path
    } catch {
        $base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } elseif ($env:HOME) { $env:HOME } else { [System.IO.Path]::GetTempPath() }
        $fallback = Join-Path $base 'RobustCloudWork'
        if (-not (Test-Path -LiteralPath $fallback)) { $null = New-Item -ItemType Directory -Path $fallback -Force }
        Write-Host "  Note: '$Path' is not writable - using '$fallback' for logs and settings." -ForegroundColor Yellow
        return $fallback
    }
}

$WorkDir          = Initialize-WorkFolder -Path $WorkDir
$LogFileDirectory = Join-Path $WorkDir 'Logs'
if (-not (Test-Path -LiteralPath $LogFileDirectory)) { $null = New-Item -ItemType Directory -Path $LogFileDirectory -Force }

$LogFile                 = Join-Path $LogFileDirectory "HybridJobLog_$ActiveUser.log"
$LongPathReport          = Join-Path $LogFileDirectory "LongPaths_FailedToSync_$ActiveUser.txt"
$FavoritesFile           = Join-Path $WorkDir "Favorites_$ActiveUser.json"
$script:SettingsFile     = Join-Path $WorkDir "Settings_$ActiveUser.json"
$script:HistoryFile      = Join-Path $WorkDir "JobHistory_$ActiveUser.json"
$script:MgTokenCacheFile = Join-Path $WorkDir "MgTokenCache_$ActiveUser.dat"

# Runtime state shared by every function in this script.
$script:JobCancelled     = $false
$script:SessionBytes     = [long]0
$script:PnPReady         = $false
$script:PnPSiteUrl       = $null
$script:Recent           = @{}
$script:OriginalTitle    = $null
$script:JobDepth         = 0
$script:Session          = @{ Jobs = 0; Files = 0; Failed = 0; Bytes = [long]0; Started = (Get-Date) }

# Microsoft Graph state for option [23] (pure REST, no module).
$script:MgGraphConnected = $false
$script:MgAccessToken    = ''
$script:MgRefreshToken   = ''
$script:MgTokenExpiry    = $null
$script:MgAccount        = ''
# Well-known Microsoft Graph PowerShell public client (device-code enabled, no app registration needed).
$script:MgClientId       = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$script:MgScopes         = 'GroupMember.ReadWrite.All User.Read.All Group.Read.All offline_access'

# Native helpers: free space for local AND UNC paths, and atomic replace-move for downloads.
# The class name is new so it never clashes with the 'Win32' type older versions compiled
# into the same PowerShell window.
if ($script:IsWin -and -not ('RHFMNative' -as [type])) {
    try {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class RHFMNative {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetDiskFreeSpaceExW(string lpDirectoryName, out ulong lpFreeBytesAvailable, out ulong lpTotalNumberOfBytes, out ulong lpTotalNumberOfFreeBytes);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MoveFileExW(string lpExistingFileName, string lpNewFileName, int dwFlags);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct FindData {
        public uint dwFileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string cAlternateFileName;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindFirstFileW(string lpFileName, out FindData lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool FindClose(IntPtr hFindFile);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr CreateFileW(string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool CloseHandle(IntPtr hObject);

    [StructLayout(LayoutKind.Sequential)]
    public struct ByHandleInfo {
        public uint FileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
        public uint VolumeSerialNumber;
        public uint FileSizeHigh;
        public uint FileSizeLow;
        public uint NumberOfLinks;
        public uint FileIndexHigh;
        public uint FileIndexLow;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetFileInformationByHandle(IntPtr hFile, out ByHandleInfo lpFileInformation);

    // Volume serial + file index: identical for two paths that reach the same file or
    // folder (mapped drive vs UNC, server alias, junction...). null when unavailable.
    public static string GetFileId(string path) {
        IntPtr handle = CreateFileW(path, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero);
        if (handle == new IntPtr(-1)) { return null; }
        try {
            ByHandleInfo info;
            if (!GetFileInformationByHandle(handle, out info)) { return null; }
            return info.VolumeSerialNumber.ToString("X8") + ":" + info.FileIndexHigh.ToString("X8") + info.FileIndexLow.ToString("X8");
        } finally { CloseHandle(handle); }
    }

    // Reparse tag of a file or folder (0 when it is not a reparse point).
    // 0xA0000003 = junction / mount point, 0xA000000C = symbolic link.
    public static uint GetReparseTag(string path) {
        FindData data;
        IntPtr handle = FindFirstFileW(path, out data);
        if (handle == new IntPtr(-1)) { return 0; }
        FindClose(handle);
        if ((data.dwFileAttributes & 0x400) == 0) { return 0; }
        return data.dwReserved0;
    }
}
"@ -ErrorAction Stop
    } catch {
        Write-Host "  Note: native helpers could not be compiled - using .NET fall-backs." -ForegroundColor Yellow
    }
}

# ======================================================================================
# SECTION 2 - DISPLAY TOOLKIT
# --------------------------------------------------------------------------------------
# Everything that draws on screen goes through these helpers so the whole app has one
# consistent look, never draws past the console width and works in both the Windows
# PowerShell 5.1 console and the ISE. Only box-drawing and block characters are used
# (they exist in every console font); there is also a plain ASCII theme.
# ======================================================================================

function Initialize-Theme {
    param([string]$Style = 'Unicode')
    if ($Style -eq 'Ascii') {
        $script:G = @{
            H = '-'; V = '|'; TL = '+'; TR = '+'; BL = '+'; BR = '+'; LT = '+'; RT = '+'
            DH = '='; DV = '|'; DTL = '+'; DTR = '+'; DBL = '+'; DBR = '+'; DLT = '+'; DRT = '+'
            Block = '#'; Shade = '.'
        }
    } else {
        $script:G = @{
            H   = [string][char]0x2500; V   = [string][char]0x2502
            TL  = [string][char]0x250C; TR  = [string][char]0x2510
            BL  = [string][char]0x2514; BR  = [string][char]0x2518
            LT  = [string][char]0x251C; RT  = [string][char]0x2524
            DH  = [string][char]0x2550; DV  = [string][char]0x2551
            DTL = [string][char]0x2554; DTR = [string][char]0x2557
            DBL = [string][char]0x255A; DBR = [string][char]0x255D
            DLT = [string][char]0x2560; DRT = [string][char]0x2563
            Block = [string][char]0x2588; Shade = [string][char]0x2591
        }
    }
    $script:ThemeStyle = $Style
}

# Usable drawing width: the console width minus one column (writing into the very last
# column makes the console wrap), capped so boxes stay readable on very wide windows.
function Get-UiWidth {
    $w = 0
    try { $w = [int]$Host.UI.RawUI.WindowSize.Width } catch { $w = 0 }
    if ($w -le 0) { try { $w = [int][Console]::WindowWidth } catch { $w = 0 } }
    if ($w -le 0) { $w = 100 }          # width unknown (some hosts) - assume a normal window
    return [Math]::Max(40, [Math]::Min($w - 1, 118))
}

# Shorten text to fit a width. -Middle keeps both ends (useful for paths).
function Get-FitText {
    param([string]$Text, [int]$Width, [switch]$Middle)
    if ($null -eq $Text) { $Text = '' }
    if ($Width -lt 5) { $Width = 5 }
    if ($Text.Length -le $Width) { return $Text }
    if ($Middle) {
        $keep  = $Width - 3
        $left  = [int][Math]::Ceiling($keep / 2)
        $right = $keep - $left
        return $Text.Substring(0, $left) + '...' + $Text.Substring($Text.Length - $right)
    }
    return $Text.Substring(0, $Width - 3) + '...'
}

# Word-wrap text to a width. Words longer than the width (paths, URLs) are hard-split so
# nothing ever runs past a border. Leading indentation of each paragraph is preserved.
function Split-TextToWidth {
    param([string]$Text, [int]$Width)
    if ($Width -lt 10) { $Width = 10 }
    $out = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrEmpty($Text)) { $out.Add(''); return ,$out.ToArray() }
    foreach ($para in ($Text -split "`r?`n")) {
        $indent = [regex]::Match($para, '^\s*').Value
        if ($indent.Length -gt ($Width / 2)) { $indent = '' }
        $avail = $Width - $indent.Length
        $body  = $para.Trim()
        if ($body.Length -eq 0) { $out.Add(''); continue }
        $line = ''
        foreach ($word in ($body -split ' ')) {
            if ($word.Length -eq 0) { continue }
            $cand = if ($line.Length -eq 0) { $word } else { "$line $word" }
            if ($cand.Length -le $avail) { $line = $cand; continue }
            if ($word.Length -le $avail) { $out.Add($indent + $line); $line = $word; continue }
            # A word longer than the line (a path or URL): continue it on this line and hard-split.
            $line = $cand
            while ($line.Length -gt $avail) { $out.Add($indent + $line.Substring(0, $avail)); $line = $line.Substring($avail) }
        }
        if ($line.Length -gt 0) { $out.Add($indent + $line) }
    }
    return ,$out.ToArray()
}

function Get-LevelColor {
    param([string]$Level)
    switch ($Level) {
        'SUCCESS' { return 'Green' }
        'ERROR'   { return 'Red' }
        'WARN'    { return 'Yellow' }
        'MUTED'   { return 'DarkGray' }
        default   { return 'White' }
    }
}

# A horizontal rule, optionally with a title: "  -- TITLE ------------------"
function Write-Rule {
    param([string]$Title = '', [string]$Color = 'DarkGray', [switch]$Double)
    $W = Get-UiWidth
    $h = if ($Double) { $script:G.DH } else { $script:G.H }
    if ($Title) {
        $t = Get-FitText -Text $Title -Width ($W - 10)
        $rest = [Math]::Max(3, $W - 2 - 4 - $t.Length)
        Write-Host ("  " + ($h * 2) + " ") -ForegroundColor $Color -NoNewline
        Write-Host $t -ForegroundColor White -NoNewline
        Write-Host (" " + ($h * $rest)) -ForegroundColor $Color
    } else {
        Write-Host ("  " + ($h * ($W - 2))) -ForegroundColor $Color
    }
}

# --------------------------------------------------------------------------------------
# Write-Panel: the workhorse box. Lines may be:
#   'text'                                    plain text (wrapped)
#   @{ Text='...'; Color='Yellow' }           coloured text (wrapped)
#   @{ Label='Files'; Value='12'; Level='SUCCESS' }   label ........ value row
#   '-'                                       divider
# --------------------------------------------------------------------------------------
function Write-Panel {
    param(
        [string]  $Title     = '',
        [object[]]$Lines     = @(),
        [string]  $Color     = 'Cyan',
        [string]  $TextColor = 'Gray',
        [switch]  $Double
    )
    $W     = Get-UiWidth
    $B     = $W - 2
    $inner = $B - 4
    if ($Double) { $h = $script:G.DH; $v = $script:G.DV; $tl = $script:G.DTL; $tr = $script:G.DTR; $bl = $script:G.DBL; $br = $script:G.DBR; $lt = $script:G.DLT; $rt = $script:G.DRT }
    else         { $h = $script:G.H;  $v = $script:G.V;  $tl = $script:G.TL;  $tr = $script:G.TR;  $bl = $script:G.BL;  $br = $script:G.BR;  $lt = $script:G.LT;  $rt = $script:G.RT }

    # Top border, with the title embedded.
    if ($Title) {
        $t = Get-FitText -Text $Title -Width ($B - 6)
        $fill = $B - 2 - 3 - $t.Length
        Write-Host ("  " + $tl + $h + " ") -ForegroundColor $Color -NoNewline
        Write-Host $t -ForegroundColor White -NoNewline
        Write-Host (" " + ($h * [Math]::Max(0, $fill)) + $tr) -ForegroundColor $Color
    } else {
        Write-Host ("  " + $tl + ($h * ($B - 2)) + $tr) -ForegroundColor $Color
    }

    $labelW = [Math]::Min(26, [int]($inner * 0.45))
    foreach ($ln in $Lines) {
        if ($null -eq $ln) { continue }
        if ($ln -is [string] -and ($ln -eq '-' -or $ln -eq '---')) {
            Write-Host ("  " + $lt + ($h * ($B - 2)) + $rt) -ForegroundColor $Color
            continue
        }
        if ($ln -is [hashtable] -and $ln.ContainsKey('Label')) {
            $label = Get-FitText -Text ([string]$ln.Label) -Width $labelW
            $value = [string]$ln.Value
            $vcol  = Get-LevelColor $ln.Level
            if (-not $ln.Level) { $vcol = 'White' }
            $leader = ' ' + ('.' * ($labelW - $label.Length)) + ' '
            $valW   = $inner - $labelW - 2
            $parts  = Split-TextToWidth -Text $value -Width $valW
            for ($i = 0; $i -lt $parts.Count; $i++) {
                Write-Host ("  " + $v + " ") -ForegroundColor $Color -NoNewline
                if ($i -eq 0) {
                    Write-Host $label -ForegroundColor Gray -NoNewline
                    Write-Host $leader -ForegroundColor DarkGray -NoNewline
                } else {
                    Write-Host (' ' * ($labelW + 2)) -NoNewline
                }
                Write-Host ($parts[$i].PadRight($valW)) -ForegroundColor $vcol -NoNewline
                Write-Host (" " + $v) -ForegroundColor $Color
            }
            continue
        }
        $text = $ln; $col = $TextColor
        if ($ln -is [hashtable]) { $text = [string]$ln.Text; if ($ln.Color) { $col = $ln.Color } }
        foreach ($part in (Split-TextToWidth -Text ([string]$text) -Width $inner)) {
            Write-Host ("  " + $v + " ") -ForegroundColor $Color -NoNewline
            Write-Host ($part.PadRight($inner)) -ForegroundColor $col -NoNewline
            Write-Host (" " + $v) -ForegroundColor $Color
        }
    }
    Write-Host ("  " + $bl + ($h * ($B - 2)) + $br) -ForegroundColor $Color
}

# Big double-line title box used at the top of screens.
function Write-Banner {
    param([string]$Title, [string]$Right = '', [string]$Color = 'Cyan', [string]$Subtitle = '')
    $W = Get-UiWidth; $B = $W - 2; $inner = $B - 4
    $r = Get-FitText -Text $Right -Width ([Math]::Max(5, [int]($inner / 2)))
    $l = Get-FitText -Text $Title -Width ($inner - $r.Length - 1)
    $gap = $inner - $l.Length - $r.Length
    Write-Host ""
    Write-Host ("  " + $script:G.DTL + ($script:G.DH * ($B - 2)) + $script:G.DTR) -ForegroundColor $Color
    Write-Host ("  " + $script:G.DV + " ") -ForegroundColor $Color -NoNewline
    Write-Host $l -ForegroundColor White -NoNewline
    Write-Host ((' ' * $gap) + $r) -ForegroundColor DarkGray -NoNewline
    Write-Host (" " + $script:G.DV) -ForegroundColor $Color
    if ($Subtitle) {
        foreach ($part in (Split-TextToWidth -Text $Subtitle -Width $inner)) {
            Write-Host ("  " + $script:G.DV + " ") -ForegroundColor $Color -NoNewline
            Write-Host ($part.PadRight($inner)) -ForegroundColor Gray -NoNewline
            Write-Host (" " + $script:G.DV) -ForegroundColor $Color
        }
    }
    Write-Host ("  " + $script:G.DBL + ($script:G.DH * ($B - 2)) + $script:G.DBR) -ForegroundColor $Color
}

# Status messages with ASCII markers that render in every console font. Long messages
# wrap with a hanging indent instead of running past the window edge.
function Write-Marked {
    param([string]$Prefix, [string]$Message, [string]$Color)
    $avail = (Get-UiWidth) - $Prefix.Length
    $parts = Split-TextToWidth -Text $Message -Width $avail
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $lead = if ($i -eq 0) { $Prefix } else { ' ' * $Prefix.Length }
        Write-Host ($lead + $parts[$i]) -ForegroundColor $Color
    }
}
function Write-Ok      { param([string]$Message) Write-Marked -Prefix '  [OK] ' -Message $Message -Color Green }
function Write-Bad     { param([string]$Message) Write-Marked -Prefix '  [X]  ' -Message $Message -Color Red }
function Write-Caution { param([string]$Message) Write-Marked -Prefix '  [!]  ' -Message $Message -Color Yellow }
function Write-Info    { param([string]$Message) Write-Marked -Prefix '  [i]  ' -Message $Message -Color Cyan }
function Write-Hint    { param([string]$Message) Write-Marked -Prefix '       ' -Message $Message -Color DarkGray }
function Write-Step    { param([string]$Step, [string]$Message, [string]$Color = 'Magenta') Write-Marked -Prefix ("  [{0}] " -f $Step) -Message $Message -Color $Color }

# A text progress bar such as  [#########.........]  for summaries.
function Get-BarText {
    param([double]$Fraction, [int]$Width = 20)
    if ($Fraction -lt 0) { $Fraction = 0 }
    if ($Fraction -gt 1) { $Fraction = 1 }
    $full = [int][Math]::Round($Fraction * $Width)
    return ($script:G.Block * $full) + ($script:G.Shade * ($Width - $full))
}

# --------------------------------------------------------------------------------------
# Number / time formatting
# --------------------------------------------------------------------------------------
function Format-FileSize {
    param([long]$bytes)
    if ($bytes -ge 1TB) { return ("{0:N2} TB" -f ($bytes / 1TB)) }
    if ($bytes -ge 1GB) { return ("{0:N2} GB" -f ($bytes / 1GB)) }
    if ($bytes -ge 1MB) { return ("{0:N2} MB" -f ($bytes / 1MB)) }
    if ($bytes -ge 1KB) { return ("{0:N1} KB" -f ($bytes / 1KB)) }
    return "$bytes bytes"
}

function Format-Count { param([long]$Number) return ("{0:N0}" -f $Number) }

# "45s", "5m 12s", "2h 15m" - truncates (never rounds up into "60m").
function Format-Duration {
    param([double]$Seconds)
    if ([double]::IsNaN($Seconds) -or [double]::IsInfinity($Seconds) -or $Seconds -lt 0) { return '' }
    $s = [long][Math]::Floor($Seconds)
    if ($s -lt 1)    { return '<1s' }
    if ($s -lt 60)   { return "${s}s" }
    if ($s -lt 3600) { return ("{0}m {1}s" -f [Math]::Floor($s / 60), ($s % 60)) }
    if ($s -lt 86400) { return ("{0}h {1}m" -f [Math]::Floor($s / 3600), [Math]::Floor(($s % 3600) / 60)) }
    return ("{0}d {1}h" -f [Math]::Floor($s / 86400), [Math]::Floor(($s % 86400) / 3600))
}

function Format-Elapsed { param([TimeSpan]$Span) return (Format-Duration $Span.TotalSeconds) }

function Get-ETA {
    param([double]$Current, [double]$Total, [TimeSpan]$Elapsed)
    if ($Current -le 0 -or $Total -le 0 -or $Current -ge $Total) { return '' }
    if ($Elapsed.TotalSeconds -lt 2) { return '' }
    $rate = $Current / $Elapsed.TotalSeconds
    if ($rate -le 0) { return '' }
    return (Format-Duration (($Total - $Current) / $rate))
}

function Format-AvgSpeed {
    param([long]$Bytes, [TimeSpan]$Elapsed)
    if ($Bytes -le 0 -or $Elapsed.TotalSeconds -lt 1) { return '' }
    return "$(Format-FileSize ([long]($Bytes / $Elapsed.TotalSeconds)))/s"
}

function Format-FilesPerMin {
    param([long]$Count, [TimeSpan]$Elapsed)
    if ($Count -le 0 -or $Elapsed.TotalSeconds -lt 1) { return '' }
    $rate = [long][Math]::Floor($Count / $Elapsed.TotalMinutes)
    if ($rate -le 0) { return '' }
    return ("{0:N0} files/min" -f $rate)
}

# --------------------------------------------------------------------------------------
# Window title (shows progress on the taskbar) and progress bars
# --------------------------------------------------------------------------------------
function Set-AppTitle {
    param([string]$Status = '')
    try {
        if ($null -eq $script:OriginalTitle) { $script:OriginalTitle = $Host.UI.RawUI.WindowTitle }
        $t = "$($script:AppName) $($script:AppVersion)"
        if ($Status) { $t = "$Status - $t" }
        $Host.UI.RawUI.WindowTitle = $t
    } catch { }
}

function Restore-AppTitle {
    try { if ($null -ne $script:OriginalTitle) { $Host.UI.RawUI.WindowTitle = $script:OriginalTitle } } catch { }
}

# Thin wrapper around Write-Progress. Id 1 = whole job / batch, Id 2 = current item,
# Id 3 = current large file. Percent is clamped so Write-Progress never throws.
function Update-Progress {
    param(
        [int]$Id = 1, [string]$Activity, [string]$Status = ' ', [double]$Percent = -1,
        [string]$Detail = '', [switch]$Completed
    )
    try {
        if ($Completed) { Write-Progress -Id $Id -Activity ($(if ($Activity) { $Activity } else { 'Working' })) -Completed; return }
        $p = @{ Id = $Id; Activity = $Activity; Status = $(if ($Status) { $Status } else { ' ' }) }
        if ($Id -gt 1) { $p.ParentId = $Id - 1 }
        if ($Percent -ge 0) { $p.PercentComplete = [int][Math]::Max(0, [Math]::Min(100, $Percent)) }
        if ($Detail) { $p.CurrentOperation = $Detail }
        Write-Progress @p
        if ($Id -eq 1) {
            $pctText = if ($Percent -ge 0) { "[{0}%] " -f [int][Math]::Min(100, $Percent) } else { '' }
            Set-AppTitle -Status ($pctText + $Activity)
        }
    } catch { }
}

function Complete-AllProgress {
    foreach ($i in 3, 2, 1) { try { Write-Progress -Id $i -Activity 'Done' -Completed } catch { } }
    Set-AppTitle
}

# ======================================================================================
# SECTION 3 - LOGGING
# ======================================================================================

function Add-LogText {
    param([string]$Text)
    $enc = New-Object System.Text.UTF8Encoding($false)
    for ($try = 0; $try -lt 3; $try++) {
        # Rotation is attempted separately so a failed rename never loses the log line.
        try {
            if ([System.IO.File]::Exists($LogFile)) {
                $fi = New-Object System.IO.FileInfo($LogFile)
                if ($fi.Length -gt ($MaxLogSizeMB * 1MB)) {
                    $rot = Join-Path $LogFileDirectory ("HybridJobLog_{0}_{1}.log" -f $ActiveUser, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                    [System.IO.File]::Move($LogFile, $rot)
                }
            }
        } catch { }
        try {
            [System.IO.File]::AppendAllText($LogFile, $Text + "`r`n", $enc)
            return
        } catch { Start-Sleep -Milliseconds 100 }
    }
}

# Write-Log keeps the same switches as earlier versions:
#   -ConsoleOnly : screen only      -FileOnly : log file only
#   -Detail      : per-file chatter, hidden on screen in Compact mode (still logged if not ConsoleOnly)
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO', [switch]$ConsoleOnly, [switch]$FileOnly, [switch]$Detail)
    $now = Get-Date
    $show = -not $FileOnly
    if ($show -and $Detail -and $script:Settings -and $script:Settings.ConsoleDetail -eq 'Compact' -and $Level -ne 'ERROR') { $show = $false }
    if ($show) {
        switch ($Level) {
            'ERROR'   { $tag = '[X] '; $col = 'Red' }
            'WARN'    { $tag = '[!] '; $col = 'Yellow' }
            'SUCCESS' { $tag = '[OK]'; $col = 'Green' }
            default   { $tag = '[i] '; $col = 'Cyan' }
        }
        Write-Marked -Prefix ("  {0} {1} " -f $now.ToString('HH:mm:ss'), $tag) -Message $Message -Color $col
    }
    if ($ConsoleOnly) { return }
    Add-LogText ("[{0}] [{1}] {2}" -f $now.ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message)
}

# Write many lines to the log in one disk write.
function Write-LogBlock {
    param([string[]]$Lines, [string]$Level = 'INFO')
    if (-not $Lines -or $Lines.Count -eq 0) { return }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        [void]$sb.Append("[$ts] [$Level] $($Lines[$i])")
        if ($i -lt $Lines.Count - 1) { [void]$sb.Append("`r`n") }
    }
    Add-LogText $sb.ToString()
}

# Result panel (double border) that is also written to the log. Same parameters as in
# earlier versions so the call sites read the same.
function Write-SummaryBox {
    param([string]$Title, [object[]]$Lines, [string]$BoxColor = 'Cyan')
    $rows = @($Lines | Where-Object { $null -ne $_ })
    Write-Panel -Title $Title -Lines $rows -Color $BoxColor -Double
    $logLines = New-Object System.Collections.Generic.List[string]
    $logLines.Add("==== $Title ====")
    foreach ($r in $rows) {
        if ($r -is [hashtable] -and $r.ContainsKey('Label')) { $logLines.Add(("  {0}: {1}" -f $r.Label, $r.Value)) }
        elseif ($r -is [hashtable]) { $logLines.Add("  $($r.Text)") }
        elseif ($r -ne '-') { $logLines.Add("  $r") }
    }
    Write-LogBlock -Lines $logLines.ToArray()
}

# Failed items: the first few on screen, every one of them in the log.
function Write-FailedReport {
    param([object]$FailedList, [int]$ShowMax = 12)
    $list = @($FailedList)
    if ($list.Count -eq 0) { return }
    $rows = New-Object System.Collections.Generic.List[object]
    $shown = 0
    foreach ($f in $list) {
        if ($shown -ge $ShowMax) { break }
        $name = if ($f.File) { [string]$f.File } else { 'Unknown item' }
        $why  = if ($f.Reason) { ([string]$f.Reason -replace '\s+', ' ') } else { 'Unknown error' }
        $rows.Add(@{ Text = "> $name"; Color = 'White' })
        $rows.Add(@{ Text = "    $why"; Color = 'DarkGray' })
        $shown++
    }
    if ($list.Count -gt $ShowMax) {
        $rows.Add('-')
        $rows.Add(@{ Text = "...and $(Format-Count ($list.Count - $ShowMax)) more. The full list is in the log file."; Color = 'Yellow' })
    }
    Write-Panel -Title ("FAILED ITEMS ({0})" -f (Format-Count $list.Count)) -Lines $rows.ToArray() -Color Red
    $logLines = New-Object System.Collections.Generic.List[string]
    $logLines.Add("==== FAILED ITEMS ($($list.Count)) ====")
    foreach ($f in $list) { $logLines.Add(("  FILE: {0} | PATH: {1} | REASON: {2}" -f $f.File, $f.Path, $f.Reason)) }
    Write-LogBlock -Lines $logLines.ToArray() -Level 'ERROR'
}

# Groups identical error messages so the real cause stands out.
function Get-DistinctErrors {
    param([object]$FailedList, [int]$Top = 0)
    $counter = @{}
    foreach ($e in @($FailedList)) {
        $reason = if ($e.Reason) { ([string]$e.Reason -replace '\s+', ' ').Trim() } else { '(no reason recorded)' }
        if ($reason.Length -gt 140) { $reason = $reason.Substring(0, 137) + '...' }
        if ($counter.ContainsKey($reason)) { $counter[$reason]++ } else { $counter[$reason] = 1 }
    }
    $sorted = @($counter.GetEnumerator() | Sort-Object Value -Descending)
    if ($Top -gt 0) { $sorted = @($sorted | Select-Object -First $Top) }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($s in $sorted) { $lines.Add(("x{0,-5} {1}" -f $s.Value, $s.Key)) }
    return ,$lines.ToArray()
}

function Get-TopExtensions {
    param([object]$Names, [int]$Top = 6)
    $counter = @{}; $total = 0
    foreach ($n in @($Names)) {
        if (-not $n) { continue }
        $ext = [System.IO.Path]::GetExtension([string]$n)
        if ([string]::IsNullOrWhiteSpace($ext)) { $ext = '(none)' } else { $ext = $ext.ToLowerInvariant() }
        if ($counter.ContainsKey($ext)) { $counter[$ext]++ } else { $counter[$ext] = 1 }
        $total++
    }
    if ($total -eq 0) { return '' }
    $parts = @($counter.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First $Top | ForEach-Object { "{0:N0} {1}" -f $_.Value, $_.Key })
    if ($counter.Count -gt $Top) { $parts += "+$($counter.Count - $Top) more" }
    return ($parts -join ', ')
}

# ======================================================================================
# SECTION 4 - SETTINGS, FAVOURITES, RECENT PATHS, JOB HISTORY
# ======================================================================================

$script:DefaultSettings = [ordered]@{
    Theme         = 'Unicode'    # Unicode | Ascii
    ShowGuides    = $true        # show the short explanation panel when an option opens
    ConsoleDetail = 'Normal'     # Normal (one line per file) | Compact (errors + summaries only)
    Sound         = $true        # chime when a job that ran for 30s+ finishes
    DefaultBuffer = ''           # temp download buffer ('' = the Windows temp folder)
    LargeFileMB   = 1024         # local copies at or above this size show their own progress bar
}

function Import-AppSettings {
    $s = [ordered]@{}
    foreach ($k in $script:DefaultSettings.Keys) { $s[$k] = $script:DefaultSettings[$k] }
    if (Test-Path -LiteralPath $script:SettingsFile) {
        try {
            $obj = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($script:SettingsFile))
            foreach ($p in $obj.PSObject.Properties) { if ($s.Contains($p.Name)) { $s[$p.Name] = $p.Value } }
        } catch { Write-Host "  Note: settings file could not be read - using defaults." -ForegroundColor Yellow }
    }
    if ($s.Theme -notin @('Unicode', 'Ascii')) { $s.Theme = 'Unicode' }
    if ($s.ConsoleDetail -notin @('Normal', 'Compact')) { $s.ConsoleDetail = 'Normal' }
    try { $s.LargeFileMB = [int]$s.LargeFileMB } catch { $s.LargeFileMB = 1024 }
    if ($s.LargeFileMB -lt 50) { $s.LargeFileMB = 50 }
    $script:Settings = $s
    $style = if ($Ascii) { 'Ascii' } else { $s.Theme }
    Initialize-Theme -Style $style
}

function Save-AppSettings {
    try {
        $json = ConvertTo-Json -InputObject ([pscustomobject]$script:Settings) -Depth 3
        [System.IO.File]::WriteAllText($script:SettingsFile, $json)
    } catch { Write-Caution "Could not save settings: $($_.Exception.Message)" }
}

function Get-LargeFileBytes { return ([long]$script:Settings.LargeFileMB) * 1MB }

# --------------------------------------------------------------------------------------
# Favourites  (file format unchanged: JSON array of {Name, Type, Value, Added})
# --------------------------------------------------------------------------------------
# PowerShell 5.1 quirk: ConvertFrom-Json emits a JSON array as ONE object. Wrapping the
# call itself in @( ) therefore gives a one-element array that contains the real array,
# which is what made favourites "merge" into one entry in older versions. Assigning
# first and then wrapping enumerates correctly in every PowerShell version.
function Get-Favorites {
    if (-not (Test-Path -LiteralPath $FavoritesFile)) { return ,@() }
    try {
        $raw = [System.IO.File]::ReadAllText($FavoritesFile)
        if ([string]::IsNullOrWhiteSpace($raw)) { return ,@() }
        $parsed = ConvertFrom-Json -InputObject $raw
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($item in @($parsed)) {
            if ($null -eq $item) { continue }
            # Repair files that older versions saved with merged (multi-value) entries.
            $names  = @($item.Name);  $types = @($item.Type)
            $values = @($item.Value); $added = @($item.Added)
            for ($i = 0; $i -lt $values.Count; $i++) {
                $n = if ($i -lt $names.Count) { "$($names[$i])" } else { "Favourite $($i + 1)" }
                $t = if ($i -lt $types.Count) { "$($types[$i])" } else { 'Any' }
                $a = if ($i -lt $added.Count) { "$($added[$i])" } else { '' }
                if ($t -notin @('Cloud', 'Local', 'Any')) { $t = 'Any' }
                if (-not [string]::IsNullOrWhiteSpace("$($values[$i])")) {
                    $list.Add([pscustomobject]@{ Name = $n; Type = $t; Value = "$($values[$i])"; Added = $a })
                }
            }
        }
        return ,$list.ToArray()
    } catch {
        Write-Caution "Favourites file could not be read ($($_.Exception.Message))."
        return ,@()
    }
}

# Each entry is serialised on its own and joined by hand: this sidesteps every PS 5.1
# ConvertTo-Json array quirk and always produces a valid JSON array.
function Save-Favorites {
    param([object[]]$Favorites)
    try {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($f in @($Favorites)) {
            if ($null -eq $f) { continue }
            $clean = [pscustomobject]@{ Name = "$($f.Name)"; Type = "$($f.Type)"; Value = "$($f.Value)"; Added = "$($f.Added)" }
            $parts.Add((ConvertTo-Json -InputObject $clean -Compress -Depth 2))
        }
        $json = if ($parts.Count -eq 0) { '[]' } else { "[`r`n  " + ($parts -join ",`r`n  ") + "`r`n]" }
        [System.IO.File]::WriteAllText($FavoritesFile, $json)
        return $true
    } catch {
        Write-Bad "Could not save favourites: $($_.Exception.Message)"
        return $false
    }
}

function Add-Favorite {
    param([string]$Name, [string]$Type, [string]$Value)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($f in (Get-Favorites)) { $list.Add($f) }
    $list.Add([pscustomobject]@{ Name = $Name.Trim(); Type = $Type; Value = $Value.Trim(); Added = (Get-Date -Format 'yyyy-MM-dd') })
    if (Save-Favorites -Favorites $list.ToArray()) {
        Write-Log "Favourites: added '$($Name.Trim())' [$Type] => '$($Value.Trim())'" 'SUCCESS' -FileOnly
        return $true
    }
    return $false
}

function Get-FavoriteTypeTag {
    param([string]$Type)
    switch ($Type) { 'Cloud' { return 'CLOUD' } 'Local' { return 'LOCAL' } default { return 'ANY  ' } }
}

function Show-FavoriteList {
    param([object[]]$Items, [switch]$WithDates)
    $W = Get-UiWidth
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $f = $Items[$i]
        $col = switch ($f.Type) { 'Cloud' { 'Yellow' } 'Local' { 'Green' } default { 'White' } }
        Write-Host ("  {0,3}  " -f ($i + 1)) -ForegroundColor Cyan -NoNewline
        Write-Host ("[{0}] " -f (Get-FavoriteTypeTag $f.Type)) -ForegroundColor $col -NoNewline
        Write-Host (Get-FitText -Text $f.Name -Width ($W - 18)) -ForegroundColor White
        $extra = if ($WithDates -and $f.Added) { "   (added $($f.Added))" } else { '' }
        $val = Get-FitText -Text $f.Value -Width ($W - 12 - $extra.Length) -Middle
        Write-Host ("            {0}{1}" -f $val, $extra) -ForegroundColor DarkGray
    }
}

# Returns the chosen favourite's Value, or $null.
function Show-FavoritePicker {
    param([string]$Type = 'Any')
    $all = (Get-Favorites)
    $filtered = @($all | Where-Object { $Type -eq 'Any' -or $_.Type -eq $Type -or $_.Type -eq 'Any' })
    if ($filtered.Count -eq 0) { $filtered = $all }
    if ($filtered.Count -eq 0) {
        Write-Caution "No favourites saved yet. Add some from the main menu with [F]."
        return $null
    }
    Write-Host ""
    Write-Rule -Title ("FAVOURITES" + $(if ($Type -ne 'Any') { " ($Type)" } else { '' })) -Color Cyan
    Show-FavoriteList -Items $filtered
    Write-Host ""
    $pick = (Read-Line -Prompt "  Number (ENTER = back) > " -LogLabel 'Favourite pick').Trim()
    if (-not $pick) { return $null }
    if ($pick -match '^\d{1,6}$' -and [int]$pick -ge 1 -and [int]$pick -le $filtered.Count) {
        $chosen = $filtered[[int]$pick - 1]
        Write-Ok "Using favourite: $($chosen.Name)"
        return $chosen.Value
    }
    Write-Caution "That number is not in the list."
    return $null
}

function Manage-Favorites {
    while ($true) {
        $favs = (Get-Favorites)
        Write-Banner -Title 'FAVOURITES' -Right "$($favs.Count) saved" -Color Cyan -Subtitle 'Saved URLs and paths appear at every path prompt when you type F.'
        if ($favs.Count -eq 0) {
            Write-Hint "No favourites yet - press A to add your first one."
        } else {
            Show-FavoriteList -Items $favs -WithDates
        }
        Write-Host ""
        Write-Marked -Prefix '  ' -Message 'A Add   R Rename   E Edit value   D Delete   U/N Move up/down   T Test paths   Q Back' -Color White
        $action = (Read-Line -Prompt "  Choice > " -LogLabel 'Favourites action').Trim().ToUpperInvariant()

        switch -Regex ($action) {
            '^A$' {
                Write-Host ""
                Write-Host "  1 Cloud URL (https://...)   2 Local / network path   3 Either" -ForegroundColor Gray
                $tc = (Read-Line -Prompt "  Type [1/2/3] > " -LogLabel 'Favourite type').Trim()
                $type = switch ($tc) { '1' { 'Cloud' } '2' { 'Local' } default { 'Any' } }
                $value = ConvertTo-CleanPath (Read-Line -Prompt "  URL or path > " -LogLabel 'Favourite value')
                if (-not $value) { Write-Caution "Nothing entered - not saved."; break }
                if ($type -eq 'Any') { $type = if ($value -match '^(?i)https?://') { 'Cloud' } else { 'Local' } }
                $suggest = 'Favourite'
                try {
                    if ($value -match '^(?i)https?://') {
                        $pc = Parse-CloudUrl -Url $value -Quiet
                        if ($pc) { $suggest = ($pc.FolderPath.TrimEnd('/') -split '/')[-1] }
                    } else { $suggest = Split-Path -Path $value.TrimEnd('\', '/') -Leaf }
                } catch { }
                if (-not $suggest) { $suggest = 'Favourite' }
                $name = (Read-Line -Prompt "  Friendly name (ENTER = '$suggest') > " -LogLabel 'Favourite name').Trim()
                if (-not $name) { $name = $suggest }
                if (Add-Favorite -Name $name -Type $type -Value $value) { Write-Ok "Saved '$name'." }
            }
            '^(R|E|D|U|N)$' {
                if ($favs.Count -eq 0) { Write-Caution "There are no favourites yet."; break }
                $n = (Read-Line -Prompt "  Which number (1-$($favs.Count))? > " -LogLabel 'Favourite number').Trim()
                if ($n -notmatch '^\d{1,6}$' -or [int]$n -lt 1 -or [int]$n -gt $favs.Count) { Write-Caution "Not a valid number."; break }
                $idx = [int]$n - 1
                $list = New-Object System.Collections.Generic.List[object]
                foreach ($f in $favs) { $list.Add([pscustomobject]@{ Name = $f.Name; Type = $f.Type; Value = $f.Value; Added = $f.Added }) }
                $target = $list[$idx]
                if ($action -eq 'R') {
                    $new = (Read-Line -Prompt "  New name for '$($target.Name)' > " -LogLabel 'Favourite rename').Trim()
                    if (-not $new) { Write-Caution "No change."; break }
                    $target.Name = $new
                } elseif ($action -eq 'E') {
                    Write-Hint "Current: $($target.Value)"
                    $new = ConvertTo-CleanPath (Read-Line -Prompt "  New URL or path (ENTER = keep) > " -LogLabel 'Favourite edit')
                    if (-not $new) { Write-Caution "No change."; break }
                    $target.Value = $new
                } elseif ($action -eq 'D') {
                    if (-not (Read-Confirmation -Message "Delete favourite '$($target.Name)'?" -Word 'YES')) { Write-Hint "Not deleted."; break }
                    $list.RemoveAt($idx)
                } elseif ($action -eq 'U') {
                    if ($idx -eq 0) { break }
                    $list.RemoveAt($idx); $list.Insert($idx - 1, $target)
                } elseif ($action -eq 'N') {
                    if ($idx -ge $list.Count - 1) { break }
                    $list.RemoveAt($idx); $list.Insert($idx + 1, $target)
                }
                if (Save-Favorites -Favorites $list.ToArray()) { Write-Ok "Favourites updated." }
            }
            '^T$' {
                Write-Host ""
                foreach ($f in $favs) {
                    if ($f.Value -match '^(?i)https?://') {
                        $ok = [bool](Parse-CloudUrl -Url $f.Value -Quiet)
                        if ($ok) { Write-Ok "$($f.Name): URL format looks valid" } else { Write-Bad "$($f.Name): URL cannot be understood" }
                    } elseif ($f.Value -match '^\\\\') {
                        Write-Hint "$($f.Name): network path - not tested (could be slow if the server is offline)"
                    } elseif (Test-Path -LiteralPath $f.Value) { Write-Ok "$($f.Name): folder exists" }
                    else { Write-Bad "$($f.Name): not found on this PC" }
                }
                Show-Pause -Message 'Press any key to continue...'
            }
            '^(Q|)$' { return }
            default { Write-Caution "Please choose A, R, E, D, U, N, T or Q." }
        }
    }
}

# --------------------------------------------------------------------------------------
# Recently used paths (this session only, never written to disk)
# --------------------------------------------------------------------------------------
function Add-RecentValue {
    param([string]$Type, [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    if (-not $script:Recent.ContainsKey($Type)) { $script:Recent[$Type] = New-Object System.Collections.Generic.List[string] }
    $l = $script:Recent[$Type]
    for ($i = $l.Count - 1; $i -ge 0; $i--) { if ($l[$i] -eq $Value) { $l.RemoveAt($i) } }
    $l.Insert(0, $Value)
    while ($l.Count -gt 8) { $l.RemoveAt($l.Count - 1) }
}

function Get-RecentValues {
    param([string]$Type)
    $keys = if ($Type -eq 'Any') { @('Cloud', 'Local') } else { @($Type) }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($k in $keys) { if ($script:Recent.ContainsKey($k)) { foreach ($v in $script:Recent[$k]) { if (-not $out.Contains($v)) { $out.Add($v) } } } }
    return ,$out.ToArray()
}

# --------------------------------------------------------------------------------------
# Job history and session scoreboard
# --------------------------------------------------------------------------------------
function Get-JobHistory {
    if (-not (Test-Path -LiteralPath $script:HistoryFile)) { return ,@() }
    try {
        $parsed = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($script:HistoryFile))
        return ,@(@($parsed) | Where-Object { $null -ne $_ })
    } catch { return ,@() }
}

function Save-JobHistory {
    param([object[]]$Entries)
    try {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($e in @($Entries)) { if ($null -ne $e) { $parts.Add((ConvertTo-Json -InputObject $e -Compress -Depth 2)) } }
        $json = if ($parts.Count -eq 0) { '[]' } else { "[`r`n" + ($parts -join ",`r`n") + "`r`n]" }
        [System.IO.File]::WriteAllText($script:HistoryFile, $json)
    } catch { }
}

function New-JobStats {
    return @{
        Success = 0; Skipped = 0; Failed = 0; TooLong = 0; Unreadable = 0; Accepted = 0
        SuccessBytes = [long]0; SkippedBytes = [long]0; TotalBytes = [long]0
        FailedItems = (New-Object System.Collections.Generic.List[hashtable])
        Manifest    = (New-Object System.Collections.Generic.List[hashtable])
        Names       = (New-Object System.Collections.Generic.List[string])
        ListingErrors = 0; Cancelled = $false
    }
}

function Add-JobFailure {
    param([hashtable]$Stats, [string]$File, [string]$Path, [string]$Reason)
    $Stats.Failed++
    $Stats.FailedItems.Add(@{ File = $File; Path = $Path; Reason = $Reason })
}

function Merge-JobStats {
    param([hashtable]$Into, [hashtable]$From)
    if ($null -eq $From) { return }
    foreach ($k in 'Success', 'Skipped', 'Failed', 'TooLong', 'Unreadable', 'Accepted', 'ListingErrors') { if ($From.ContainsKey($k)) { $Into[$k] += [int]$From[$k] } }
    foreach ($k in 'SuccessBytes', 'SkippedBytes', 'TotalBytes') { if ($From.ContainsKey($k)) { $Into[$k] += [long]$From[$k] } }
    if ($From.FailedItems) { foreach ($x in $From.FailedItems) { $Into.FailedItems.Add($x) } }
    if ($From.Manifest)    { foreach ($x in $From.Manifest)    { $Into.Manifest.Add($x) } }
    if ($From.Names -and $Into.Names.Count -lt 200000) { foreach ($x in $From.Names) { $Into.Names.Add($x) } }
    if ($From.Cancelled) { $Into.Cancelled = $true }
}

# Records a finished job in the session scoreboard and the persistent history, and plays
# the completion chime for long jobs.
function Register-JobResult {
    param(
        [string]$Operation, [string]$Source = '', [string]$Destination = '',
        [hashtable]$Stats, [TimeSpan]$Elapsed, [string]$Outcome = ''
    )
    if (-not $Outcome) {
        if ($Stats -and $Stats.Cancelled) { $Outcome = 'Cancelled' }
        elseif ($Stats -and ($Stats.Failed -gt 0 -or $Stats.ListingErrors -gt 0 -or $Stats.Unreadable -gt 0)) { $Outcome = 'Completed with errors' }
        else { $Outcome = 'Success' }
    }
    $ok = 0; $fail = 0; $skip = 0; $bytes = [long]0
    if ($Stats) { $ok = [int]$Stats.Success; $fail = [int]$Stats.Failed; $skip = [int]$Stats.Skipped; $bytes = [long]$Stats.SuccessBytes }
    $script:Session.Jobs++
    $script:Session.Files  += $ok
    $script:Session.Failed += $fail
    $script:Session.Bytes  += $bytes
    $entry = [pscustomobject]@{
        When = (Get-Date -Format 'yyyy-MM-dd HH:mm'); Operation = $Operation; Source = $Source; Destination = $Destination
        Ok = $ok; Skipped = $skip; Failed = $fail; Bytes = $bytes; Seconds = [int]$Elapsed.TotalSeconds; Outcome = $Outcome
    }
    $hist = New-Object System.Collections.Generic.List[object]
    foreach ($h in (Get-JobHistory)) { $hist.Add($h) }
    $hist.Add($entry)
    $total = $hist.Count
    while ($hist.Count -gt 500) { $hist.RemoveAt(0) }
    Save-JobHistory -Entries $hist.ToArray()
    Write-Log ("JOB DONE: {0} | {1} | ok={2} skipped={3} failed={4} | {5} | {6}" -f $Operation, $Outcome, $ok, $skip, $fail, (Format-FileSize $bytes), (Format-Elapsed $Elapsed)) 'INFO' -FileOnly

    if ($script:Settings.Sound -and $Elapsed.TotalSeconds -ge 30) {
        try { if ($fail -gt 0) { [Console]::Beep(440, 300) } else { [Console]::Beep(784, 120); [Console]::Beep(1046, 180) } } catch { }
    }
    $milestones = @(10, 25, 50, 100, 250, 500)
    if ($milestones -contains $total) { Write-Host ("  *** Milestone: that was job number {0} with this tool. Nice work! ***" -f $total) -ForegroundColor Magenta }
    Set-AppTitle
}

# ======================================================================================
# SECTION 5 - KEYBOARD, PROMPTS AND PATH INPUT
# ======================================================================================

$script:IsISE = ($Host.Name -match 'ISE')

function Clear-KeyBuffer {
    try { while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) } } catch { }
}

# While a job runs, Ctrl+C is read as a normal key so it can cancel the job cleanly
# (finishing the current file and printing the summary) instead of killing the script.
function Set-CtrlCAsInput {
    param([bool]$Enabled)
    try { [Console]::TreatControlCAsInput = $Enabled } catch { }
}

# Every engine calls Start-JobUi / Stop-JobUi. Only the outermost call resets the cancel
# flag, so pressing C inside a nested item really stops the whole batch.
function Start-JobUi {
    if ($script:JobDepth -le 0) {
        $script:JobDepth = 0
        $script:JobCancelled = $false
        Clear-KeyBuffer
        Set-CtrlCAsInput $true
    }
    $script:JobDepth++
}

function Stop-JobUi {
    if ($script:JobDepth -gt 0) { $script:JobDepth-- }
    if ($script:JobDepth -le 0) {
        $script:JobDepth = 0
        Set-CtrlCAsInput $false
        Clear-KeyBuffer
        Complete-AllProgress
    }
}

# Reads a line without the ": " that Read-Host appends. Works in the console and the ISE.
function Read-Line {
    param([string]$Prompt = '  > ', [string]$LogLabel = '', [switch]$NoLog, [string]$Color = 'Cyan')
    # The key buffer is NOT cleared here: a multi-line paste arrives as queued keys and
    # every pasted line must reach the following prompts. Stray keys pressed during a job
    # are discarded by Stop-JobUi instead.
    Set-CtrlCAsInput $false
    Write-Host $Prompt -ForegroundColor $Color -NoNewline
    $v = $null
    try { $v = $Host.UI.ReadLine() } catch { try { $v = Read-Host } catch { $v = $null } }
    if ($script:JobDepth -gt 0) { Set-CtrlCAsInput $true }
    if ($null -eq $v) { $v = 'Q' }   # end of input (non-interactive host): behave like Q
    if (-not $NoLog) {
        $lbl = if ($LogLabel) { $LogLabel } else { $Prompt.Trim() }
        Write-Log "[USER INPUT] ${lbl}: '$v'" 'INFO' -FileOnly
    }
    return $v
}

function Show-Pause {
    param([string]$Message = 'Press any key to return to the main menu...')
    Write-Host ""
    Write-Host "  $Message" -ForegroundColor DarkGray
    Clear-KeyBuffer
    try { [void][Console]::ReadKey($true) } catch { try { $null = Read-Host } catch { } }
}

# Returns $true only when the exact word is typed (case-sensitive, like earlier versions).
function Read-Confirmation {
    param([string]$Message, [string]$Word = 'YES', [string]$Color = 'Yellow')
    Write-Host ""
    foreach ($line in (Split-TextToWidth -Text $Message -Width ((Get-UiWidth) - 4))) { Write-Host "  $line" -ForegroundColor $Color }
    $ans = Read-Line -Prompt ("  Type {0} to continue (anything else cancels) > " -f $Word) -LogLabel 'Confirmation'
    return ($ans.Trim() -ceq $Word)
}

function Read-YesNo {
    param([string]$Question, [bool]$Default = $false)
    $hint = if ($Default) { '[Y/n]' } else { '[y/N]' }
    $ans = (Read-Line -Prompt ("  {0} {1} > " -f $Question, $hint) -LogLabel $Question).Trim()
    if (-not $ans) { return $Default }
    return ($ans -match '^(?i)y(es)?$')
}

# Returns the upper-case letter chosen from -Valid, or '' when the user just pressed ENTER/Q.
function Read-Choice {
    param([string]$Prompt, [string[]]$Valid)
    while ($true) {
        $a = (Read-Line -Prompt "  $Prompt > " -LogLabel $Prompt).Trim().ToUpperInvariant()
        if (-not $a -or $a -eq 'Q') { return '' }
        if ($Valid -contains $a) { return $a }
        Write-Caution ("Please type one of: {0} (or Q to go back)." -f ($Valid -join ', '))
    }
}

# Pause / cancel hotkeys while a job is running: P = pause, C or Ctrl+C = cancel.
# Safe to call thousands of times per second (only reads the keyboard when a key is waiting).
function Check-Interrupt {
    param([string]$CurrentFile = '', [string]$CurrentFolder = '')
    if ($script:JobCancelled) { return 'CANCEL' }
    try {
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::C) {
                $script:JobCancelled = $true
                Write-Host ""
                Write-Log "CANCEL requested by user. Last folder: [$CurrentFolder] | last file: [$CurrentFile]. Finishing safely..." 'WARN'
                return 'CANCEL'
            }
            if ($key.Key -eq [ConsoleKey]::P) {
                Write-Host ""
                $lines = @()
                if ($CurrentFolder) { $lines += @{ Label = 'Folder'; Value = $CurrentFolder } }
                if ($CurrentFile)   { $lines += @{ Label = 'File';   Value = $CurrentFile } }
                $lines += @{ Text = 'Press C to cancel the rest of the job, or any other key to resume.'; Color = 'Yellow' }
                Write-Panel -Title 'JOB PAUSED' -Lines $lines -Color Magenta
                Write-Log "JOB PAUSED. Folder: [$CurrentFolder] | file: [$CurrentFile]" 'WARN' -FileOnly
                $resume = [Console]::ReadKey($true)
                if ($resume.Key -eq [ConsoleKey]::C) {
                    $script:JobCancelled = $true
                    Write-Log "CANCEL requested while paused. Finishing safely..." 'WARN'
                    return 'CANCEL'
                }
                Write-Log "Job resumed." 'INFO'
                return 'RESUME'
            }
        }
    } catch { }
    return 'CONTINUE'
}

# One status line that is rewritten in place (falls back to normal lines in the ISE).
function Write-StatusLine {
    param([string]$Text, [string]$Color = 'DarkGray')
    $w = (Get-UiWidth) - 1
    $t = Get-FitText -Text $Text -Width $w
    if ($script:IsISE) { Write-Host $t -ForegroundColor $Color }
    else { Write-Host ("`r" + $t.PadRight($w)) -NoNewline -ForegroundColor $Color }
}

function Clear-StatusLine {
    if (-not $script:IsISE) { Write-Host ("`r" + (' ' * ((Get-UiWidth) - 1)) + "`r") -NoNewline }
}

# Sleep that keeps listening for P / C, with a visible countdown.
function Start-InterruptibleSleep {
    param([int]$Seconds, [string]$Label = 'Waiting')
    for ($s = $Seconds; $s -gt 0; $s--) {
        if ((Check-Interrupt) -eq 'CANCEL') { break }
        if (-not $script:IsISE -or ($s % 10 -eq 0) -or $s -eq $Seconds) {
            Write-StatusLine -Text ("  {0} - {1}s left   (P = pause, C = cancel)" -f $Label, $s)
        }
        Start-Sleep -Seconds 1
    }
    Clear-StatusLine
}

# --------------------------------------------------------------------------------------
# Path cleaning and prompts
# --------------------------------------------------------------------------------------

# Strips quotes added by drag-and-drop / "Copy as path", trims spaces and trailing slashes,
# but never turns a drive root like C:\ into "C:" (which means "current folder on C:").
function ConvertTo-CleanPath {
    param([string]$Path)
    if ($null -eq $Path) { return '' }
    $p = $Path.Trim()
    if ($p.StartsWith('& ')) { $p = $p.Substring(2).Trim() }   # ISE drag-and-drop adds "& "
    while ($p.Length -ge 2 -and (($p[0] -eq '"' -and $p[$p.Length - 1] -eq '"') -or ($p[0] -eq "'" -and $p[$p.Length - 1] -eq "'"))) {
        $p = $p.Substring(1, $p.Length - 2).Trim()
    }
    $p = $p.Trim('"').Trim()
    if ($p -match '^(?i)https?://') { return $p }
    if ($p -match '^[A-Za-z]:$')   { return $p + '\' }
    if ($p -match '^[A-Za-z]:\\$') { return $p }
    if ($p.Length -gt 3)           { $p = $p.TrimEnd('\', '/') }
    return $p
}

function Resolve-LocalPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    try { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) } catch { return $Path }
}

function Test-WritableFolder {
    param([string]$Path)
    try {
        $long = Get-LongPath $Path
        if (-not [System.IO.Directory]::Exists($long)) { $null = [System.IO.Directory]::CreateDirectory($long) }
        $probe = Join-Path $long ("rhfm_probe_{0}.tmp" -f [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, 'x')
        [System.IO.File]::Delete($probe)
        return $true
    } catch { return $false }
}

# Windows folder-picker dialog for [B] Browse.
function Show-FolderPicker {
    param([string]$Description = 'Choose a folder')
    if (-not $script:IsWin) { Write-Caution "The folder browser is only available on Windows."; return $null }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $Description
        $dlg.ShowNewFolderButton = $true
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        Write-Hint "A folder browser window has opened (it may be behind this window)..."
        $res = $dlg.ShowDialog($owner)
        $owner.Dispose()
        if ($res -eq [System.Windows.Forms.DialogResult]::OK -and $dlg.SelectedPath) { return $dlg.SelectedPath }
        Write-Hint "No folder chosen."
    } catch { Write-Caution "The folder browser could not open here: $($_.Exception.Message)" }
    return $null
}

# The smart prompt behind every URL / path question.
#   F = favourites   R = recently used   B = browse (local)   Q = cancel (returns 'Q')
#   ENTER returns '' when -AllowEmpty is set.
function Read-PathInput {
    param(
        [string]$Prompt,
        [ValidateSet('Cloud', 'Local', 'Any')][string]$Type = 'Any',
        [string]$LogLabel = '',
        [switch]$AllowEmpty,
        [string]$EmptyHint = 'skip',
        [switch]$FreeText
    )
    while ($true) {
        $recent = (Get-RecentValues -Type $Type)
        $hints = New-Object System.Collections.Generic.List[string]
        $hints.Add('F = favourites')
        if ($recent.Count -gt 0) { $hints.Add('R = recent') }
        if ($Type -ne 'Cloud' -and $script:IsWin) { $hints.Add('B = browse') }
        if ($AllowEmpty) { $hints.Add("ENTER = $EmptyHint") }
        $hints.Add('Q = cancel')
        Write-Host ""
        foreach ($l in (Split-TextToWidth -Text $Prompt -Width ((Get-UiWidth) - 4))) { Write-Host "  $l" -ForegroundColor White }
        Write-Marked -Prefix '  ' -Message ($hints -join '   ') -Color DarkGray
        $label = if ($LogLabel) { $LogLabel } else { $Prompt }
        $val = ConvertTo-CleanPath (Read-Line -Prompt '  > ' -LogLabel $label)

        if ($val -match '^[Qq]$') { return 'Q' }
        if ($val -match '^[Ff]$') {
            $picked = Show-FavoritePicker -Type $Type
            if ($picked) { Add-RecentValue -Type $(if ($picked -match '^(?i)https?://') { 'Cloud' } else { 'Local' }) -Value $picked; return $picked }
            continue
        }
        if ($val -match '^[Rr]$') {
            if ($recent.Count -eq 0) { Write-Caution "Nothing used yet in this session."; continue }
            if ($recent.Count -eq 1) { Write-Ok "Using: $($recent[0])"; return $recent[0] }
            Write-Host ""
            for ($i = 0; $i -lt $recent.Count; $i++) {
                Write-Host ("  {0,3}  {1}" -f ($i + 1), (Get-FitText -Text $recent[$i] -Width ((Get-UiWidth) - 8) -Middle)) -ForegroundColor Gray
            }
            $n = (Read-Line -Prompt '  Number (ENTER = back) > ' -LogLabel 'Recent pick').Trim()
            if ($n -match '^\d{1,6}$' -and [int]$n -ge 1 -and [int]$n -le $recent.Count) {
                Write-Ok "Using: $($recent[[int]$n - 1])"
                return $recent[[int]$n - 1]
            }
            continue
        }
        if ($val -match '^[Bb]$' -and $Type -ne 'Cloud') {
            $picked = Show-FolderPicker -Description $Prompt
            if ($picked) { Write-Ok "Selected: $picked"; Add-RecentValue -Type 'Local' -Value $picked; return $picked }
            continue
        }
        if (-not $val) {
            if ($AllowEmpty) { return '' }
            Write-Caution "Please enter a value, or Q to cancel."
            continue
        }
        if (-not $FreeText) {
            if ($Type -eq 'Cloud' -and $val -notmatch '^(?i)https?://') {
                Write-Caution "That does not look like a web address - SharePoint / OneDrive links start with https://"
                Write-Hint "Open the folder in your browser, copy the whole address bar and paste it here."
                continue
            }
            if ($Type -eq 'Local' -and $val -match '^(?i)https?://') {
                Write-Caution "That is a web address - this question needs a folder on your PC or network."
                Write-Hint "Examples: D:\Archive   C:\Users\$ActiveUser\Documents   \\server\share\folder"
                continue
            }
        }
        if (-not $FreeText) { Add-RecentValue -Type $(if ($val -match '^(?i)https?://') { 'Cloud' } else { 'Local' }) -Value $val }
        return $val
    }
}

# Asks for a local / UNC folder and validates it. Returns the full path, '' (when
# -AllowEmpty and ENTER was pressed) or $null when the user cancelled.
function Read-LocalFolder {
    param(
        [string]$Prompt, [string]$LogLabel = '', [switch]$MustExist, [switch]$CreateIfMissing,
        [switch]$AllowEmpty, [string]$EmptyHint = 'skip', [switch]$AllowFile
    )
    while ($true) {
        $v = Read-PathInput -Prompt $Prompt -Type Local -LogLabel $LogLabel -AllowEmpty:$AllowEmpty -EmptyHint $EmptyHint
        if ($v -eq 'Q') { return $null }
        if ($v -eq '')  { return '' }
        $full = Resolve-LocalPath $v
        $long = Get-LongPath $full
        $isDir  = [System.IO.Directory]::Exists($long)
        $isFile = [System.IO.File]::Exists($long)
        if ($isFile -and -not $AllowFile) { Write-Caution "That is a file - please choose a folder."; continue }
        if ($isDir -or $isFile) { return $full }
        if ($MustExist) {
            Write-Bad "Not found: $full"
            Write-Hint "Check it in File Explorer. Network shares must be reachable and OneDrive folders synced."
            continue
        }
        if ($CreateIfMissing) {
            try {
                $null = [System.IO.Directory]::CreateDirectory($long)
                Write-Ok "Created folder: $full"
                return $full
            } catch {
                Write-Bad "Could not create '$full': $($_.Exception.Message)"
                continue
            }
        }
        return $full
    }
}

# Asks for a SharePoint / OneDrive URL and parses it. Returns the parsed hashtable,
# '' (when -AllowEmpty and ENTER) or $null when cancelled.
function Read-CloudFolder {
    param([string]$Prompt, [string]$LogLabel = '', [switch]$AllowEmpty, [string]$EmptyHint = 'skip')
    while ($true) {
        $v = Read-PathInput -Prompt $Prompt -Type Cloud -LogLabel $LogLabel -AllowEmpty:$AllowEmpty -EmptyHint $EmptyHint
        if ($v -eq 'Q') { return $null }
        if ($v -eq '')  { return '' }
        $p = Parse-CloudUrl -Url $v
        if (-not $p) {
            Write-Hint "Open the folder in your browser, copy the whole address bar and paste it here."
            continue
        }
        Write-Hint ("Site   : {0}" -f $p.SiteUrl)
        Write-Hint ("Folder : {0}" -f $p.FolderPath)
        if ($p.Note) { Write-Caution $p.Note }
        return $p
    }
}

# The temp buffer is not asked for every time any more: the Settings value (or the
# Windows temp folder) is used, and only if it is not writable is the user asked.
function Get-BufferFolder {
    $default = $script:Settings.DefaultBuffer
    if ([string]::IsNullOrWhiteSpace($default)) { $default = [System.IO.Path]::GetTempPath().TrimEnd('\', '/') }
    if (Test-WritableFolder $default) {
        Write-Hint "Temp buffer: $default   (change it in [S] Settings)"
        return $default
    }
    Write-Caution "The temp buffer '$default' is not writable."
    while ($true) {
        $v = Read-LocalFolder -Prompt 'Choose a temp buffer folder (e.g. C:\Temp):' -LogLabel 'Buffer' -CreateIfMissing
        if ($null -eq $v -or $v -eq '') { return $null }
        if (Test-WritableFolder $v) { return $v }
        Write-Bad "Cannot write to '$v' - please choose another folder."
    }
}

# --------------------------------------------------------------------------------------
# Item selection ("3", "2-5", "1,3,7", "1 3 7", "10-3", "A")
# --------------------------------------------------------------------------------------
function ConvertFrom-SelectionText {
    param([string]$Text, [int]$Count)
    $res = @{ Ok = $false; Indices = @(); Error = '' }
    $t = ("$Text").Trim()
    if ($Count -lt 1) { $res.Error = 'There is nothing to select.'; return $res }
    if ($t -match '^(?i)(a|all|\*)$') { $res.Indices = @(0..($Count - 1)); $res.Ok = $true; return $res }
    $t = $t -replace '\s*-\s*', '-'
    $set = New-Object 'System.Collections.Generic.SortedSet[int]'
    foreach ($part in ($t -split '[,;\s]+')) {
        if (-not $part) { continue }
        if ($part -match '^(\d{1,6})-(\d{1,6})$') {
            $a = [int]$matches[1]; $b = [int]$matches[2]
            if ($a -gt $b) { $tmp = $a; $a = $b; $b = $tmp }
            if ($a -lt 1 -or $b -gt $Count) { $res.Error = "Range '$part' is outside 1-$Count."; return $res }
            for ($i = $a; $i -le $b; $i++) { [void]$set.Add($i - 1) }
        } elseif ($part -match '^\d{1,6}$') {
            $n = [int]$part
            if ($n -lt 1 -or $n -gt $Count) { $res.Error = "Item '$part' is outside 1-$Count."; return $res }
            [void]$set.Add($n - 1)
        } else {
            $res.Error = "'$part' is not a number or a range like 2-5."
            return $res
        }
    }
    if ($set.Count -eq 0) { $res.Error = 'Nothing was selected.'; return $res }
    $res.Indices = [int[]]@($set)
    $res.Ok = $true
    return $res
}

# Shared prompt under every batch index. Returns @{ Action = Select|Quit|Refresh|Search; Indices }.
function Read-IndexSelection {
    param([int]$Count, [string]$Verb = 'process', [switch]$NoSearch)
    $help = "Pick items to $Verb :  3  |  2-5  |  1,3,7  |  A = all"
    if (-not $NoSearch) { $help += "  |  S = search" }
    $help += "  |  ENTER = refresh  |  Q = back"
    Write-Host ""
    foreach ($l in (Split-TextToWidth -Text $help -Width ((Get-UiWidth) - 4))) { Write-Host "  $l" -ForegroundColor DarkGray }
    while ($true) {
        $raw = (Read-Line -Prompt '  Select > ' -LogLabel 'Batch selection').Trim()
        if (-not $raw) { return @{ Action = 'Refresh'; Indices = @() } }
        if ($raw -match '^(?i)q$') { return @{ Action = 'Quit'; Indices = @() } }
        if ($raw -match '^(?i)s$' -and -not $NoSearch) { return @{ Action = 'Search'; Indices = @() } }
        $sel = ConvertFrom-SelectionText -Text $raw -Count $Count
        if ($sel.Ok) {
            Write-Hint ("Selected {0} of {1} item(s)." -f $sel.Indices.Count, $Count)
            return @{ Action = 'Select'; Indices = $sel.Indices }
        }
        Write-Caution $sel.Error
    }
}

# ======================================================================================
# SECTION 6 - LOCAL PATHS, FOLDER WALKING AND STORAGE
# ======================================================================================

# Adds the \\?\ prefix to long paths so .NET can reach past the 260-character limit.
function Get-LongPath {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path) -or -not $script:IsWin) { return $Path }
    if ($Path.StartsWith('\\?\')) { return $Path }
    $p = $Path -replace '/', '\'
    if (-not [System.IO.Path]::IsPathRooted($p) -or $p -match '^[A-Za-z]:[^\\]' -or $p -match '^[A-Za-z]:$') { $p = Resolve-LocalPath $p }
    if ($p.Length -lt 240) { return $p }
    if ($p.StartsWith('\\')) { return '\\?\UNC\' + $p.Substring(2) }
    return '\\?\' + $p
}

# The opposite of Get-LongPath, for display.
function Remove-LongPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    if ($Path.StartsWith('\\?\UNC\')) { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))     { return $Path.Substring(4) }
    return $Path
}

function Join-LocalPath {
    param([string]$Base, [string]$Child)
    if ([string]::IsNullOrEmpty($Child)) { return $Base }
    return [System.IO.Path]::Combine($Base, $Child)
}

# Path of an item relative to a root folder (case-insensitive, long-path aware).
function Get-ChildRelativePath {
    param([string]$Root, [string]$FullName)
    $r = (Remove-LongPrefix $Root).TrimEnd('\', '/')
    $f = Remove-LongPrefix $FullName
    if ($f.Length -gt $r.Length -and $f.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $f.Substring($r.Length).TrimStart('\', '/')
    }
    return [System.IO.Path]::GetFileName($f)
}

# Normalised form for comparing two locations (resolves mapped drives to their UNC path).
function Get-ComparablePath {
    param([string]$Path)
    $p = (Remove-LongPrefix (Resolve-LocalPath $Path))
    try { $p = [System.IO.Path]::GetFullPath($p) } catch { }   # normalises "..", trailing dots, etc.
    $p = $p.TrimEnd('\', '/')
    if ($p -match '^([A-Za-z]):(.*)$') {
        $drv = $matches[1]; $tail = $matches[2]
        try {
            $d = Get-PSDrive -Name $drv -ErrorAction Stop
            if ($d.DisplayRoot -and $d.DisplayRoot.StartsWith('\\')) { $p = $d.DisplayRoot.TrimEnd('\') + $tail }
        } catch { }
    }
    # \\localhost\C$\x or \\THISPC\C$\x  ->  C:\x
    $hostNames = @('localhost', '127.0.0.1', "$env:COMPUTERNAME") | Where-Object { $_ }
    foreach ($hn in $hostNames) {
        $m = [regex]::Match($p, '^\\\\' + [regex]::Escape($hn) + '\\([A-Za-z])\$(.*)$', 'IgnoreCase')
        if ($m.Success) { $p = $m.Groups[1].Value + ':' + $m.Groups[2].Value; break }
    }
    return $p.ToLowerInvariant()
}

# True when two paths reach the same file or folder on disk (Windows only; uses the
# volume serial number and file index, so aliases, mapped drives and junctions match).
function Test-SamePhysicalItem {
    param([string]$A, [string]$B)
    if (-not ($script:IsWin -and ('RHFMNative' -as [type]))) { return $false }
    try {
        $ia = [RHFMNative]::GetFileId((Get-LongPath $A))
        $ib = [RHFMNative]::GetFileId((Get-LongPath $B))
        return ([bool]$ia -and [bool]$ib -and $ia -eq $ib)
    } catch { return $false }
}

# Returns a reason string when source and destination overlap (same folder, or one inside
# the other); '' when they are safe to use together. A MOVE onto itself would delete data.
function Test-PathOverlap {
    param([string]$Source, [string]$Destination)
    $s = Get-ComparablePath $Source
    $d = Get-ComparablePath $Destination
    if ($s -eq $d) { return 'The source and destination are the same folder.' }
    if ([System.IO.Directory]::Exists((Get-LongPath $Destination)) -and (Test-SamePhysicalItem $Source $Destination)) { return 'The source and destination are the same folder (reached through two different paths).' }
    if ($d.StartsWith($s + '\') -or $d.StartsWith($s + '/')) { return 'The destination is inside the source folder.' }
    if ($s.StartsWith($d + '\') -or $s.StartsWith($d + '/')) { return 'The source is inside the destination folder.' }
    return ''
}

# Folders that must never be wiped or bulk-deleted.
function Test-ProtectedFolder {
    param([string]$Path)
    $p = Get-ComparablePath $Path
    if ($p -match '^[a-z]:$') { return $true }
    if ($p -match '^\\\\[^\\]+\\[^\\]+$') { return $true }   # the root of a network share
    $protected = @($env:SystemRoot, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:USERPROFILE, $WorkDir)
    if ($env:USERPROFILE) { $protected += (Split-Path -Path $env:USERPROFILE -Parent) }
    foreach ($x in $protected) { if ($x -and ($p -eq (Get-ComparablePath $x))) { return $true } }
    return $false
}

function Get-SafeWindowsName {
    param([string]$Name)
    if ($null -eq $Name) { $Name = '' }
    $n = $Name -replace '[<>:"/\\|?*\x00-\x1F]', '-'
    $n = $n.TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($n)) { return "Unnamed_Item_$([guid]::NewGuid().ToString().Substring(0, 8))" }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($n)
    if ($base -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$') { $n = '_' + $n }
    return $n
}

# Makes a name unique inside one folder (case-insensitive), e.g. "Report (2).docx".
function Get-UniqueLocalName {
    param([string]$Name, [System.Collections.Generic.HashSet[string]]$Used)
    if ($Used.Add($Name)) { return $Name }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $ext  = [System.IO.Path]::GetExtension($Name)
    for ($i = 2; $i -lt 10000; $i++) {
        $cand = "$base ($i)$ext"
        if ($Used.Add($cand)) { return $cand }
    }
    return $Name
}

function Test-SameFileTime {
    param([datetime]$A, [datetime]$B)
    # FAT/exFAT and many NAS devices store times with 2-second precision.
    return ([Math]::Abs(($A.ToUniversalTime() - $B.ToUniversalTime()).TotalSeconds) -le 2)
}

function Test-IsLinkFolder {
    param([System.IO.DirectoryInfo]$Dir)
    if (-not ($Dir.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) { return $false }
    # OneDrive / cloud placeholders are reparse points too - only junctions and symbolic
    # links are treated as links (never followed, never emptied).
    if ($script:IsWin -and ('RHFMNative' -as [type])) {
        try {
            $tag = [RHFMNative]::GetReparseTag($Dir.FullName)
            if ($tag -eq [uint32]2684354563 -or $tag -eq [uint32]2684354572) { return $true }
            if ($tag -ne 0) { return $false }
        } catch { }
    }
    $lt = $null; $known = $false
    try { $lt = $Dir.LinkType; $known = $true } catch {
        try { $lt = (Get-Item -LiteralPath (Remove-LongPrefix $Dir.FullName) -Force -ErrorAction Stop).LinkType; $known = $true } catch { }
    }
    if ($known) { return ($lt -in @('Junction', 'SymbolicLink')) }
    # Could not tell: treat it as a link (never followed) unless it is a cloud placeholder.
    $cloud = [int]$Dir.Attributes -band (0x1000 -bor 0x40000 -bor 0x400000 -bor 0x100000)
    return ($cloud -eq 0)
}

# --------------------------------------------------------------------------------------
# Get-LocalTree: walks a folder tree folder-by-folder. Unreadable folders are recorded
# (not silently skipped), junctions / symbolic links are never followed, P / C work,
# and a heartbeat line shows progress on very large trees.
# Returns @{ Root; Files = List[FileInfo]; Dirs = List[DirectoryInfo]; Links; Errors; Bytes; Cancelled }
# --------------------------------------------------------------------------------------
function Get-LocalTree {
    param([string]$Path, [string]$Label = 'Scanning', [string]$Pattern = '*', [switch]$Quiet)
    $rootLong = Get-LongPath (Resolve-LocalPath $Path)
    $res = @{
        Root   = $rootLong
        Files  = (New-Object System.Collections.Generic.List[System.IO.FileInfo])
        Dirs   = (New-Object System.Collections.Generic.List[System.IO.DirectoryInfo])
        Links  = (New-Object System.Collections.Generic.List[System.IO.DirectoryInfo])
        Errors = (New-Object System.Collections.Generic.List[hashtable])
        Bytes  = [long]0
        Cancelled = $false
    }
    if (-not [System.IO.Directory]::Exists($rootLong)) {
        $res.Errors.Add(@{ File = (Split-Path -Path $Path -Leaf); Path = $Path; Reason = 'Folder not found' })
        return $res
    }
    $res.Root = (New-Object System.IO.DirectoryInfo($rootLong)).FullName
    $queue = New-Object System.Collections.Generic.Queue[System.IO.DirectoryInfo]
    $queue.Enqueue((New-Object System.IO.DirectoryInfo($res.Root)))
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $tick = [System.Diagnostics.Stopwatch]::StartNew()
    $shown = $false
    while ($queue.Count -gt 0) {
        $dir = $queue.Dequeue()
        if ((Check-Interrupt -CurrentFolder $dir.FullName) -eq 'CANCEL') { $res.Cancelled = $true; break }
        try {
            foreach ($f in $dir.EnumerateFiles($Pattern)) { $res.Files.Add($f); $res.Bytes += $f.Length }
            foreach ($d in $dir.EnumerateDirectories()) {
                if (Test-IsLinkFolder $d) { $res.Links.Add($d); continue }
                $res.Dirs.Add($d)
                $queue.Enqueue($d)
            }
        } catch {
            $res.Errors.Add(@{ File = $dir.Name; Path = (Remove-LongPrefix $dir.FullName); Reason = "Folder could not be read: $($_.Exception.Message)" })
        }
        if (-not $Quiet -and $tick.ElapsedMilliseconds -ge 1500) {
            $shown = $true
            Write-StatusLine -Text ("  [..] {0}: {1} files, {2} folders, {3} ({4})" -f $Label, (Format-Count $res.Files.Count), (Format-Count $res.Dirs.Count), (Format-FileSize $res.Bytes), (Format-Duration $sw.Elapsed.TotalSeconds))
            $tick.Restart()
        }
    }
    if ($shown) {
        Clear-StatusLine
        Write-Hint ("{0}: {1} files, {2} folders, {3}" -f $Label, (Format-Count $res.Files.Count), (Format-Count $res.Dirs.Count), (Format-FileSize $res.Bytes))
    }
    foreach ($e in $res.Errors) { Write-Log "UNREADABLE: $($e.Path) - $($e.Reason)" 'WARN' -FileOnly }
    foreach ($l in $res.Links)  { Write-Log "LINK SKIPPED (junction / symbolic link, not followed): $(Remove-LongPrefix $l.FullName)" 'INFO' -FileOnly }
    return $res
}

function Get-LocalDirectoryMetrics {
    param([string]$Path)
    $long = Get-LongPath $Path
    if (-not [System.IO.Directory]::Exists($long)) { return $null }
    $t = Get-LocalTree -Path $Path -Quiet
    return @{ Size = (Format-FileSize $t.Bytes); RawBytes = $t.Bytes; FileCount = $t.Files.Count; FolderCount = $t.Dirs.Count; Errors = $t.Errors.Count }
}

# LOCAL / MAP (mapped network drive) / UNC / OneDrive - shown in summaries.
function Get-PathStorageType {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = Remove-LongPrefix $Path
    if ($p -match '\\OneDrive(\s|-|_|\\|$)') { return 'OneDrive' }
    if ($p -match '^\\\\') { return 'UNC' }
    if ($p -match '^([A-Za-z]):') {
        try {
            $d = Get-PSDrive -Name $matches[1] -ErrorAction Stop
            if ($d.DisplayRoot -and $d.DisplayRoot -match '^\\\\') { return 'MAP' }
        } catch { }
        return 'LOCAL'
    }
    return ''
}

function Format-PathWithStorageTag {
    param([string]$Path)
    $t = Get-PathStorageType $Path
    if ($t) { return "[$t] $Path" }
    return $Path
}

# Free bytes available to this user on a local OR network path; -1 when unknown.
function Get-PathFreeSpace {
    param([string]$Path)
    $probe = Remove-LongPrefix (Resolve-LocalPath $Path)
    for ($i = 0; $i -lt 64 -and $probe; $i++) {
        if ([System.IO.Directory]::Exists((Get-LongPath $probe))) { break }
        $parent = Split-Path -Path $probe -Parent
        if (-not $parent -or $parent -eq $probe) { break }
        $probe = $parent
    }
    if ($script:IsWin -and ('RHFMNative' -as [type]) -and $probe) {
        try {
            [UInt64]$free = 0; [UInt64]$total = 0; [UInt64]$totalFree = 0
            $target = $probe
            if (-not $target.EndsWith('\')) { $target += '\' }
            if ([RHFMNative]::GetDiskFreeSpaceExW($target, [ref]$free, [ref]$total, [ref]$totalFree)) { return [long]$free }
        } catch { }
    }
    try {
        $root = [System.IO.Path]::GetPathRoot($probe)
        $drive = New-Object System.IO.DriveInfo($root)
        if ($drive.IsReady) { return [long]$drive.AvailableFreeSpace }
    } catch { }
    return [long]-1
}

function Show-DestinationFreeSpace {
    param([string]$DestPath)
    $free = Get-PathFreeSpace -Path $DestPath
    if ($free -ge 0) { Write-Log "Free space at destination: $(Format-FileSize $free)" 'INFO' }
    else { Write-Log "Free space at destination could not be determined." 'INFO' -FileOnly }
}

# Pre-flight storage check. -NeededBytes is what will actually be written (already net of
# files that are skipped because they are identical). Returns $true to continue.
function Invoke-PreflightStorageCheck {
    param([string]$DestRoot, [long]$NeededBytes)
    $free = Get-PathFreeSpace -Path $DestRoot
    if ($free -lt 0) {
        Write-Caution "Free space on the destination could not be determined."
        return (Read-YesNo -Question 'Continue anyway?' -Default $true)
    }
    Write-Log ("Storage check: need {0}, free {1} on destination." -f (Format-FileSize $NeededBytes), (Format-FileSize $free)) 'INFO'
    if ($NeededBytes -le $free) { return $true }
    Write-Panel -Title 'NOT ENOUGH FREE SPACE' -Color Red -Lines @(
        @{ Label = 'Needed';    Value = (Format-FileSize $NeededBytes); Level = 'ERROR' }
        @{ Label = 'Available'; Value = (Format-FileSize $free); Level = 'WARN' }
        'The job will probably stop part-way when the disk fills up.'
    )
    $go = Read-YesNo -Question 'Start anyway?' -Default $false
    Write-Log "Storage warning override: $go" 'WARN' -FileOnly
    return $go
}

# Deletes a file or folder (folders must already be empty unless -Recurse) using long
# paths, clearing read-only attributes when needed. Returns $true on success.
function Invoke-ForceDelete {
    param([string]$Path, [switch]$IsFolder, [switch]$Recurse)
    try {
        if ($IsFolder) { [System.IO.Directory]::Delete($Path, [bool]$Recurse) }
        else           { [System.IO.File]::Delete($Path) }
        return $true
    } catch {
        try {
            if ($IsFolder) {
                $di = New-Object System.IO.DirectoryInfo($Path)
                $di.Attributes = [System.IO.FileAttributes]::Normal
                [System.IO.Directory]::Delete($Path, [bool]$Recurse)
            } else {
                $fi = New-Object System.IO.FileInfo($Path)
                $fi.Attributes = [System.IO.FileAttributes]::Normal
                [System.IO.File]::Delete($Path)
            }
            return $true
        } catch { return $false }
    }
}

# After a MOVE: removes folders that are now EMPTY (deepest first). It never deletes
# recursively, so any file that was not moved (locked, too long, failed) stays put.
function Remove-EmptyDirectories {
    param([string]$TargetFolder, [switch]$IncludeRoot)
    $tree = Get-LocalTree -Path $TargetFolder -Quiet
    $dirs = @($tree.Dirs | Sort-Object { $_.FullName.Length } -Descending)
    $removed = 0
    foreach ($d in $dirs) {
        try {
            $hasEntries = [System.IO.Directory]::EnumerateFileSystemEntries($d.FullName).GetEnumerator().MoveNext()
            if (-not $hasEntries) { [System.IO.Directory]::Delete($d.FullName, $false); $removed++ }
        } catch { }
    }
    if ($IncludeRoot) {
        try {
            if (-not [System.IO.Directory]::EnumerateFileSystemEntries($tree.Root).GetEnumerator().MoveNext()) {
                [System.IO.Directory]::Delete($tree.Root, $false); $removed++
            }
        } catch { }
    }
    if ($removed -gt 0) { Write-Log "Removed $removed empty source folder(s) after the move." 'INFO' -Detail }
}

# Reveal a file or folder in Explorer WITHOUT running it.
function Show-InExplorer {
    param([string]$Path, [switch]$OpenFolder)
    try {
        $p = Remove-LongPrefix $Path
        if ($OpenFolder) { Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$p`"" -ErrorAction Stop }
        else { Start-Process -FilePath 'explorer.exe' -ArgumentList "/select,`"$p`"" -ErrorAction Stop }
        return $true
    } catch { Write-Bad "Could not open Explorer: $($_.Exception.Message)"; return $false }
}

function Set-ClipboardText {
    param([string]$Text)
    try { Set-Clipboard -Value $Text -ErrorAction Stop; return $true }
    catch {
        try { $Text | & clip.exe; return $true } catch { Write-Bad "Clipboard is not available: $($_.Exception.Message)"; return $false }
    }
}

# ======================================================================================
# SECTION 7 - SHAREPOINT / ONEDRIVE CORE
# --------------------------------------------------------------------------------------
# Authentication is exactly the method that has always worked for this script:
#   Connect-PnPOnline -Url <site> -UseWebLogin   (PnP.PowerShell 1.12.0)
# What is new: the connection is remembered and reused while you stay on the same site,
# and every listing / transfer call goes through one throttle-aware retry helper.
# ======================================================================================

$script:PnPRequiredVersion = '1.12.0'

# Loads (and if needed installs) PnP.PowerShell 1.12.0 the first time a cloud option is used.
function Ensure-PnPModule {
    if ($script:PnPReady) { return $true }
    $req = $script:PnPRequiredVersion
    $loaded = @(Get-Module -Name PnP.PowerShell)
    if ($loaded.Count -gt 0) {
        if (@($loaded | Where-Object { $_.Version.ToString() -eq $req }).Count -gt 0) { $script:PnPReady = $true; return $true }
        Write-Bad "PnP.PowerShell $($loaded[0].Version) is already loaded in this window, but version $req is required."
        Write-Hint "Close this PowerShell window, open a new one and start the script again."
        return $false
    }
    $available = @(Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version.ToString() -eq $req })
    if ($available.Count -eq 0) {
        Write-Info "One-time setup: installing PnP.PowerShell $req for your account. This can take a minute..."
        Write-Log "Installing PnP.PowerShell $req (CurrentUser scope)." 'WARN' -FileOnly
        try {
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop
            }
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Install-Module -Name PnP.PowerShell -RequiredVersion $req -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Bad "PnP.PowerShell could not be installed: $($_.Exception.Message)"
            Write-Hint "Check the internet connection, or ask IT to install PnP.PowerShell $req."
            Write-Log "PnP.PowerShell install failed: $($_.Exception.Message)" 'ERROR' -FileOnly
            return $false
        }
    }
    try {
        Import-Module PnP.PowerShell -RequiredVersion $req -DisableNameChecking -ErrorAction Stop -WarningAction SilentlyContinue
        $script:PnPReady = $true
        Write-Log "PnP.PowerShell $req loaded." 'INFO' -FileOnly
        return $true
    } catch {
        Write-Bad "PnP.PowerShell could not be loaded: $($_.Exception.Message)"
        Write-Log "PnP.PowerShell import failed: $($_.Exception.Message)" 'ERROR' -FileOnly
        return $false
    }
}

# Connects with -UseWebLogin (browser sign-in). Re-uses the open connection when it is
# already for the same site and still works. Throws when the connection fails.
function Connect-CloudSite {
    param([Parameter(Mandatory = $true)][string]$SiteUrl, [switch]$Force)
    if (-not (Ensure-PnPModule)) { throw "PnP.PowerShell $($script:PnPRequiredVersion) is not available." }
    $target = $SiteUrl.TrimEnd('/')
    if (-not $Force -and $script:PnPSiteUrl -and ($script:PnPSiteUrl -ieq $target)) {
        try { $null = Get-PnPWeb -ErrorAction Stop; return } catch {
            Write-Log "The saved connection to $target has expired - signing in again." 'WARN'
        }
    }
    Write-Info "Connecting to $target"
    Write-Hint "A Microsoft sign-in window may appear - sign in, then come back to this window."
    $oldWarn = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    try {
        Connect-PnPOnline -Url $target -UseWebLogin -ErrorAction Stop
    } catch {
        $script:PnPSiteUrl = $null
        Write-Log "Connection to $target failed: $($_.Exception.Message)" 'ERROR' -FileOnly
        throw
    } finally {
        $WarningPreference = $oldWarn
    }
    $script:PnPSiteUrl = $target
    Write-Log "Connected to $target" 'SUCCESS'
}

function Disconnect-CloudSite {
    if ($script:PnPSiteUrl) {
        try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch { }
        Write-Log "Disconnected from $($script:PnPSiteUrl)" 'INFO' -FileOnly
        $script:PnPSiteUrl = $null
    }
}

# Turns any SharePoint / OneDrive browser address into @{ SiteUrl; FolderPath; Note }.
# Understands /sites/, /teams/ and /personal/ sites, ?id= / ?RootFolder= links (including
# names with & or #), view pages (Forms/AllItems.aspx), viewid= and "/:f:/r/" links.
function Parse-CloudUrl {
    param([string]$Url, [switch]$Quiet)
    $u = ConvertTo-CleanPath $Url
    if ($u -notmatch '^(?i)https?://[^/\s]+') {
        if (-not $Quiet) { Write-Bad "That is not a web address. SharePoint / OneDrive links start with https://" }
        return $null
    }
    try {
        $uri = New-Object System.Uri($u)
        $authority = 'https://' + $uri.Host
        $serverRel = $null
        $note = ''
        if ($uri.Query) {
            foreach ($pair in $uri.Query.TrimStart('?').Split('&')) {
                $kv = $pair.Split(@('='), 2, [System.StringSplitOptions]::None)
                if ($kv.Count -eq 2 -and ($kv[0] -ieq 'id' -or $kv[0] -ieq 'RootFolder') -and $kv[1]) {
                    $serverRel = [System.Uri]::UnescapeDataString($kv[1])
                    break
                }
            }
        }
        if (-not $serverRel) {
            $rawPath = $uri.AbsolutePath
            if ($rawPath -match '^/:[a-zA-Z]:/[a-qs-zA-QS-Z]/') {
                if (-not $Quiet) {
                    Write-Bad "That is a sharing link, which does not contain the folder path."
                    Write-Hint "Open the link in your browser, wait for the folder to load, then copy the address bar."
                }
                return $null
            }
            $rawPath = [regex]::Replace($rawPath, '^/:[a-zA-Z]:/[rR]/', '/')
            $serverRel = [System.Uri]::UnescapeDataString($rawPath)
        }
        if (-not $serverRel.StartsWith('/')) { $serverRel = '/' + $serverRel }
        $serverRel = [regex]::Replace($serverRel, '(?i)/Forms/[^/]+\.aspx$', '')
        $site = ''
        if ($serverRel -match '^(?i)(/(sites|teams|personal)/[^/]+)') { $site = $matches[1] }
        if ($serverRel -match '(?i)/_layouts/') {
            $lib = if ($site -match '^(?i)/personal/') { 'Documents' } else { 'Shared Documents' }
            $serverRel = "$site/$lib"
            $note = "That link is a page, not a folder - using the default '$lib' library. Paste a folder link if you meant a different folder."
        } elseif ($serverRel -match '(?i)\.aspx$') {
            $note = "That link points to a web page (.aspx), not a folder or file. Open the folder in the browser and copy that address instead."
        }
        $serverRel = $serverRel.TrimEnd('/')
        if (-not $serverRel) { $serverRel = '/' }
        return @{ SiteUrl = ($authority + $site); FolderPath = $serverRel; Note = $note }
    } catch {
        if (-not $Quiet) { Write-Bad "The address could not be understood: $($_.Exception.Message)" }
        return $null
    }
}

function Get-SitePath {
    param([string]$SiteUrl)
    try { return ([System.Uri]$SiteUrl).AbsolutePath.TrimEnd('/') } catch { return '' }
}

# Server-relative (/sites/x/Shared Documents/A) -> site-relative (Shared Documents/A).
function Get-SiteRelativePath {
    param([string]$SiteUrl, [string]$ServerRelativePath)
    $sp = Get-SitePath $SiteUrl
    $fp = "$ServerRelativePath"
    if ($sp -and $fp.StartsWith($sp, [System.StringComparison]::OrdinalIgnoreCase)) { return $fp.Substring($sp.Length).Trim('/') }
    return $fp.Trim('/')
}

function Join-CloudPath {
    param([string]$Base, [string]$Child)
    if ([string]::IsNullOrEmpty($Child)) { return $Base.TrimEnd('/') }
    if ([string]::IsNullOrEmpty($Base))  { return $Child.Trim('/') }
    return ($Base.TrimEnd('/') + '/' + $Child.TrimStart('/'))
}

# Browser link for a server-relative path (every segment encoded, so # % & work).
function Get-CloudWebUrl {
    param([string]$SiteUrl, [string]$ServerRelativePath, [switch]$IsFile)
    $authority = ([System.Uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $enc = (($ServerRelativePath -split '/') | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/'
    $url = $authority + $enc
    if ($IsFile) { $url += '?web=1' }
    return $url
}

function Test-IsCloudFolder { param($Item) return ($null -ne $Item -and $Item.GetType().Name -eq 'Folder') }

# Reads a CSOM property that may not be loaded (returns $Default instead of throwing).
function Get-CloudProp {
    param($Item, [string]$Name, $Default = $null)
    try { $v = $Item.$Name; if ($null -ne $v) { return $v } } catch { }
    return $Default
}

# Site-relative root folders of every list / library on a site (cached per site).
$script:LibraryRootCache = @{}
function Get-CloudLibraryRoots {
    param([string]$SiteUrl = '')
    if (-not $SiteUrl) { $SiteUrl = $script:PnPSiteUrl }
    if (-not $SiteUrl) { return $null }
    $key = $SiteUrl.TrimEnd('/').ToLowerInvariant()
    if ($script:LibraryRootCache.ContainsKey($key)) { return ,$script:LibraryRootCache[$key] }
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        $glrLists = Invoke-WithRetry -OperationName 'Read libraries' -Quiet -ScriptBlock { Get-PnPList -ErrorAction Stop }
        foreach ($l in @($glrLists)) {
            $u = Get-CloudProp (Get-CloudProp $l 'RootFolder') 'ServerRelativeUrl'
            if ($u) { $null = $set.Add((Get-SiteRelativePath -SiteUrl $SiteUrl -ServerRelativePath $u)) }
        }
    } catch { $set = $null }
    $script:LibraryRootCache[$key] = $set
    return ,$set
}

# True for the root folder of a library / list (or a site): such folders are never
# deleted - only their contents. Libraries of sub-sites ("subsite/Library") are covered too.
function Test-IsLibraryRootPath {
    param([string]$SiteUrl, [string]$SiteRelativePath)
    $rel = "$SiteRelativePath".Trim('/')
    if ($rel -eq '' -or $rel -notmatch '/') { return $true }
    $roots = Get-CloudLibraryRoots -SiteUrl $SiteUrl
    if ($null -eq $roots) { return $false }
    if ($roots.Contains($rel)) { return $true }
    $segs = $rel -split '/'
    if ($segs.Count -eq 2 -and -not $roots.Contains($segs[0])) { return $true }
    return $false
}

# The hidden "Forms" folder at the root of a document library holds the library's view
# pages. It is skipped everywhere (it cannot be deleted and contains no user files). Only
# a Forms folder directly inside a KNOWN library root is skipped - a user folder that
# happens to be called "Forms" deeper down is always included.
function Test-IsSystemFolder {
    param($Item, [string]$ParentSiteRel)
    if ($Item.Name -ne 'Forms' -or -not (Test-IsCloudFolder $Item)) { return $false }
    $parent = "$ParentSiteRel".Trim('/')
    if ($parent -eq '') { return $false }
    if ($parent -notmatch '/') { return $true }
    $roots = Get-CloudLibraryRoots
    return ($null -ne $roots -and $roots.Contains($parent))
}

# Errors that will never succeed on retry - fail fast instead of waiting a minute.
$script:PermanentErrorPatterns = @(
    '\(404\)', '(?i)not found', '(?i)does not exist', '\(401\)', '\(403\)', '(?i)access (is )?denied',
    '(?i)unauthori[sz]ed', '(?i)Forms directory of a document library', '(?i)checked out',
    '(?i)locked for (shared )?use', '(?i)already exists', '(?i)blocked file', '(?i)is not a valid',
    '(?i)path.{0,30}too long'
)

# Runs a script block with throttle-aware retries (honours Retry-After, backs off, and
# lets the user pause / cancel during the wait). Returns the script block's output.
# Local variable names are deliberately unusual: script blocks see this function's
# variables, so common names here could shadow the caller's.
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName = 'Cloud operation',
        [int]$MaxAttempts = 5,
        [string[]]$FastFailPatterns = @(),
        [switch]$Quiet
    )
    $iwrAttempt = 0
    while ($true) {
        $iwrAttempt++
        try {
            return (& $ScriptBlock)
        } catch {
            $iwrMsg = $_.Exception.Message
            if ($_.Exception -is [System.OperationCanceledException]) { throw }
            $iwrStop = $false
            foreach ($iwrPat in (@($FastFailPatterns) + @($script:PermanentErrorPatterns))) {
                if ($iwrPat -and $iwrMsg -match $iwrPat) { $iwrStop = $true; break }
            }
            if ($iwrStop -or $iwrAttempt -ge $MaxAttempts -or $script:JobCancelled) { throw }
            $iwrThrottle = ($iwrMsg -match '(?i)429|503|throttl|too many requests|service unavailable|server is busy|timed? ?out|temporar')
            $iwrWait = if ($iwrThrottle) { 10 * $iwrAttempt } else { 3 * $iwrAttempt }
            if ($iwrMsg -match '(?i)retry.?after\D{0,12}(\d{1,3})') { $iwrWait = [Math]::Min(120, [Math]::Max(1, [int]$matches[1])) }
            $iwrWhat = if ($iwrThrottle) { 'is being throttled by Microsoft' } else { 'failed' }
            Write-Log ("{0} {1} (attempt {2}/{3}): {4}" -f $OperationName, $iwrWhat, $iwrAttempt, $MaxAttempts, $iwrMsg) 'WARN' -FileOnly
            if (-not $Quiet -or $iwrThrottle) {
                Write-Caution ("{0} {1} - retrying ({2}/{3})" -f (Get-FitText -Text $OperationName -Width 50), $iwrWhat, ($iwrAttempt + 1), $MaxAttempts)
            }
            Start-InterruptibleSleep -Seconds $iwrWait -Label 'Waiting before retry'
            if ($script:JobCancelled) { throw (New-Object System.OperationCanceledException('Cancelled by the user during a retry wait.')) }
        }
    }
}

# Lists one folder (files + sub-folders). Throws if the folder cannot be listed after
# retries - callers count that as an error instead of treating the folder as empty.
function Get-CloudFolderItems {
    param([string]$SiteRelativeUrl, [switch]$IncludeSystem)
    $gcfRaw = Invoke-WithRetry -OperationName "List /$SiteRelativeUrl" -Quiet -ScriptBlock {
        Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeUrl -ErrorAction Stop
    }
    $gcfOut = New-Object System.Collections.Generic.List[object]
    foreach ($gcfItem in @($gcfRaw)) {
        if ($null -eq $gcfItem) { continue }
        if (-not $IncludeSystem -and (Test-IsSystemFolder -Item $gcfItem -ParentSiteRel $SiteRelativeUrl)) { continue }
        $gcfOut.Add($gcfItem)
    }
    return ,$gcfOut.ToArray()
}

# Returns the file object when the site-relative path is a FILE, otherwise $null.
function Get-CloudFileIfFile {
    param([string]$SiteRelativeUrl)
    if ([string]::IsNullOrWhiteSpace($SiteRelativeUrl)) { return $null }
    try {
        $f = Get-PnPFile -Url $SiteRelativeUrl -ErrorAction Stop
        if ($f -and (Get-CloudProp $f 'Name')) { return $f }
    } catch { }
    return $null
}

# Sorted listing for the batch screens: folders first, then files, by name.
function Get-CloudIndexItems {
    param([string]$SiteRelativeUrl)
    $items = Get-CloudFolderItems -SiteRelativeUrl $SiteRelativeUrl
    return ,@($items | Sort-Object @{ Expression = { -not (Test-IsCloudFolder $_) } }, @{ Expression = { $_.Name } })
}

# Full recursive inventory of a cloud folder (or a single file).
# Returns @{ Files = List[@{RelPath;Name;Length;ServerRelativeUrl;Modified}]; Folders; Errors;
#            Bytes; IsFile; Cancelled; Incomplete }. Incomplete = something could not be listed.
function Get-CloudInventory {
    param([string]$SiteUrl, [string]$FolderPath, [string]$Tag = 'Scanning cloud', [switch]$Quiet)
    $res = @{
        Files   = (New-Object System.Collections.Generic.List[hashtable])
        Folders = (New-Object System.Collections.Generic.List[string])
        Errors  = (New-Object System.Collections.Generic.List[hashtable])
        Bytes = [long]0; IsFile = $false; Cancelled = $false; Incomplete = $false
    }
    $base = Get-SiteRelativePath -SiteUrl $SiteUrl -ServerRelativePath $FolderPath
    $single = Get-CloudFileIfFile -SiteRelativeUrl $base
    if ($single) {
        $len = [long](Get-CloudProp $single 'Length' 0)
        $res.Files.Add(@{ RelPath = $single.Name; Name = $single.Name; Length = $len; ServerRelativeUrl = $FolderPath; Modified = (Get-CloudProp $single 'TimeLastModified') })
        $res.Bytes = $len; $res.IsFile = $true
        return $res
    }
    $queue = New-Object System.Collections.Generic.Queue[hashtable]
    $queue.Enqueue(@{ Site = $base; Rel = '' })
    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $tick = [System.Diagnostics.Stopwatch]::StartNew(); $shown = $false
    while ($queue.Count -gt 0) {
        $cur = $queue.Dequeue()
        if ((Check-Interrupt -CurrentFolder "/$($cur.Site)") -eq 'CANCEL') { $res.Cancelled = $true; break }
        Start-Sleep -Milliseconds 100   # gentle pacing between folders (anti-throttling)
        try { $items = Get-CloudFolderItems -SiteRelativeUrl $cur.Site }
        catch {
            if ($_.Exception -is [System.OperationCanceledException]) { $res.Cancelled = $true; break }
            $res.Errors.Add(@{ File = $(if ($cur.Rel) { $cur.Rel } else { '(top folder)' }); Path = "/$($cur.Site)"; Reason = "Folder could not be listed: $($_.Exception.Message)" })
            continue
        }
        foreach ($it in $items) {
            $rel = if ($cur.Rel) { "$($cur.Rel)/$($it.Name)" } else { $it.Name }
            if (Test-IsCloudFolder $it) {
                $res.Folders.Add($rel)
                $queue.Enqueue(@{ Site = (Join-CloudPath $cur.Site $it.Name); Rel = $rel })
            } else {
                $len = [long](Get-CloudProp $it 'Length' 0)
                $srv = Get-CloudProp $it 'ServerRelativeUrl' (Join-CloudPath $FolderPath $rel)
                $res.Files.Add(@{ RelPath = $rel; Name = $it.Name; Length = $len; ServerRelativeUrl = $srv; Modified = (Get-CloudProp $it 'TimeLastModified') })
                $res.Bytes += $len
            }
        }
        if (-not $Quiet -and $tick.ElapsedMilliseconds -ge 1500) {
            $shown = $true
            Write-StatusLine -Text ("  [..] {0}: {1} files in {2} folders, {3} ({4})" -f $Tag, (Format-Count $res.Files.Count), (Format-Count ($res.Folders.Count + 1)), (Format-FileSize $res.Bytes), (Format-Duration $sw.Elapsed.TotalSeconds))
            $tick.Restart()
        }
    }
    if ($shown) { Clear-StatusLine }
    if (-not $Quiet) {
        Write-Hint ("{0}: {1} files in {2} folders, {3}" -f $Tag, (Format-Count $res.Files.Count), (Format-Count ($res.Folders.Count + 1)), (Format-FileSize $res.Bytes))
    }
    $res.Incomplete = ($res.Errors.Count -gt 0 -or $res.Cancelled)
    foreach ($e in $res.Errors) { Write-Log "LISTING ERROR $($e.Path): $($e.Reason)" 'WARN' }
    return $res
}

# Server-side copy / move (Copy-PnPFile / Move-PnPFile). PnP 1.12 uses SharePoint copy
# jobs for folders and cross-site items; a failed job is returned as data, not thrown,
# so the job log is inspected and turned into an error here.
function Invoke-CloudCopyMove {
    param([string]$SourceUrl, [string]$TargetUrl, [switch]$Move, [string]$Name = '')
    $ccmVerb = if ($Move) { 'Move' } else { 'Copy' }
    $null = Invoke-WithRetry -OperationName "$ccmVerb $Name" -ScriptBlock {
        if ($Move) { $ccmOut = Move-PnPFile -SourceUrl $SourceUrl -TargetUrl $TargetUrl -Force -Overwrite -ErrorAction Stop }
        else       { $ccmOut = Copy-PnPFile -SourceUrl $SourceUrl -TargetUrl $TargetUrl -Force -Overwrite -ErrorAction Stop }
        $ccmErrors = @(@($ccmOut) | Where-Object { $_ -and $_.PSObject.Properties['Event'] -and ("$($_.Event)" -match '(?i)error|fatal|cancel') })
        if ($ccmErrors.Count -gt 0) {
            $ccmText = ($ccmErrors | ForEach-Object { if ($_.PSObject.Properties['Message']) { "$($_.Message)" } else { "$_" } }) -join '; '
            throw "SharePoint reported that the $($ccmVerb.ToLower()) job failed: $ccmText"
        }
    }
}

# Sends one cloud file or folder to the site Recycle Bin. Refuses to remove the root of
# a document library (that would remove the whole library).
function Remove-CloudItem {
    param([string]$SiteUrl, [string]$ServerRelativeUrl, [switch]$IsFolder)
    $rciSrv = $ServerRelativeUrl.TrimEnd('/')
    $rciRel = Get-SiteRelativePath -SiteUrl $SiteUrl -ServerRelativePath $rciSrv
    if ($IsFolder -and (Test-IsLibraryRootPath -SiteUrl $SiteUrl -SiteRelativePath $rciRel)) {
        throw "Refusing to delete '$rciSrv': it is the root of a document library or site. Its contents can be deleted instead."
    }
    $rciCut    = $rciSrv.LastIndexOf('/')
    $rciParent = $rciSrv.Substring(0, $rciCut)
    $rciLeaf   = $rciSrv.Substring($rciCut + 1)
    if ($IsFolder) {
        $null = Invoke-WithRetry -OperationName "Delete folder $rciLeaf" -ScriptBlock {
            $null = Remove-PnPFolder -Name $rciLeaf -Folder $rciParent -Force -Recycle -ErrorAction Stop
        }
    } else {
        $null = Invoke-WithRetry -OperationName "Delete file $rciLeaf" -ScriptBlock {
            $null = Remove-PnPFile -ServerRelativeUrl $rciSrv -Force -Recycle -ErrorAction Stop
        }
    }
}

# Pulls "file X checked out by Y" out of SharePoint's error text.
function Get-CheckoutInfo {
    param([string]$Message)
    $file = if ($Message -match '(?i)The file "([^"]+)"') { Split-Path -Path $matches[1] -Leaf } else { 'a file inside' }
    $user = if ($Message -match '(?i)by\s+i:0#\.f\|membership\|(\S+)') { $matches[1].TrimEnd('.') }
            elseif ($Message -match '(?i)by\s+(.{3,60}?)\.?$') { $matches[1].Trim() }
            else { 'another user' }
    return @{ File = $file; User = $user }
}

# Index rows for a cloud listing (used by every cloud batch screen).
function ConvertTo-CloudIndexRows {
    param([object[]]$Items)
    $rows = New-Object System.Collections.Generic.List[hashtable]
    foreach ($it in $Items) {
        $mod = Get-CloudProp $it 'TimeLastModified'
        $modText = if ($mod) { ([datetime]$mod).ToString('yyyy-MM-dd') } else { '' }
        if (Test-IsCloudFolder $it) {
            $cnt = Get-CloudProp $it 'ItemCount'
            $info = if ($null -ne $cnt) { "{0:N0} items" -f [long]$cnt } else { '' }
            $rows.Add(@{ Kind = 'FOLDER'; Name = $it.Name; Info = $info; Date = $modText; Bytes = -1 })
        } else {
            $len = [long](Get-CloudProp $it 'Length' 0)
            $rows.Add(@{ Kind = 'FILE'; Name = $it.Name; Info = (Format-FileSize $len); Date = $modText; Bytes = $len })
        }
    }
    return ,$rows.ToArray()
}

# ======================================================================================
# SECTION 8 - SHARED SCREENS: ITEM INDEX, SUMMARIES, SEARCH RESULTS
# ======================================================================================

# Table used by every batch screen. Rows: @{ Kind = FOLDER|FILE; Name; Info; Date; Bytes }
function Show-ItemIndex {
    param([string]$Title, [object[]]$Rows, [string]$Color = 'Cyan', [string]$InfoHeader = 'SIZE / ITEMS')
    $W = Get-UiWidth
    $numW  = [Math]::Max(2, ([string]$Rows.Count).Length)
    $infoW = 14; $dateW = 10; $barW = 8
    $showDate = $true; $showBar = $true
    $nameW = $W - 2 - ($numW + 2) - 8 - ($infoW + 1) - ($dateW + 1) - ($barW + 1)
    if ($nameW -lt 18) { $showBar = $false; $nameW += ($barW + 1) }
    if ($nameW -lt 18) { $showDate = $false; $nameW += ($dateW + 1) }
    $maxBytes = [long]0
    foreach ($r in $Rows) { if ($r.Bytes -gt $maxBytes) { $maxBytes = [long]$r.Bytes } }
    if ($maxBytes -le 0) { $showBar = $false }

    Write-Host ""
    Write-Rule -Title $Title -Color $Color
    $hdr = ("  {0} {1} {2} {3}" -f ('#'.PadLeft($numW + 1)), 'TYPE   ', 'NAME'.PadRight($nameW), $InfoHeader.PadLeft($infoW))
    if ($showDate) { $hdr += ' ' + 'MODIFIED'.PadRight($dateW) }
    Write-Host $hdr -ForegroundColor DarkGray
    $audit = New-Object System.Collections.Generic.List[string]
    $audit.Add("--- $Title ---")
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $r = $Rows[$i]
        $isFolder = ($r.Kind -eq 'FOLDER')
        Write-Host ("  {0} " -f ("{0}." -f ($i + 1)).PadLeft($numW + 1)) -ForegroundColor Cyan -NoNewline
        if ($isFolder) { Write-Host 'FOLDER ' -ForegroundColor Yellow -NoNewline } else { Write-Host 'FILE   ' -ForegroundColor DarkCyan -NoNewline }
        Write-Host (' ' + (Get-FitText -Text $r.Name -Width $nameW).PadRight($nameW)) -ForegroundColor White -NoNewline
        Write-Host (' ' + (Get-FitText -Text ([string]$r.Info) -Width $infoW).PadLeft($infoW)) -ForegroundColor Gray -NoNewline
        if ($showDate) { Write-Host (' ' + ([string]$r.Date).PadRight($dateW)) -ForegroundColor DarkGray -NoNewline }
        if ($showBar -and $r.Bytes -gt 0) { Write-Host (' ' + (Get-BarText -Fraction ($r.Bytes / $maxBytes) -Width $barW)) -ForegroundColor DarkCyan -NoNewline }
        Write-Host ""
        $audit.Add(("[{0}] [{1}] {2} [{3}] {4}" -f ($i + 1), $r.Kind, $r.Name, $r.Info, $r.Date))
    }
    Write-LogBlock -Lines $audit.ToArray() -Level 'MENU'
}

function Write-BatchItemHeader {
    param([int]$Index, [int]$Total, [string]$Name, [string]$Verb, [string]$Color = 'Magenta')
    Write-Host ""
    Write-Rule -Title ("[{0}/{1}] {2}  {3}" -f $Index, $Total, $Verb, $Name) -Color $Color
    Write-Log ("BATCH ITEM {0}/{1} {2}: {3}" -f $Index, $Total, $Verb, $Name) 'INFO' -FileOnly
}

# Failure rate, distinct error messages and a file-type breakdown under a summary.
function Show-OperationDiagnostics {
    param([hashtable]$Stats)
    $total = $Stats.Success + $Stats.Failed
    if ($total -le 0) { return }
    $lines = New-Object System.Collections.Generic.List[object]
    $color = 'DarkGray'
    if ($Stats.Failed -gt 0) {
        $rate = $Stats.Failed / $total
        $color = if ($rate -ge 0.25) { 'Red' } elseif ($rate -ge 0.10) { 'Yellow' } else { 'DarkYellow' }
        $lines.Add(@{ Text = ("Failure rate: {0:N1}% ({1} of {2})" -f ($rate * 100), (Format-Count $Stats.Failed), (Format-Count $total)); Color = $color })
        $lines.Add('Distinct errors (most frequent first):')
        foreach ($e in (Get-DistinctErrors -FailedList $Stats.FailedItems -Top 8)) { $lines.Add(@{ Text = "  $e"; Color = 'Gray' }) }
    }
    if ($Stats.Success -ge 50 -and $Stats.Names.Count -gt 0) {
        if ($lines.Count -gt 0) { $lines.Add('-') }
        $lines.Add(@{ Text = "By file type: $(Get-TopExtensions -Names $Stats.Names)"; Color = 'Gray' })
    }
    if ($lines.Count -eq 0) { return }
    $title = if ($Stats.Failed -gt 0) { 'DIAGNOSTICS' } else { 'DETAILS' }
    Write-Panel -Title $title -Lines $lines.ToArray() -Color $color
}

# One summary for every transfer engine (copy, move, download, upload).
function Show-TransferSummary {
    param(
        [string]$Title, [hashtable]$Stats, [TimeSpan]$Elapsed, [string]$Target = '', [string]$Storage = '',
        [string]$Color = 'Green', [string]$OkLabel = 'Transferred', [object[]]$ExtraRows = @()
    )
    if ($Stats.FailedItems.Count -gt 0) { Write-FailedReport -FailedList $Stats.FailedItems }
    $rows = New-Object System.Collections.Generic.List[object]
    if ($Target)  { $rows.Add(@{ Label = 'Target';  Value = $Target }) }
    if ($Storage) { $rows.Add(@{ Label = 'Storage'; Value = $Storage }) }
    $rows.Add(@{ Label = $OkLabel; Value = ("{0} files  |  {1}" -f (Format-Count $Stats.Success), (Format-FileSize $Stats.SuccessBytes)); Level = 'SUCCESS' })
    $rows.Add(@{ Label = 'Skipped (already there)'; Value = ("{0} files  |  {1}" -f (Format-Count $Stats.Skipped), (Format-FileSize $Stats.SkippedBytes)); Level = 'INFO' })
    $rows.Add(@{ Label = 'Failed'; Value = (Format-Count $Stats.Failed); Level = $(if ($Stats.Failed -gt 0) { 'ERROR' } else { 'SUCCESS' }) })
    if ($Stats.Unreadable -gt 0)    { $rows.Add(@{ Label = 'Unreadable folders'; Value = "$($Stats.Unreadable) (see failed items)"; Level = 'ERROR' }) }
    if ($Stats.ListingErrors -gt 0) { $rows.Add(@{ Label = 'Folders not listed'; Value = "$($Stats.ListingErrors) (see failed items)"; Level = 'ERROR' }) }
    if ($Stats.TooLong -gt 0)       { $rows.Add(@{ Label = 'Skipped (path too long)'; Value = "$($Stats.TooLong) (listed in LongPaths log)"; Level = 'WARN' }) }
    if ($Stats.Accepted -gt 0)      { $rows.Add(@{ Label = 'Office size differences'; Value = "$($Stats.Accepted) (SharePoint metadata - accepted)"; Level = 'WARN' }) }
    foreach ($x in $ExtraRows) { $rows.Add($x) }
    $rows.Add(@{ Label = 'Time taken'; Value = (Format-Elapsed $Elapsed) })
    $spd = Format-AvgSpeed -Bytes $Stats.SuccessBytes -Elapsed $Elapsed
    if ($spd) { $rows.Add(@{ Label = 'Average speed'; Value = $spd }) }
    $fpm = Format-FilesPerMin -Count $Stats.Success -Elapsed $Elapsed
    if ($fpm -and $Stats.Success -ge 50) { $rows.Add(@{ Label = 'Throughput'; Value = $fpm }) }
    if ($script:SessionBytes -gt 0) { $rows.Add(@{ Label = 'This session so far'; Value = (Format-FileSize $script:SessionBytes) }) }

    $bad = ($Stats.Failed -gt 0 -or $Stats.Unreadable -gt 0 -or $Stats.ListingErrors -gt 0)
    if ($Stats.Cancelled) { $state = 'CANCELLED'; $boxColor = 'Yellow' }
    elseif ($bad)         { $state = 'FINISHED WITH ERRORS'; $boxColor = 'Red' }
    else                  { $state = 'COMPLETE'; $boxColor = $Color }
    Write-Host ""
    Write-SummaryBox -Title "$Title - $state" -Lines $rows.ToArray() -BoxColor $boxColor
    Show-OperationDiagnostics -Stats $Stats
    if (-not $Stats.Cancelled -and -not $bad -and ($Stats.Success + $Stats.Skipped) -gt 0) {
        Write-Host ("  " + (Get-BarText -Fraction 1 -Width 12) + "  All clear - every file is in place, no errors.") -ForegroundColor Green
    }
}

# --------------------------------------------------------------------------------------
# Search results (local and cloud)
# Each result: @{ Type = FILE|FOLDER; Name; Path; Parent; Size; Modified; Url; ParentUrl }
# --------------------------------------------------------------------------------------
function Show-SearchResults {
    param([object[]]$Results, [string]$Term, [ValidateSet('Local', 'Cloud')][string]$Kind)
    if ($Results.Count -eq 0) { Write-Caution "No matches for '$Term'."; return }
    $max = 200
    while ($true) {
        $W = Get-UiWidth
        Write-Host ""
        Write-Rule -Title ("{0} MATCHES FOR '{1}'" -f (Format-Count $Results.Count), $Term) -Color Cyan
        $numW = ([string][Math]::Min($Results.Count, $max)).Length + 1
        $pathW = [Math]::Max(10, $W - 2 - $numW - 1 - 7 - 25)
        for ($i = 0; $i -lt [Math]::Min($Results.Count, $max); $i++) {
            $r = $Results[$i]
            $meta = @()
            if ($r.Type -eq 'FILE') { $meta += (Format-FileSize ([long]$r.Size)) }
            if ($r.Modified) { try { $meta += ([datetime]$r.Modified).ToString('yyyy-MM-dd') } catch { } }
            Write-Host ("  {0} " -f ("$($i + 1).").PadLeft($numW)) -ForegroundColor Cyan -NoNewline
            if ($r.Type -eq 'FOLDER') { Write-Host 'FOLDER ' -ForegroundColor Yellow -NoNewline } else { Write-Host 'FILE   ' -ForegroundColor DarkCyan -NoNewline }
            Write-Host (Get-FitText -Text $r.Path -Width $pathW -Middle).PadRight($pathW) -ForegroundColor White -NoNewline
            Write-Host (' ' + (Get-FitText -Text ($meta -join '  ') -Width 24)) -ForegroundColor DarkGray
        }
        if ($Results.Count -gt $max) { Write-Caution ("Showing the first {0}. Export to CSV to see all {1}." -f $max, (Format-Count $Results.Count)) }
        $opt = (Read-Line -Prompt '  Number = more options  |  E = export CSV  |  ENTER = done > ' -LogLabel 'Search results').Trim()
        if (-not $opt -or $opt -match '^(?i)q$') { return }
        if ($opt -match '^(?i)e$') {
            $csv = Join-Path $LogFileDirectory ("{0}Search_{1}.csv" -f $Kind, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            try {
                $Results | ForEach-Object {
                    [pscustomobject]@{
                        Type = $_.Type; Name = $_.Name; Path = $_.Path; ParentFolder = $_.Parent
                        Size = $(if ($_.Type -eq 'FILE') { Format-FileSize ([long]$_.Size) } else { '' })
                        SizeBytes = $(if ($_.Type -eq 'FILE') { [long]$_.Size } else { '' })
                        Modified = $(if ($_.Modified) { try { ([datetime]$_.Modified).ToString('yyyy-MM-dd HH:mm') } catch { '' } } else { '' })
                        WebLink = $_.Url
                    }
                } | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
                Write-Ok "Exported $(Format-Count $Results.Count) results to $csv"
                if (Read-YesNo -Question 'Open the CSV now?' -Default $false) { try { Invoke-Item -LiteralPath $csv -ErrorAction Stop } catch { Write-Bad "Could not open it: $($_.Exception.Message)" } }
            } catch { Write-Bad "Export failed: $($_.Exception.Message)" }
            continue
        }
        if ($opt -match '^\d{1,6}$' -and [int]$opt -ge 1 -and [int]$opt -le [Math]::Min($Results.Count, $max)) {
            $r = $Results[[int]$opt - 1]
            $rows = @(@{ Label = 'Name'; Value = $r.Name }, @{ Label = 'Type'; Value = $r.Type })
            if ($Kind -eq 'Local') { $rows += @{ Label = 'Full path'; Value = $r.Path }; $rows += @{ Label = 'Parent folder'; Value = $r.Parent } }
            else { $rows += @{ Label = 'Web link'; Value = $r.Url }; $rows += @{ Label = 'Parent folder'; Value = $r.ParentUrl } }
            if ($r.Type -eq 'FILE') { $rows += @{ Label = 'Size'; Value = (Format-FileSize ([long]$r.Size)) } }
            if ($r.Modified) { try { $rows += @{ Label = 'Modified'; Value = ([datetime]$r.Modified).ToString('yyyy-MM-dd HH:mm') } } catch { } }
            Write-Panel -Title 'MATCH DETAILS' -Lines $rows -Color Magenta
            while ($true) {
                if ($Kind -eq 'Local') { $a = Read-Choice -Prompt 'S = show in Explorer   F = open its folder   C = copy path   ENTER = back' -Valid @('S', 'F', 'C') }
                else { $a = Read-Choice -Prompt 'O = open in browser   F = open its folder   C = copy link   ENTER = back' -Valid @('O', 'F', 'C') }
                if (-not $a) { break }
                switch ($a) {
                    'S' { if (Show-InExplorer -Path $r.Path) { Write-Ok 'Shown in Explorer.' } }
                    'F' {
                        if ($Kind -eq 'Local') { if (Show-InExplorer -Path $r.Parent -OpenFolder) { Write-Ok 'Folder opened.' } }
                        else { try { Start-Process $r.ParentUrl -ErrorAction Stop; Write-Ok 'Opened in your browser.' } catch { Write-Bad $_.Exception.Message } }
                    }
                    'O' { try { Start-Process $r.Url -ErrorAction Stop; Write-Ok 'Opened in your browser.' } catch { Write-Bad $_.Exception.Message } }
                    'C' {
                        $txt = if ($Kind -eq 'Local') { $r.Path } else { $r.Url }
                        if (Set-ClipboardText $txt) { Write-Ok 'Copied to the clipboard.' }
                    }
                }
            }
            continue
        }
        Write-Caution 'Type a number from the list, E, or press ENTER.'
    }
}

# Turns a search term into a wildcard pattern. Plain words become *word*; [ and ] are
# treated literally (they are common in file names).
function New-SearchPattern {
    param([string]$Term)
    $t = $Term.Trim()
    $escaped = $t -replace '([\[\]`])', '`$1'
    if ($t -notmatch '[\*\?]') { $escaped = "*$escaped*" }
    return (New-Object System.Management.Automation.WildcardPattern($escaped, [System.Management.Automation.WildcardOptions]::IgnoreCase))
}

# ======================================================================================
# SECTION 9 - LOCAL ENGINES (copy / move / verify / batch / wipe / search)
# ======================================================================================

# Atomic replace: MoveFileEx(REPLACE_EXISTING | COPY_ALLOWED | WRITE_THROUGH) so the old
# destination file is only replaced once the new one is complete.
function Move-FileReplace {
    param([string]$Source, [string]$Dest)
    if ($script:IsWin -and ('RHFMNative' -as [type])) {
        if ([RHFMNative]::MoveFileExW($Source, $Dest, 0x1 -bor 0x2 -bor 0x8)) { return }
        $code = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw (New-Object System.ComponentModel.Win32Exception($code))
    }
    if ([System.IO.File]::Exists($Dest)) { [System.IO.File]::Delete($Dest) }
    [System.IO.File]::Move($Source, $Dest)
}

# Copies one file. Large files are streamed through a temporary ".rhfmpart" file with
# their own progress bar and can be cancelled part-way (the partial file is removed).
function Copy-LocalFile {
    param([string]$Source, [string]$Dest, [long]$Size, [string]$Name, [int]$BarId = 2)
    if ($Size -lt (Get-LargeFileBytes)) { [System.IO.File]::Copy($Source, $Dest, $true); return }
    $tmp = $Dest + '.rhfmpart'
    $in = $null; $out = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $ui = [System.Diagnostics.Stopwatch]::StartNew()
    [long]$done = 0
    try {
        $in  = [System.IO.File]::Open($Source, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $out = [System.IO.File]::Open($tmp, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $buf = New-Object byte[] (4MB)
        while (($read = $in.Read($buf, 0, $buf.Length)) -gt 0) {
            $out.Write($buf, 0, $read)
            $done += $read
            if ($ui.ElapsedMilliseconds -ge 400) {
                if ((Check-Interrupt -CurrentFile $Name) -eq 'CANCEL') { throw (New-Object System.OperationCanceledException('Cancelled during a large file copy.')) }
                $eta = Get-ETA -Current $done -Total $Size -Elapsed $sw.Elapsed
                Update-Progress -Id $BarId -Activity "Large file: $Name" -Status ("{0} of {1}  {2}" -f (Format-FileSize $done), (Format-FileSize $Size), (Format-AvgSpeed -Bytes $done -Elapsed $sw.Elapsed)) -Percent ($done * 100.0 / [Math]::Max(1, $Size)) -Detail $(if ($eta) { "About $eta left" } else { '' })
                $ui.Restart()
            }
        }
        $out.Flush()
    } catch {
        if ($out) { $out.Dispose(); $out = $null }
        try { if ([System.IO.File]::Exists($tmp)) { [System.IO.File]::Delete($tmp) } } catch { }
        throw
    } finally {
        if ($in)  { $in.Dispose() }
        if ($out) { $out.Dispose() }
        Update-Progress -Id $BarId -Activity 'Large file' -Completed
    }
    Move-FileReplace -Source $tmp -Dest $Dest
}

function Write-LongPathReport {
    param([string]$Source, [string]$Dest, [int]$Length)
    try {
        $line = "[{0}] OVER LIMIT ({1} chars) - Src: {2} | Dst: {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Length, $Source, $Dest
        [System.IO.File]::AppendAllText($LongPathReport, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
    } catch { }
}

# Copies or moves a folder (or a single file) into a destination folder, skipping files
# that are already identical (same size and time), so interrupted jobs can be resumed.
function Perform-LocalTransfer {
    param([string]$Src, [string]$Dst, [switch]$IsMove, [switch]$Nested, [switch]$SkipStorageCheck)
    $mode  = if ($IsMove) { 'MOVE' } else { 'COPY' }
    $stats = New-JobStats
    $sw    = [System.Diagnostics.Stopwatch]::StartNew()
    $srcFull = Resolve-LocalPath $Src
    $dstFull = Resolve-LocalPath $Dst
    $targetName = Split-Path -Path $srcFull -Leaf
    if (-not $targetName) { $targetName = $srcFull }
    $barId = if ($Nested) { 2 } else { 1 }
    Start-JobUi
    try {
        $srcLong = Get-LongPath $srcFull
        $isFileSource = [System.IO.File]::Exists($srcLong)
        $overlap = ''
        if ($isFileSource) {
            $destFile = Join-LocalPath $dstFull (Split-Path -Path $srcFull -Leaf)
            if ((Get-ComparablePath $destFile) -eq (Get-ComparablePath $srcFull)) { $overlap = 'The file is already in that folder (source and destination are the same file).' }
        } else {
            $overlap = Test-PathOverlap -Source $srcFull -Destination $dstFull
        }
        if (-not $overlap -and -not $isFileSource -and (Test-IsLinkFolder (New-Object System.IO.DirectoryInfo($srcLong)))) {
            $overlap = 'It is a junction / symbolic link to another location - links are never followed.'
        }
        if ($overlap) {
            Write-Bad "${targetName}: $overlap Nothing was changed."
            Add-JobFailure -Stats $stats -File $targetName -Path $srcFull -Reason $overlap
            return $stats
        }
        if ($Nested) { Write-Log "LOCAL ${mode}: $srcFull -> $dstFull" 'INFO' -FileOnly }
        else {
            Write-Log "LOCAL ${mode} STARTED" 'INFO'
            Write-Log "  From: $(Format-PathWithStorageTag $srcFull)" 'INFO'
            Write-Log "  To  : $(Format-PathWithStorageTag $dstFull)" 'INFO'
        }

        # 1. Inventory
        if ($isFileSource) {
            $fi = New-Object System.IO.FileInfo($srcLong)
            $files = @($fi); $dirs = @(); $root = $fi.DirectoryName; $totalBytes = $fi.Length
        } else {
            $tree = Get-LocalTree -Path $srcFull -Label 'Scanning source'
            if ($tree.Cancelled) { $stats.Cancelled = $true }
            foreach ($e in $tree.Errors) { $stats.Unreadable++; Add-JobFailure -Stats $stats -File $e.File -Path $e.Path -Reason $e.Reason }
            $files = $tree.Files; $dirs = $tree.Dirs; $root = $tree.Root; $totalBytes = $tree.Bytes
        }
        $stats.TotalBytes = $totalBytes

        # 2. Storage check (skipped for a move within the same drive - nothing extra is written)
        if (-not $SkipStorageCheck -and -not $stats.Cancelled -and $totalBytes -gt 0) {
            $sameDrive = $IsMove -and ([System.IO.Path]::GetPathRoot((Remove-LongPrefix $srcFull)) -ieq [System.IO.Path]::GetPathRoot($dstFull))
            if (-not $sameDrive) {
                $free = Get-PathFreeSpace -Path $dstFull
                $need = $totalBytes
                if ($free -ge 0 -and $need -gt $free) {
                    Write-Hint 'Checking which files are already at the destination...'
                    $need = [long]0
                    foreach ($f in $files) {
                        $d = Get-LongPath (Join-LocalPath $dstFull (Get-ChildRelativePath $root $f.FullName))
                        if ([System.IO.File]::Exists($d)) { $dl = (New-Object System.IO.FileInfo($d)).Length; if ($f.Length -gt $dl) { $need += ($f.Length - $dl) } }
                        else { $need += $f.Length }
                    }
                }
                if ($free -lt 0 -or $need -gt $free) {
                    if (-not (Invoke-PreflightStorageCheck -DestRoot $dstFull -NeededBytes $need)) {
                        Write-Log "Cancelled at the storage check." 'WARN'
                        $stats.Cancelled = $true
                        return $stats
                    }
                } else { Write-Log ("Storage check passed: {0} to write, {1} free." -f (Format-FileSize $need), (Format-FileSize $free)) 'INFO' -Detail }
            }
        }

        # 3. Folders first (keeps empty folders), then files
        $dstLongRoot = Get-LongPath $dstFull
        if (-not [System.IO.Directory]::Exists($dstLongRoot)) { $null = [System.IO.Directory]::CreateDirectory($dstLongRoot) }
        foreach ($d in $dirs) {
            if ($stats.Cancelled) { break }
            $rel = Get-ChildRelativePath $root $d.FullName
            $std = Join-LocalPath $dstFull $rel
            if ($std.Length -ge $CloudPathLimit) { continue }
            try { $null = [System.IO.Directory]::CreateDirectory((Get-LongPath $std)) } catch { }
        }

        $total = $files.Count; $i = 0; [long]$bytesSeen = 0
        $tick = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($f in $files) {
            if ($stats.Cancelled) { break }
            if ((Check-Interrupt -CurrentFile $f.Name -CurrentFolder (Remove-LongPrefix $f.DirectoryName)) -eq 'CANCEL') { $stats.Cancelled = $true; break }
            $i++
            $bytesSeen += $f.Length
            if ($tick.ElapsedMilliseconds -ge 500 -or $i -eq $total) {
                $eta = Get-ETA -Current $bytesSeen -Total $totalBytes -Elapsed $sw.Elapsed
                $act = if ($Nested) { "$mode $targetName" } else { "Local $($mode.ToLower()): $targetName" }
                Update-Progress -Id $barId -Activity $act -Status ("File {0} of {1}  -  {2} of {3}" -f (Format-Count $i), (Format-Count $total), (Format-FileSize $bytesSeen), (Format-FileSize $totalBytes)) -Percent ($i * 100.0 / [Math]::Max(1, $total)) -Detail ($(if ($eta) { "About $eta left  |  " } else { '' }) + 'P = pause, C = cancel')
                $tick.Restart()
            }
            $rel = Get-ChildRelativePath $root $f.FullName
            $destStd = Join-LocalPath $dstFull $rel
            if ($destStd.Length -ge $CloudPathLimit) {
                $stats.TooLong++
                Write-Log ("SKIPPED (path {0} chars, limit {1}): {2}" -f $destStd.Length, $CloudPathLimit, $rel) 'WARN' -Detail
                Write-LongPathReport -Source (Remove-LongPrefix $f.FullName) -Dest $destStd -Length $destStd.Length
                continue
            }
            $destLong = Get-LongPath $destStd
            try {
                $destDir = [System.IO.Path]::GetDirectoryName($destLong)
                if (-not [System.IO.Directory]::Exists($destDir)) { $null = [System.IO.Directory]::CreateDirectory($destDir) }
                $exists = [System.IO.File]::Exists($destLong)
                if ($exists) {
                    $di = New-Object System.IO.FileInfo($destLong)
                    if ($di.Length -eq $f.Length -and (Test-SameFileTime $di.LastWriteTimeUtc $f.LastWriteTimeUtc)) {
                        if ($IsMove) {
                            if (Test-SamePhysicalItem $f.FullName $destLong) {
                                Add-JobFailure -Stats $stats -File $f.Name -Path $rel -Reason 'Source and destination are the same file on disk (reached through two paths) - nothing was deleted.'
                                continue
                            }
                            if (-not (Invoke-ForceDelete -Path $f.FullName)) {
                                Add-JobFailure -Stats $stats -File $f.Name -Path $rel -Reason 'Already at the destination, but the source copy could not be removed (locked or read-only).'
                                continue
                            }
                        }
                        $stats.Skipped++; $stats.SkippedBytes += $f.Length
                        Write-Log "Identical, skipped: $rel" 'INFO' -ConsoleOnly -Detail
                        continue
                    }
                }
                $created = $f.CreationTimeUtc; $written = $f.LastWriteTimeUtc
                if ($IsMove -and -not $exists) {
                    [System.IO.File]::Move($f.FullName, $destLong)
                } else {
                    Copy-LocalFile -Source $f.FullName -Dest $destLong -Size $f.Length -Name $f.Name -BarId ($barId + 1)
                    try { [System.IO.File]::SetCreationTimeUtc($destLong, $created); [System.IO.File]::SetLastWriteTimeUtc($destLong, $written) } catch { }
                    if ($IsMove -and -not (Invoke-ForceDelete -Path $f.FullName)) {
                        Add-JobFailure -Stats $stats -File $f.Name -Path $rel -Reason 'Copied to the destination, but the source could not be removed (locked or read-only).'
                        continue
                    }
                }
                try { [System.IO.File]::SetCreationTimeUtc($destLong, $created); [System.IO.File]::SetLastWriteTimeUtc($destLong, $written) } catch { }
                $stats.Success++; $stats.SuccessBytes += $f.Length; $script:SessionBytes += $f.Length
                if ($stats.Names.Count -lt 200000) { $stats.Names.Add($f.Name) }
                Write-Log ("{0} {1} ({2})" -f $(if ($IsMove) { 'Moved ' } else { 'Copied' }), $rel, (Format-FileSize $f.Length)) 'SUCCESS' -ConsoleOnly -Detail
            } catch [System.OperationCanceledException] {
                $stats.Cancelled = $true
                break
            } catch {
                Write-Log "FAILED: $rel - $($_.Exception.Message)" 'ERROR'
                Add-JobFailure -Stats $stats -File $f.Name -Path $rel -Reason $_.Exception.Message
            }
        }
        Update-Progress -Id $barId -Activity 'Local transfer' -Completed

        if ($IsMove -and -not $stats.Cancelled -and -not $isFileSource) { Remove-EmptyDirectories -TargetFolder $srcFull -IncludeRoot }
    } catch {
        Write-Log "LOCAL $mode ERROR ($targetName): $($_.Exception.Message)" 'ERROR'
        Add-JobFailure -Stats $stats -File $targetName -Path $srcFull -Reason $_.Exception.Message
    } finally {
        $sw.Stop()
        Stop-JobUi
    }

    if ($Nested) {
        Write-Log ("{0} {1}: {2} done, {3} skipped, {4} failed, {5}" -f $mode, $targetName, $stats.Success, $stats.Skipped, $stats.Failed, (Format-FileSize $stats.SuccessBytes)) $(if ($stats.Failed -gt 0) { 'WARN' } else { 'INFO' }) -FileOnly
    } else {
        $st = @((Get-PathStorageType $srcFull), (Get-PathStorageType $dstFull)) | Where-Object { $_ }
        Show-TransferSummary -Title "LOCAL $mode" -Stats $stats -Elapsed $sw.Elapsed -Target $targetName -Storage ($st -join ' -> ') -Color Green -OkLabel $(if ($IsMove) { 'Moved' } else { 'Copied' })
        Register-JobResult -Operation "Local $($mode.ToLower())" -Source $srcFull -Destination $dstFull -Stats $stats -Elapsed $sw.Elapsed
    }
    return $stats
}

function Get-FileSha256 {
    param([string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        return ([System.BitConverter]::ToString($sha.ComputeHash($fs)) -replace '-', '')
    } finally { if ($fs) { $fs.Dispose() }; $sha.Dispose() }
}

# Compares two folders (or a file against a folder): missing, different and extra files.
# -Hash also compares SHA-256 checksums of same-size files (slower, proves the content).
function Perform-LocalVerify {
    param([string]$Src, [string]$Dst, [switch]$Nested, [switch]$Hash)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $srcFull = Resolve-LocalPath $Src; $dstFull = Resolve-LocalPath $Dst
    $r = @{
        Perfect = 0; Mismatch = 0; MissingInDst = 0; MissingInSrc = 0; SrcBytes = [long]0; DstBytes = [long]0
        SrcFiles = 0; DstFiles = 0; SrcDirs = 0; DstDirs = 0; Unreadable = 0; Cancelled = $false
        Differences = (New-Object System.Collections.Generic.List[hashtable])
    }
    Start-JobUi
    try {
        if (-not $Nested) {
            Write-Log "LOCAL VERIFY STARTED$(if ($Hash) { ' (with SHA-256 checksums)' })" 'INFO'
            Write-Log "  Source     : $(Format-PathWithStorageTag $srcFull)" 'INFO'
            Write-Log "  Destination: $(Format-PathWithStorageTag $dstFull)" 'INFO'
        }
        $cmp = [System.StringComparer]::OrdinalIgnoreCase
        $srcMap = New-Object 'System.Collections.Generic.Dictionary[string,System.IO.FileInfo]' ($cmp)
        $dstMap = New-Object 'System.Collections.Generic.Dictionary[string,System.IO.FileInfo]' ($cmp)
        $srcLong = Get-LongPath $srcFull
        if ([System.IO.File]::Exists($srcLong)) {
            $fi = New-Object System.IO.FileInfo($srcLong)
            $srcMap[$fi.Name] = $fi
            $df = Get-LongPath (Join-LocalPath $dstFull $fi.Name)
            if ([System.IO.File]::Exists($df)) { $dstMap[$fi.Name] = (New-Object System.IO.FileInfo($df)) }
        } else {
            $st = Get-LocalTree -Path $srcFull -Label 'Scanning source'
            $r.Cancelled = $st.Cancelled
            $dt = $null
            if (-not $r.Cancelled) { $dt = Get-LocalTree -Path $dstFull -Label 'Scanning destination'; $r.Cancelled = $dt.Cancelled }
            foreach ($f in $st.Files) { $srcMap[(Get-ChildRelativePath $st.Root $f.FullName)] = $f }
            $r.SrcDirs = $st.Dirs.Count; $r.Unreadable += $st.Errors.Count
            foreach ($e in $st.Errors) { $r.Differences.Add(@{ Status = 'UNREADABLE'; Path = $e.Path; Detail = $e.Reason }) }
            if ($dt) {
                foreach ($f in $dt.Files) { $dstMap[(Get-ChildRelativePath $dt.Root $f.FullName)] = $f }
                $r.DstDirs = $dt.Dirs.Count; $r.Unreadable += $dt.Errors.Count
                foreach ($e in $dt.Errors) { $r.Differences.Add(@{ Status = 'UNREADABLE'; Path = $e.Path; Detail = $e.Reason }) }
            }
        }
        $r.SrcFiles = $srcMap.Count; $r.DstFiles = $dstMap.Count
        if (-not $r.Cancelled) {
            $i = 0; $tick = [System.Diagnostics.Stopwatch]::StartNew()
            foreach ($key in $srcMap.Keys) {
                $i++
                if ((Check-Interrupt -CurrentFile $key) -eq 'CANCEL') { $r.Cancelled = $true; break }
                if ($tick.ElapsedMilliseconds -ge 500) {
                    Update-Progress -Id $(if ($Nested) { 2 } else { 1 }) -Activity 'Verifying' -Status ("File {0} of {1}" -f (Format-Count $i), (Format-Count $srcMap.Count)) -Percent ($i * 100.0 / [Math]::Max(1, $srcMap.Count)) -Detail $key
                    $tick.Restart()
                }
                $s = $srcMap[$key]; $r.SrcBytes += $s.Length
                if (-not $dstMap.ContainsKey($key)) {
                    $r.MissingInDst++; $r.Differences.Add(@{ Status = 'MISSING'; Path = $key; Detail = (Format-FileSize $s.Length) })
                    continue
                }
                $d = $dstMap[$key]; $r.DstBytes += $d.Length
                $why = ''
                if ($s.Length -ne $d.Length) { $why = "size {0} vs {1}" -f $s.Length, $d.Length }
                elseif (-not (Test-SameFileTime $s.LastWriteTimeUtc $d.LastWriteTimeUtc)) {
                    $why = "modified {0} vs {1}" -f $s.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'), $d.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                }
                if (-not $why -and $Hash) {
                    try { if ((Get-FileSha256 $s.FullName) -ne (Get-FileSha256 $d.FullName)) { $why = 'content differs (SHA-256)' } }
                    catch { $why = "could not read for checksum: $($_.Exception.Message)" }
                }
                if ($why) { $r.Mismatch++; $r.Differences.Add(@{ Status = 'DIFFERENT'; Path = $key; Detail = $why }) }
                else { $r.Perfect++ }
            }
            if (-not $r.Cancelled) {
                foreach ($key in $dstMap.Keys) {
                    if (-not $srcMap.ContainsKey($key)) {
                        $r.MissingInSrc++; $r.DstBytes += $dstMap[$key].Length
                        $r.Differences.Add(@{ Status = 'EXTRA'; Path = $key; Detail = (Format-FileSize $dstMap[$key].Length) })
                    }
                }
            }
        }
        if ($r.Differences.Count -gt 0) {
            Write-LogBlock -Lines @($r.Differences | ForEach-Object { "VERIFY {0}: {1} ({2})" -f $_.Status, $_.Path, $_.Detail }) -Level 'WARN'
        }
    } catch {
        Write-Log "VERIFY ERROR: $($_.Exception.Message)" 'ERROR'
        $r.Unreadable++
        $r.Differences.Add(@{ Status = 'UNREADABLE'; Path = $srcFull; Detail = $_.Exception.Message })
    } finally {
        $sw.Stop()
        Stop-JobUi
    }
    if (-not $Nested) { Show-VerifyResult -Title 'LOCAL VERIFY' -Result $r -Elapsed $sw.Elapsed -Note $(if ($Hash) { 'Compared size, date and SHA-256 checksum.' } else { 'Compared size and date (within 2 seconds).' }) }
    return $r
}

# Summary + first differences + optional CSV export for any verify result.
function Show-VerifyResult {
    param([string]$Title, [hashtable]$Result, [TimeSpan]$Elapsed, [string]$Note = '', [switch]$NoExport)
    $r = $Result
    $diffs = @($r.Differences)
    if ($diffs.Count -gt 0) {
        $show = New-Object System.Collections.Generic.List[object]
        foreach ($d in ($diffs | Select-Object -First 15)) {
            $c = switch ($d.Status) { 'MISSING' { 'Red' } 'DIFFERENT' { 'Red' } 'UNREADABLE' { 'Red' } default { 'Yellow' } }
            $show.Add(@{ Text = ("{0,-10} {1}" -f $d.Status, $d.Path); Color = $c })
            if ($d.Detail) { $show.Add(@{ Text = "           $($d.Detail)"; Color = 'DarkGray' }) }
        }
        if ($diffs.Count -gt 15) { $show.Add('-'); $show.Add(@{ Text = "...and $(Format-Count ($diffs.Count - 15)) more (all of them are in the log)."; Color = 'Yellow' }) }
        Write-Panel -Title ("DIFFERENCES ({0})" -f (Format-Count $diffs.Count)) -Lines $show.ToArray() -Color Yellow
    }
    $clean = ($r.Mismatch -eq 0 -and $r.MissingInDst -eq 0 -and $r.Unreadable -eq 0 -and -not $r.Cancelled -and -not $r.Incomplete)
    $rows = @(
        @{ Label = 'Source';      Value = ("{0} files  |  {1}" -f (Format-Count $r.SrcFiles), (Format-FileSize $r.SrcBytes)) }
        @{ Label = 'Destination'; Value = ("{0} files  |  {1}" -f (Format-Count $r.DstFiles), (Format-FileSize $r.DstBytes)) }
        @{ Label = 'Identical';   Value = (Format-Count $r.Perfect); Level = 'SUCCESS' }
        @{ Label = 'Different';   Value = (Format-Count $r.Mismatch); Level = $(if ($r.Mismatch) { 'ERROR' } else { 'SUCCESS' }) }
        @{ Label = 'Missing in destination'; Value = (Format-Count $r.MissingInDst); Level = $(if ($r.MissingInDst) { 'ERROR' } else { 'SUCCESS' }) }
        @{ Label = 'Extra in destination';   Value = (Format-Count $r.MissingInSrc); Level = $(if ($r.MissingInSrc) { 'WARN' } else { 'SUCCESS' }) }
    )
    if ($r.Unreadable) { $rows += @{ Label = 'Unreadable / unlisted folders'; Value = (Format-Count $r.Unreadable); Level = 'ERROR' } }
    if ($Note) { $rows += @{ Label = 'Method'; Value = $Note } }
    $rows += @{ Label = 'Time taken'; Value = (Format-Elapsed $Elapsed) }
    $state = if ($r.Cancelled) { 'CANCELLED' } elseif ($clean) { 'MATCH' } else { 'DIFFERENCES FOUND' }
    $col = if ($r.Cancelled) { 'Yellow' } elseif ($clean) { 'Green' } else { 'Red' }
    Write-SummaryBox -Title "$Title - $state" -Lines $rows -BoxColor $col
    if ($clean -and $r.MissingInSrc -eq 0) { Write-Host ("  " + (Get-BarText -Fraction 1 -Width 12) + "  Perfect match - the two locations are identical.") -ForegroundColor Green }
    if (-not $NoExport -and $diffs.Count -gt 0) {
        if (Read-YesNo -Question 'Export the list of differences to CSV?' -Default $false) {
            $csv = Join-Path $LogFileDirectory ("VerifyDifferences_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            try {
                $diffs | ForEach-Object { [pscustomobject]@{ Status = $_.Status; Path = $_.Path; Detail = $_.Detail } } | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
                Write-Ok "Saved: $csv"
            } catch { Write-Bad "Export failed: $($_.Exception.Message)" }
        }
    }
}

# Deletes a local file or folder tree. Junctions / symbolic links are removed as links only
# (their targets are never touched). Updates -Stats.
function Remove-LocalItem {
    param([string]$Path, [hashtable]$Stats)
    $long = Get-LongPath (Resolve-LocalPath $Path)
    $name = Split-Path -Path $Path -Leaf
    if (Test-ProtectedFolder $Path) {
        Write-Bad "'$Path' is a drive / share root or a system / profile folder - it is never deleted."
        Add-JobFailure -Stats $Stats -File $name -Path $Path -Reason 'Protected folder - refused.'
        return
    }
    if ([System.IO.File]::Exists($long)) {
        $len = (New-Object System.IO.FileInfo($long)).Length
        if (Invoke-ForceDelete -Path $long) { $Stats.Success++; $Stats.SuccessBytes += $len; Write-Log "Deleted file: $name" 'SUCCESS' -ConsoleOnly -Detail }
        else { Add-JobFailure -Stats $Stats -File $name -Path $Path -Reason 'File is locked, read-only or access is denied.' }
        return
    }
    if (-not [System.IO.Directory]::Exists($long)) { Add-JobFailure -Stats $Stats -File $name -Path $Path -Reason 'Not found.'; return }
    $di = New-Object System.IO.DirectoryInfo($long)
    if (Test-IsLinkFolder $di) {
        if (Invoke-ForceDelete -Path $long -IsFolder) { $Stats.Success++; Write-Log "Removed link (target untouched): $name" 'SUCCESS' }
        else { Add-JobFailure -Stats $Stats -File $name -Path $Path -Reason 'Link could not be removed.' }
        return
    }
    $tree = Get-LocalTree -Path $Path -Label "Scanning $name"
    if ($tree.Cancelled) { $Stats.Cancelled = $true; return }
    foreach ($e in $tree.Errors) { $Stats.Unreadable++; Add-JobFailure -Stats $Stats -File $e.File -Path $e.Path -Reason $e.Reason }
    $i = 0; $tick = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($f in $tree.Files) {
        if ((Check-Interrupt -CurrentFile $f.Name) -eq 'CANCEL') { $Stats.Cancelled = $true; return }
        $i++
        if ($tick.ElapsedMilliseconds -ge 500) { Update-Progress -Id 2 -Activity "Deleting $name" -Status ("File {0} of {1}" -f $i, $tree.Files.Count) -Percent ($i * 100.0 / [Math]::Max(1, $tree.Files.Count)); $tick.Restart() }
        $len = $f.Length
        if (Invoke-ForceDelete -Path $f.FullName) { $Stats.Success++; $Stats.SuccessBytes += $len; Write-Log "Deleted: $($f.Name)" 'SUCCESS' -ConsoleOnly -Detail }
        else { Add-JobFailure -Stats $Stats -File $f.Name -Path (Remove-LongPrefix $f.FullName) -Reason 'Locked, read-only or access denied.' }
    }
    foreach ($l in $tree.Links) {
        if (-not (Invoke-ForceDelete -Path $l.FullName -IsFolder)) { Add-JobFailure -Stats $Stats -File $l.Name -Path (Remove-LongPrefix $l.FullName) -Reason 'Link could not be removed.' }
    }
    foreach ($d in @($tree.Dirs | Sort-Object { $_.FullName.Length } -Descending)) {
        if (-not (Invoke-ForceDelete -Path $d.FullName -IsFolder)) { Add-JobFailure -Stats $Stats -File $d.Name -Path (Remove-LongPrefix $d.FullName) -Reason 'Folder not empty or locked.' }
    }
    if (-not (Invoke-ForceDelete -Path $tree.Root -IsFolder)) { Add-JobFailure -Stats $Stats -File $name -Path $Path -Reason 'Folder could not be removed (something inside is still open or locked).' }
    Update-Progress -Id 2 -Activity 'Deleting' -Completed
}

function Get-LocalIndexRows {
    param([object[]]$Items, [switch]$NoSizes)
    $rows = New-Object System.Collections.Generic.List[hashtable]
    $n = 0
    $script:JobCancelled = $false
    foreach ($it in $Items) {
        $n++
        if ($it -is [System.IO.DirectoryInfo] -and (Test-IsLinkFolder $it)) {
            $rows.Add(@{ Kind = 'FOLDER'; Name = "$($it.Name)  [link]"; Info = 'link'; Date = $it.LastWriteTime.ToString('yyyy-MM-dd'); Bytes = -1; Item = $it; IsLink = $true })
            continue
        }
        if ($it -is [System.IO.DirectoryInfo]) {
            if ($NoSizes) { $rows.Add(@{ Kind = 'FOLDER'; Name = $it.Name; Info = ''; Date = $it.LastWriteTime.ToString('yyyy-MM-dd'); Bytes = -1; Item = $it }); continue }
            Write-StatusLine -Text ("  Measuring {0}/{1}: {2}" -f $n, $Items.Count, $it.Name)
            $t = Get-LocalTree -Path $it.FullName -Quiet
            if ($t.Cancelled) { $script:JobCancelled = $false }
            $info = "{0} | {1:N0}" -f (Format-FileSize $t.Bytes), ($t.Files.Count)
            $rows.Add(@{ Kind = 'FOLDER'; Name = $it.Name; Info = $info; Date = $it.LastWriteTime.ToString('yyyy-MM-dd'); Bytes = $t.Bytes; Item = $it })
        } else {
            $rows.Add(@{ Kind = 'FILE'; Name = $it.Name; Info = (Format-FileSize $it.Length); Date = $it.LastWriteTime.ToString('yyyy-MM-dd'); Bytes = $it.Length; Item = $it })
        }
    }
    Clear-StatusLine
    return ,$rows.ToArray()
}

# Batch screen for a local folder: pick items, then Copy / Move / Delete / Verify.
function Perform-LocalBatch {
    param([string]$Src, [string]$Dst)
    Write-Log "LOCAL BATCH: $(Format-PathWithStorageTag $Src)$(if ($Dst) { " -> $(Format-PathWithStorageTag $Dst)" })" 'INFO'
    while ($true) {
        try {
            $srcDi = New-Object System.IO.DirectoryInfo((Get-LongPath (Resolve-LocalPath $Src)))
            $items = @($srcDi.GetFileSystemInfos() |
                Sort-Object @{ Expression = { -not ($_ -is [System.IO.DirectoryInfo]) } }, @{ Expression = { $_.Name } })
        } catch { Write-Bad "Could not read '$Src': $($_.Exception.Message)"; return }
        if ($items.Count -eq 0) { Write-Caution 'This folder is empty.'; return }
        Write-Panel -Title 'LOCAL BATCH' -Color Green -Lines @(
            @{ Label = 'Source';      Value = $Src }
            @{ Label = 'Destination'; Value = $(if ($Dst) { $Dst } else { '(not set - you will be asked if needed)' }) }
        )
        Write-Hint 'Measuring folder sizes (sizes include sub-folders; the second number is the file count)...'
        $rows = Get-LocalIndexRows -Items $items
        Show-ItemIndex -Title "CONTENTS OF $(Split-Path -Path $Src -Leaf)" -Rows $rows -Color Green -InfoHeader 'SIZE | FILES'
        $sel = Read-IndexSelection -Count $rows.Count
        if ($sel.Action -eq 'Quit') { return }
        if ($sel.Action -eq 'Refresh') { continue }
        if ($sel.Action -eq 'Search') { Invoke-LocalSearchLoop -SearchPath $Src; continue }

        $act = Read-Choice -Prompt 'C = copy   M = move   D = delete   V = verify   H = verify with checksums' -Valid @('C', 'M', 'D', 'V', 'H')
        if (-not $act) { continue }
        $chosen = @($sel.Indices | ForEach-Object { $rows[$_] })
        $verb = switch ($act) { 'C' { 'COPY' } 'M' { 'MOVE' } 'D' { 'DELETE' } 'V' { 'VERIFY' } 'H' { 'VERIFY (CHECKSUMS)' } }

        if ($act -in @('D', 'M')) {
            $prot = @($chosen | Where-Object { Test-ProtectedFolder $_.Item.FullName })
            if ($prot.Count -gt 0) {
                foreach ($x in $prot) { Write-Bad "'$($x.Name)' is a system, profile or root folder - it cannot be deleted or moved here." }
                $chosen = @($chosen | Where-Object { $prot -notcontains $_ })
                if ($chosen.Count -eq 0) { continue }
            }
        }
        $links = @($chosen | Where-Object { $_.IsLink })
        if ($links.Count -gt 0 -and $act -ne 'D') {
            foreach ($x in $links) { Write-Caution "'$($x.Item.Name)' is a junction / link to another location - skipped (links are never followed)." }
            $chosen = @($chosen | Where-Object { -not $_.IsLink })
            if ($chosen.Count -eq 0) { continue }
        }
        if ($act -eq 'D') {
            $bytes = [long]0; foreach ($c in $chosen) { if ($c.Bytes -gt 0) { $bytes += $c.Bytes } }
            if (-not (Read-Confirmation -Message ("PERMANENTLY delete {0} item(s) ({1})? This bypasses the Recycle Bin and cannot be undone." -f $chosen.Count, (Format-FileSize $bytes)) -Color Red)) { Write-Hint 'Nothing deleted.'; continue }
        } else {
            if (-not $Dst) {
                $Dst = Read-LocalFolder -Prompt 'Destination folder for these items:' -LogLabel 'Batch destination' -CreateIfMissing
                if (-not $Dst) { $Dst = $null; continue }
            }
            $blocked = $false
            foreach ($c in $chosen) {
                $ov = Test-PathOverlap -Source $c.Item.FullName -Destination $(if ($c.Kind -eq 'FOLDER') { Join-LocalPath $Dst $c.Name } else { $Dst })
                if ($ov -and $c.Kind -eq 'FOLDER') { Write-Bad "'$($c.Name)': $ov"; $blocked = $true }
            }
            if ($blocked) { Write-Caution 'Choose a destination outside the selected folders.'; continue }
            if ($act -in @('C', 'M')) {
                $bytes = [long]0; foreach ($c in $chosen) { if ($c.Bytes -gt 0) { $bytes += $c.Bytes } }
                $free = Get-PathFreeSpace -Path $Dst
                if (($free -ge 0 -and $bytes -gt $free) -or $free -lt 0) {
                    if (-not (Invoke-PreflightStorageCheck -DestRoot $Dst -NeededBytes $bytes)) { continue }
                }
            }
        }

        $batch = New-JobStats
        $vr = @{ Perfect = 0; Mismatch = 0; MissingInDst = 0; MissingInSrc = 0; SrcBytes = [long]0; DstBytes = [long]0; SrcFiles = 0; DstFiles = 0; Unreadable = 0; Cancelled = $false; Differences = (New-Object System.Collections.Generic.List[hashtable]) }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Start-JobUi
        try {
            $n = 0
            foreach ($c in $chosen) {
                if ($script:JobCancelled) { $batch.Cancelled = $true; break }
                $n++
                Update-Progress -Id 1 -Activity "Local batch $($verb.ToLower())" -Status ("Item {0} of {1}: {2}" -f $n, $chosen.Count, $c.Name) -Percent (($n - 1) * 100.0 / $chosen.Count) -Detail 'P = pause, C = cancel'
                Write-BatchItemHeader -Index $n -Total $chosen.Count -Name $c.Name -Verb $verb -Color Green
                $full = $c.Item.FullName
                switch ($act) {
                    'D' { Remove-LocalItem -Path $full -Stats $batch }
                    { $_ -in @('C', 'M') } {
                        $target = if ($c.Kind -eq 'FOLDER') { Join-LocalPath $Dst $c.Name } else { $Dst }
                        $res = Perform-LocalTransfer -Src $full -Dst $target -IsMove:($act -eq 'M') -Nested -SkipStorageCheck
                        Merge-JobStats -Into $batch -From $res
                        Write-Host ("  {0} done, {1} skipped, {2} failed" -f $res.Success, $res.Skipped, $res.Failed) -ForegroundColor $(if ($res.Failed) { 'Yellow' } else { 'DarkGray' })
                    }
                    { $_ -in @('V', 'H') } {
                        $target = if ($c.Kind -eq 'FOLDER') { Join-LocalPath $Dst $c.Name } else { $Dst }
                        if ($c.Kind -eq 'FOLDER' -and -not [System.IO.Directory]::Exists((Get-LongPath $target))) {
                            Write-Bad "Missing in destination: $($c.Name)"
                            $vr.MissingInDst++; $vr.Differences.Add(@{ Status = 'MISSING'; Path = $c.Name; Detail = 'whole folder missing' })
                        } else {
                            $one = Perform-LocalVerify -Src $full -Dst $target -Nested -Hash:($act -eq 'H')
                            foreach ($k in 'Perfect', 'Mismatch', 'MissingInDst', 'MissingInSrc', 'SrcFiles', 'DstFiles', 'Unreadable') { $vr[$k] += $one[$k] }
                            $vr.SrcBytes += $one.SrcBytes; $vr.DstBytes += $one.DstBytes
                            foreach ($d in $one.Differences) { $vr.Differences.Add(@{ Status = $d.Status; Path = "$($c.Name)\$($d.Path)"; Detail = $d.Detail }) }
                            if ($one.Cancelled) { $vr.Cancelled = $true }
                            Write-Host ("  {0} identical, {1} different, {2} missing, {3} extra" -f $one.Perfect, $one.Mismatch, $one.MissingInDst, $one.MissingInSrc) -ForegroundColor $(if ($one.Mismatch -or $one.MissingInDst) { 'Yellow' } else { 'DarkGray' })
                        }
                    }
                }
                if ($script:JobCancelled) { $batch.Cancelled = $true; $vr.Cancelled = $true }
            }
        } finally { $sw.Stop(); Stop-JobUi }

        if ($act -in @('V', 'H')) {
            Show-VerifyResult -Title 'LOCAL BATCH VERIFY' -Result $vr -Elapsed $sw.Elapsed -Note $(if ($act -eq 'H') { 'Size, date and SHA-256 checksum.' } else { 'Size and date (within 2 seconds).' })
            $vs = New-JobStats; $vs.Success = $vr.Perfect; $vs.Failed = $vr.Mismatch + $vr.MissingInDst; $vs.Cancelled = $vr.Cancelled
            Register-JobResult -Operation 'Local batch verify' -Source $Src -Destination $Dst -Stats $vs -Elapsed $sw.Elapsed
        } elseif ($act -eq 'D') {
            Show-TransferSummary -Title 'LOCAL BATCH DELETE' -Stats $batch -Elapsed $sw.Elapsed -Color Green -OkLabel 'Deleted'
            Register-JobResult -Operation 'Local batch delete' -Source $Src -Stats $batch -Elapsed $sw.Elapsed
        } else {
            Show-TransferSummary -Title "LOCAL BATCH $verb" -Stats $batch -Elapsed $sw.Elapsed -Color Green -OkLabel $(if ($act -eq 'M') { 'Moved' } else { 'Copied' }) -Storage ((@((Get-PathStorageType $Src), (Get-PathStorageType $Dst)) | Where-Object { $_ }) -join ' -> ')
            Register-JobResult -Operation "Local batch $($verb.ToLower())" -Source $Src -Destination $Dst -Stats $batch -Elapsed $sw.Elapsed
        }
        Show-Pause -Message 'Press any key to return to the folder list...'
    }
}

function Perform-LocalWipe {
    param([string]$TargetFolder)
    $full = Resolve-LocalPath $TargetFolder
    if (Test-ProtectedFolder $full) {
        Write-Bad "'$full' is a drive root, share root or system / profile folder. It cannot be wiped with this tool."
        return
    }
    Write-Hint 'Measuring what will be deleted...'
    $m = Get-LocalTree -Path $full -Label 'Scanning'
    if ($m.Cancelled) { Write-Caution 'Cancelled.'; return }
    Write-Panel -Title 'ABOUT TO DELETE PERMANENTLY' -Color Red -Lines @(
        @{ Label = 'Folder';  Value = $full; Level = 'ERROR' }
        @{ Label = 'Files';   Value = (Format-Count $m.Files.Count) }
        @{ Label = 'Folders'; Value = (Format-Count $m.Dirs.Count) }
        @{ Label = 'Size';    Value = (Format-FileSize $m.Bytes) }
        'This bypasses the Windows Recycle Bin. Junctions / links inside are removed as links only.'
    )
    if (-not (Read-Confirmation -Message 'Delete this folder and everything in it?' -Color Red)) { Write-Hint 'Nothing deleted.'; Write-Log 'Local wipe cancelled at confirmation.' 'INFO' -FileOnly; return }
    Write-Log "LOCAL WIPE STARTED: $full" 'WARN'
    $stats = New-JobStats
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Start-JobUi
    try { Remove-LocalItem -Path $full -Stats $stats } finally { $sw.Stop(); Stop-JobUi }
    Show-TransferSummary -Title 'LOCAL WIPE' -Stats $stats -Elapsed $sw.Elapsed -Target $full -Color Green -OkLabel 'Deleted'
    Register-JobResult -Operation 'Local wipe' -Source $full -Stats $stats -Elapsed $sw.Elapsed
}

function Perform-LocalSearch {
    param([string]$SearchPath, [string]$SearchTerm)
    $root = Get-LongPath (Resolve-LocalPath $SearchPath)
    if (-not [System.IO.Directory]::Exists($root)) { Write-Bad "Folder not found: $SearchPath"; return }
    $wp = New-SearchPattern $SearchTerm
    Write-Log "LOCAL SEARCH for '$SearchTerm' in $SearchPath" 'INFO' -FileOnly
    $results = New-Object System.Collections.Generic.List[hashtable]
    $folders = 0; $files = 0; $errors = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $tick = [System.Diagnostics.Stopwatch]::StartNew()
    $queue = New-Object System.Collections.Generic.Queue[System.IO.DirectoryInfo]
    $queue.Enqueue((New-Object System.IO.DirectoryInfo($root)))
    $cancelled = $false
    Start-JobUi
    try {
        while ($queue.Count -gt 0) {
            $dir = $queue.Dequeue()
            if ((Check-Interrupt -CurrentFolder $dir.FullName) -eq 'CANCEL') { $cancelled = $true; break }
            $folders++
            try {
                foreach ($f in $dir.EnumerateFiles()) {
                    $files++
                    if ($wp.IsMatch($f.Name)) {
                        $results.Add(@{ Type = 'FILE'; Name = $f.Name; Path = (Remove-LongPrefix $f.FullName); Parent = (Remove-LongPrefix $dir.FullName); Size = $f.Length; Modified = $f.LastWriteTime })
                        Write-Log "Match: $(Remove-LongPrefix $f.FullName)" 'SUCCESS' -ConsoleOnly -Detail
                    }
                }
                foreach ($d in $dir.EnumerateDirectories()) {
                    if ($wp.IsMatch($d.Name)) {
                        $results.Add(@{ Type = 'FOLDER'; Name = $d.Name; Path = (Remove-LongPrefix $d.FullName); Parent = (Remove-LongPrefix $dir.FullName); Size = 0; Modified = $d.LastWriteTime })
                        Write-Log "Match: $(Remove-LongPrefix $d.FullName)\" 'SUCCESS' -ConsoleOnly -Detail
                    }
                    if (-not (Test-IsLinkFolder $d)) { $queue.Enqueue($d) }
                }
            } catch { $errors++ }
            if ($tick.ElapsedMilliseconds -ge 500) {
                Update-Progress -Id 1 -Activity "Searching for '$SearchTerm'" -Status ("{0} matches  |  {1} folders, {2} files scanned" -f (Format-Count $results.Count), (Format-Count $folders), (Format-Count $files)) -Detail (Remove-LongPrefix $dir.FullName)
                $tick.Restart()
            }
        }
    } finally { $sw.Stop(); Stop-JobUi }
    $state = if ($cancelled) { 'stopped early' } else { 'complete' }
    Write-Log ("Search {0}: {1} matches, {2} folders and {3} files scanned in {4}{5}" -f $state, $results.Count, $folders, $files, (Format-Elapsed $sw.Elapsed), $(if ($errors) { ", $errors folders could not be read" } else { '' })) $(if ($cancelled) { 'WARN' } else { 'SUCCESS' })
    if ($results.Count -gt 0) { Write-LogBlock -Lines @($results | ForEach-Object { "MATCH [$($_.Type)] $($_.Path)" }) }
    Show-SearchResults -Results $results.ToArray() -Term $SearchTerm -Kind Local
}

# Repeated searches in one folder (used by option 12 and the [S] key in batch screens).
function Invoke-LocalSearchLoop {
    param([string]$SearchPath)
    while ($true) {
        Write-Host ""
        Write-Hint "Search in: $SearchPath"
        Write-Hint 'Examples:  invoice   *.pdf   *2024*report*   Budget [Final].xlsx'
        $term = (Read-Line -Prompt '  Search for (ENTER = back) > ' -LogLabel 'Search term').Trim()
        if (-not $term -or $term -match '^(?i)q$') { return }
        Perform-LocalSearch -SearchPath $SearchPath -SearchTerm $term
    }
}

# ======================================================================================
# SECTION 10 - CLOUD ENGINES (copy / move / batch / verify / wipe / search)
# ======================================================================================

# Library view link for a cloud folder (opens reliably in the browser).
function Get-CloudFolderViewUrl {
    param([string]$SiteUrl, [string]$ServerRelativeFolder)
    $rel = Get-SiteRelativePath -SiteUrl $SiteUrl -ServerRelativePath $ServerRelativeFolder
    $lib = ($rel -split '/')[0]
    if (-not $lib) { return Get-CloudWebUrl -SiteUrl $SiteUrl -ServerRelativePath $ServerRelativeFolder }
    $authority = ([System.Uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $libPath = (Get-SitePath $SiteUrl) + '/' + $lib
    $encLib = (($libPath -split '/') | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/'
    return "$authority$encLib/Forms/AllItems.aspx?id=$([System.Uri]::EscapeDataString($ServerRelativeFolder))"
}

function Get-SiteRecycleBinUrl {
    param([string]$SiteUrl)
    return ($SiteUrl.TrimEnd('/') + '/_layouts/15/RecycleBin.aspx')
}

function Test-SameTenant {
    param([string]$UrlA, [string]$UrlB)
    try {
        $a = ([System.Uri]$UrlA).Host.ToLowerInvariant() -replace '-my\.sharepoint\.com$', '.sharepoint.com'
        $b = ([System.Uri]$UrlB).Host.ToLowerInvariant() -replace '-my\.sharepoint\.com$', '.sharepoint.com'
        return ($a -eq $b)
    } catch { return $false }
}

# Compares two inventory maps (lower-cased relative path -> @{ Length; Path }).
# Same result shape as local verify.
function Compare-CloudInventories {
    param([hashtable]$SrcMap, [hashtable]$DstMap)
    $r = @{
        Perfect = 0; Mismatch = 0; MissingInDst = 0; MissingInSrc = 0; SrcBytes = [long]0; DstBytes = [long]0
        SrcFiles = $SrcMap.Count; DstFiles = $DstMap.Count; Unreadable = 0; Cancelled = $false; Incomplete = $false
        Differences = (New-Object System.Collections.Generic.List[hashtable])
    }
    foreach ($k in $SrcMap.Keys) {
        $s = $SrcMap[$k]; $r.SrcBytes += [long]$s.Length
        if (-not $DstMap.ContainsKey($k)) { $r.MissingInDst++; $r.Differences.Add(@{ Status = 'MISSING'; Path = $s.Path; Detail = (Format-FileSize ([long]$s.Length)) }); continue }
        $d = $DstMap[$k]; $r.DstBytes += [long]$d.Length
        if ([long]$s.Length -ne [long]$d.Length) { $r.Mismatch++; $r.Differences.Add(@{ Status = 'DIFFERENT'; Path = $s.Path; Detail = ("size {0} vs {1} bytes" -f $s.Length, $d.Length) }) }
        else { $r.Perfect++ }
    }
    foreach ($k in $DstMap.Keys) {
        if (-not $SrcMap.ContainsKey($k)) {
            $d = $DstMap[$k]; $r.MissingInSrc++; $r.DstBytes += [long]$d.Length
            $r.Differences.Add(@{ Status = 'EXTRA'; Path = $d.Path; Detail = (Format-FileSize ([long]$d.Length)) })
        }
    }
    return $r
}

function ConvertTo-InventoryMap {
    param($Inventory, [string]$Prefix = '')
    $m = @{}
    if (-not $Inventory) { return $m }
    foreach ($f in $Inventory.Files) {
        $p = if ($Prefix) { "$Prefix/$($f.RelPath)" } else { $f.RelPath }
        $m[$p.ToLowerInvariant()] = @{ Length = [long]$f.Length; Path = $p }
    }
    return $m
}

# Option 1 / 2: server-side copy or move of a folder (or file) INTO a destination folder.
function Perform-CloudTransfer {
    param([hashtable]$SrcData, [hashtable]$DstData, [switch]$IsMove)
    $mode = if ($IsMove) { 'MOVE' } else { 'COPY' }
    if (-not (Test-SameTenant $SrcData.SiteUrl $DstData.SiteUrl)) {
        Write-Bad 'The two addresses belong to different Microsoft 365 tenants - SharePoint cannot copy between tenants.'
        Write-Hint 'Use [13] Download from the first tenant, then [15] Upload to the second.'
        return
    }
    if ($SrcData.SiteUrl -ine $DstData.SiteUrl) {
        Write-Caution 'Source and destination are on different sites. SharePoint will run a cross-site copy job, which can take a while.'
    }
    $leaf = ($SrcData.FolderPath -split '/')[-1]
    $stats = New-JobStats
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $errText = ''
    Start-JobUi
    try {
        Connect-CloudSite -SiteUrl $SrcData.SiteUrl
        if ($SrcData.SiteUrl -ieq $DstData.SiteUrl) {
            $dRel = Get-SiteRelativePath -SiteUrl $DstData.SiteUrl -ServerRelativePath $DstData.FolderPath
            if ($dRel -and -not (Ensure-CloudFolder -SiteRelativePath $dRel)) { Write-Caution "The destination folder could not be checked or created; trying anyway." }
        }
        Write-Log "CLOUD $mode STARTED: $($SrcData.FolderPath) -> $($DstData.FolderPath)" 'INFO'
        Write-Info "Microsoft's servers are doing the $($mode.ToLower()) - nothing passes through this PC. Please keep this window open."
        Update-Progress -Id 1 -Activity "Cloud $($mode.ToLower()): $leaf" -Status 'SharePoint is working on it...'
        Invoke-CloudCopyMove -SourceUrl $SrcData.FolderPath -TargetUrl $DstData.FolderPath -Move:$IsMove -Name $leaf
        $stats.Success = 1
        Write-Log "CLOUD $mode finished: '$leaf' is now in '$($DstData.FolderPath)'" 'SUCCESS'
    } catch {
        $errText = $_.Exception.Message
        if ($_.Exception -is [System.OperationCanceledException]) { $stats.Cancelled = $true }
        Add-JobFailure -Stats $stats -File $leaf -Path $SrcData.FolderPath -Reason $errText
        Write-Log "CLOUD $mode FAILED: $errText" 'ERROR'
    } finally { $sw.Stop(); Stop-JobUi }

    $rows = @(
        @{ Label = 'Source';      Value = $SrcData.FolderPath }
        @{ Label = 'Copied into'; Value = $DstData.FolderPath }
        @{ Label = 'Result';      Value = $(if ($stats.Success) { "Done - now at $($DstData.FolderPath)/$leaf" } else { "Failed: $errText" }); Level = $(if ($stats.Success) { 'SUCCESS' } else { 'ERROR' }) }
        @{ Label = 'Time taken';  Value = (Format-Elapsed $sw.Elapsed) }
    )
    if ($IsMove) { $rows[1].Label = 'Moved into' }
    Write-SummaryBox -Title "CLOUD $mode - $(if ($stats.Success) { 'COMPLETE' } else { 'FAILED' })" -Lines $rows -BoxColor $(if ($stats.Success) { 'Green' } else { 'Red' })
    Register-JobResult -Operation "Cloud $($mode.ToLower())" -Source $SrcData.FolderPath -Destination $DstData.FolderPath -Stats $stats -Elapsed $sw.Elapsed
    if ($stats.Success -and -not $IsMove) {
        if (Read-YesNo -Question 'Verify the copy now (compares every file name and size)?' -Default $false) {
            Perform-CloudVerify -SrcData $SrcData -DstData @{ SiteUrl = $DstData.SiteUrl; FolderPath = (Join-CloudPath $DstData.FolderPath $leaf) }
        }
    }
}

# Option 5: compare two cloud folders without downloading anything.
function Perform-CloudVerify {
    param([hashtable]$SrcData, [hashtable]$DstData)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $r = $null
    Start-JobUi
    try {
        Write-Log "CLOUD VERIFY: $($SrcData.FolderPath) <-> $($DstData.FolderPath)" 'INFO'
        Connect-CloudSite -SiteUrl $SrcData.SiteUrl
        $si = Get-CloudInventory -SiteUrl $SrcData.SiteUrl -FolderPath $SrcData.FolderPath -Tag 'Scanning source'
        $di = $null
        if (-not $si.Cancelled) {
            Connect-CloudSite -SiteUrl $DstData.SiteUrl
            $di = Get-CloudInventory -SiteUrl $DstData.SiteUrl -FolderPath $DstData.FolderPath -Tag 'Scanning destination'
        }
        $r = Compare-CloudInventories -SrcMap (ConvertTo-InventoryMap $si) -DstMap (ConvertTo-InventoryMap $di)
        $r.Cancelled = ($si.Cancelled -or ($di -and $di.Cancelled))
        foreach ($inv in @($si, $di)) {
            if (-not $inv) { continue }
            foreach ($e in $inv.Errors) { $r.Unreadable++; $r.Differences.Add(@{ Status = 'UNREADABLE'; Path = $e.Path; Detail = $e.Reason }) }
        }
        $r.Incomplete = ($r.Unreadable -gt 0)
        if ($r.Differences.Count -gt 0) { Write-LogBlock -Lines @($r.Differences | ForEach-Object { "VERIFY {0}: {1} ({2})" -f $_.Status, $_.Path, $_.Detail }) -Level 'WARN' }
    } catch {
        Write-Bad "Verification could not run: $($_.Exception.Message)"
        Write-Log "CLOUD VERIFY FAILED: $($_.Exception.Message)" 'ERROR' -FileOnly
        return
    } finally { $sw.Stop(); Stop-JobUi }
    Show-VerifyResult -Title 'CLOUD VERIFY' -Result $r -Elapsed $sw.Elapsed -Note 'Compared every file name and exact size in bytes.'
    $vs = New-JobStats; $vs.Success = $r.Perfect; $vs.Failed = $r.Mismatch + $r.MissingInDst + $r.Unreadable; $vs.Cancelled = $r.Cancelled
    Register-JobResult -Operation 'Cloud verify' -Source $SrcData.FolderPath -Destination $DstData.FolderPath -Stats $vs -Elapsed $sw.Elapsed
}

# Option 4: send a cloud folder (or file) to the Recycle Bin.
function Perform-CloudWipe {
    param([hashtable]$TargetData)
    try { Connect-CloudSite -SiteUrl $TargetData.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; return }
    $rel = Get-SiteRelativePath -SiteUrl $TargetData.SiteUrl -ServerRelativePath $TargetData.FolderPath
    $file = Get-CloudFileIfFile -SiteRelativeUrl $rel
    $isRoot = (-not $file) -and (Test-IsLibraryRootPath -SiteUrl $TargetData.SiteUrl -SiteRelativePath $rel)
    $children = @()
    if (-not $file) {
        try { $children = (Get-CloudFolderItems -SiteRelativeUrl $rel) }
        catch { Write-Bad "The folder could not be opened: $($_.Exception.Message)"; return }
    }
    $lines = @(@{ Label = 'Target'; Value = $TargetData.FolderPath; Level = 'ERROR' })
    if ($file) { $lines += @{ Label = 'Type'; Value = "File ($(Format-FileSize ([long](Get-CloudProp $file 'Length' 0))))" } }
    else {
        $nf = @($children | Where-Object { Test-IsCloudFolder $_ }).Count
        $lines += @{ Label = 'Contains'; Value = ("{0} folders and {1} files at the top level" -f $nf, ($children.Count - $nf)) }
    }
    if ($isRoot) { $lines += @{ Text = 'This is the root of a document library. The library itself is kept; everything INSIDE it will be deleted.'; Color = 'Yellow' } }
    $lines += 'Deleted items go to the SharePoint Recycle Bin and can be restored for about 93 days.'
    Write-Panel -Title 'ABOUT TO DELETE (RECYCLE BIN)' -Lines $lines -Color Red
    if (-not (Read-Confirmation -Message 'Send this to the SharePoint Recycle Bin?' -Color Red)) { Write-Hint 'Nothing deleted.'; return }

    $stats = New-JobStats
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Start-JobUi
    try {
        Write-Log "CLOUD WIPE STARTED: $($TargetData.FolderPath)" 'WARN'
        if ($isRoot) {
            $n = 0
            foreach ($c in $children) {
                if ((Check-Interrupt -CurrentFile $c.Name) -eq 'CANCEL') { $stats.Cancelled = $true; break }
                $n++
                Update-Progress -Id 1 -Activity 'Cloud wipe' -Status ("Item {0} of {1}: {2}" -f $n, $children.Count, $c.Name) -Percent ($n * 100.0 / $children.Count)
                try {
                    Remove-CloudItem -SiteUrl $TargetData.SiteUrl -ServerRelativeUrl (Join-CloudPath $TargetData.FolderPath $c.Name) -IsFolder:(Test-IsCloudFolder $c)
                    $stats.Success++; Write-Log "Recycled: $($c.Name)" 'SUCCESS' -Detail
                } catch { Add-JobFailure -Stats $stats -File $c.Name -Path (Join-CloudPath $TargetData.FolderPath $c.Name) -Reason $_.Exception.Message }
            }
        } else {
            try {
                Remove-CloudItem -SiteUrl $TargetData.SiteUrl -ServerRelativeUrl $TargetData.FolderPath -IsFolder:(-not $file)
                $stats.Success++
                Write-Log "Recycled: $($TargetData.FolderPath)" 'SUCCESS'
            } catch { Add-JobFailure -Stats $stats -File (($TargetData.FolderPath -split '/')[-1]) -Path $TargetData.FolderPath -Reason $_.Exception.Message }
        }
    } finally { $sw.Stop(); Stop-JobUi }
    Show-TransferSummary -Title 'CLOUD WIPE' -Stats $stats -Elapsed $sw.Elapsed -Target $TargetData.FolderPath -Color Green -OkLabel 'Sent to Recycle Bin'
    Write-Hint "Recycle Bin: $(Get-SiteRecycleBinUrl $TargetData.SiteUrl)"
    Register-JobResult -Operation 'Cloud wipe' -Source $TargetData.FolderPath -Stats $stats -Elapsed $sw.Elapsed
}

# Option 6 and [S] in cloud batch screens. Assumes the site is connected.
function Perform-CloudSearch {
    param([hashtable]$CloudData, [string]$SearchTerm)
    $wp = New-SearchPattern $SearchTerm
    $base = Get-SiteRelativePath -SiteUrl $CloudData.SiteUrl -ServerRelativePath $CloudData.FolderPath
    $results = New-Object System.Collections.Generic.List[hashtable]
    $folders = 0; $files = 0; $errors = 0; $cancelled = $false
    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $tick = [System.Diagnostics.Stopwatch]::StartNew()
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue($base)
    Write-Log "CLOUD SEARCH for '$SearchTerm' in $($CloudData.FolderPath)" 'INFO' -FileOnly
    Start-JobUi
    try {
        while ($queue.Count -gt 0) {
            $cur = $queue.Dequeue()
            if ((Check-Interrupt -CurrentFolder "/$cur") -eq 'CANCEL') { $cancelled = $true; break }
            $folders++
            Start-Sleep -Milliseconds 100
            try { $items = Get-CloudFolderItems -SiteRelativeUrl $cur }
            catch {
                if ($_.Exception -is [System.OperationCanceledException]) { $cancelled = $true; break }
                $errors++; Write-Log "Search could not list /${cur}: $($_.Exception.Message)" 'WARN' -FileOnly; continue
            }
            $parentSrv = (Get-SitePath $CloudData.SiteUrl) + '/' + $cur
            foreach ($it in $items) {
                $isFolder = Test-IsCloudFolder $it
                if ($isFolder) { $queue.Enqueue((Join-CloudPath $cur $it.Name)) } else { $files++ }
                if ($wp.IsMatch($it.Name)) {
                    $srv = Get-CloudProp $it 'ServerRelativeUrl' (Join-CloudPath $parentSrv $it.Name)
                    $results.Add(@{
                        Type = $(if ($isFolder) { 'FOLDER' } else { 'FILE' }); Name = $it.Name; Path = $srv; Parent = $parentSrv
                        Size = $(if ($isFolder) { 0 } else { [long](Get-CloudProp $it 'Length' 0) }); Modified = (Get-CloudProp $it 'TimeLastModified')
                        Url = $(if ($isFolder) { Get-CloudFolderViewUrl -SiteUrl $CloudData.SiteUrl -ServerRelativeFolder $srv } else { Get-CloudWebUrl -SiteUrl $CloudData.SiteUrl -ServerRelativePath $srv -IsFile })
                        ParentUrl = (Get-CloudFolderViewUrl -SiteUrl $CloudData.SiteUrl -ServerRelativeFolder $parentSrv)
                    })
                    Write-Log "Match: $srv" 'SUCCESS' -ConsoleOnly -Detail
                }
            }
            if ($tick.ElapsedMilliseconds -ge 500) {
                Update-Progress -Id 1 -Activity "Searching SharePoint for '$SearchTerm'" -Status ("{0} matches  |  {1} folders, {2} files scanned" -f (Format-Count $results.Count), (Format-Count $folders), (Format-Count $files)) -Detail "/$cur"
                $tick.Restart()
            }
        }
    } finally { $sw.Stop(); Stop-JobUi }
    $state = if ($cancelled) { 'stopped early' } else { 'complete' }
    Write-Log ("Cloud search {0}: {1} matches, {2} folders and {3} files scanned in {4}" -f $state, $results.Count, $folders, $files, (Format-Elapsed $sw.Elapsed)) $(if ($cancelled) { 'WARN' } else { 'SUCCESS' })
    if ($errors) { Write-Caution "$errors folder(s) could not be listed (see the log) - results may be incomplete." }
    if ($results.Count -gt 0) { Write-LogBlock -Lines @($results | ForEach-Object { "MATCH [$($_.Type)] $($_.Path)" }) }
    Show-SearchResults -Results $results.ToArray() -Term $SearchTerm -Kind Cloud
}

function Invoke-CloudSearchLoop {
    param([hashtable]$CloudData)
    while ($true) {
        Write-Host ""
        Write-Hint "Search in: $($CloudData.FolderPath)"
        Write-Hint 'Examples:  invoice   *.pdf   *2024*report*   Budget [Final].xlsx'
        $term = (Read-Line -Prompt '  Search for (ENTER = back) > ' -LogLabel 'Cloud search term').Trim()
        if (-not $term -or $term -match '^(?i)q$') { return }
        Perform-CloudSearch -CloudData $CloudData -SearchTerm $term
    }
}

# Option 3: batch screen for a cloud folder - Copy / Move / Delete / Verify picked items.
function Perform-CloudBatch {
    param([hashtable]$SrcData, $DstData)
    if ($DstData -isnot [hashtable]) { $DstData = $null }
    try { Connect-CloudSite -SiteUrl $SrcData.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; return }
    $srcRel = Get-SiteRelativePath -SiteUrl $SrcData.SiteUrl -ServerRelativePath $SrcData.FolderPath
    Write-Log "CLOUD BATCH: $($SrcData.FolderPath)" 'INFO'
    while ($true) {
        $script:JobCancelled = $false
        try { Connect-CloudSite -SiteUrl $SrcData.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; return }
        Write-Panel -Title 'CLOUD BATCH' -Color Yellow -Lines @(
            @{ Label = 'Source'; Value = $SrcData.FolderPath }
            @{ Label = 'Destination'; Value = $(if ($DstData) { $DstData.FolderPath } else { '(not set - you will be asked if needed)' }) }
        )
        Write-Hint 'Reading the folder...'
        try { $items = Get-CloudIndexItems -SiteRelativeUrl $srcRel }
        catch { Write-Bad "The folder could not be listed: $($_.Exception.Message)"; return }
        if ($items.Count -eq 0) { Write-Caution 'This folder is empty.'; return }
        $rows = ConvertTo-CloudIndexRows -Items $items
        Show-ItemIndex -Title "CONTENTS OF $(($SrcData.FolderPath -split '/')[-1])" -Rows $rows -Color Yellow
        $sel = Read-IndexSelection -Count $rows.Count
        if ($sel.Action -eq 'Quit') { return }
        if ($sel.Action -eq 'Refresh') { continue }
        if ($sel.Action -eq 'Search') { Invoke-CloudSearchLoop -CloudData $SrcData; continue }
        $act = Read-Choice -Prompt 'C = copy   M = move   D = delete (Recycle Bin)   V = verify against destination' -Valid @('C', 'M', 'D', 'V')
        if (-not $act) { continue }
        $chosen = @($sel.Indices | ForEach-Object { $items[$_] })
        $verb = switch ($act) { 'C' { 'COPY' } 'M' { 'MOVE' } 'D' { 'DELETE' } 'V' { 'VERIFY' } }
        if ($act -ne 'D' -and -not $DstData) {
            $d = Read-CloudFolder -Prompt 'Destination folder URL for these items:' -LogLabel 'Cloud batch destination'
            if ($null -eq $d -or $d -eq '') { continue }
            if (-not (Test-SameTenant $SrcData.SiteUrl $d.SiteUrl)) { Write-Bad 'That destination is in a different tenant. Use download + upload instead.'; continue }
            $DstData = $d
        }
        if ($act -eq 'D' -and -not (Read-Confirmation -Message ("Send {0} item(s) to the SharePoint Recycle Bin?" -f $chosen.Count) -Color Red)) { Write-Hint 'Nothing deleted.'; continue }

        $stats = New-JobStats
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $checkedOut = New-Object System.Collections.Generic.List[string]
        if ($act -eq 'V') {
            Invoke-CloudBatchVerify -SrcData $SrcData -DstData $DstData -Items $chosen
            continue
        }
        Start-JobUi
        try {
            if ($act -in @('C', 'M') -and $SrcData.SiteUrl -ieq $DstData.SiteUrl) {
                $dRel = Get-SiteRelativePath -SiteUrl $DstData.SiteUrl -ServerRelativePath $DstData.FolderPath
                if ($dRel) { $null = Ensure-CloudFolder -SiteRelativePath $dRel }
            }
            $n = 0
            foreach ($it in $chosen) {
                if ((Check-Interrupt -CurrentFile $it.Name -CurrentFolder $SrcData.FolderPath) -eq 'CANCEL') { $stats.Cancelled = $true; break }
                $n++
                $isFolder = Test-IsCloudFolder $it
                Update-Progress -Id 1 -Activity "Cloud batch $($verb.ToLower())" -Status ("Item {0} of {1}: {2}" -f $n, $chosen.Count, $it.Name) -Percent (($n - 1) * 100.0 / $chosen.Count) -Detail 'P = pause, C = cancel'
                Write-BatchItemHeader -Index $n -Total $chosen.Count -Name $it.Name -Verb $verb -Color Yellow
                $srcUrl = Join-CloudPath $SrcData.FolderPath $it.Name
                try {
                    if ($act -eq 'D') {
                        Remove-CloudItem -SiteUrl $SrcData.SiteUrl -ServerRelativeUrl $srcUrl -IsFolder:$isFolder
                        Write-Ok "Sent to the Recycle Bin: $($it.Name)"
                    } else {
                        # Folders (and anything going to another site) are copied INTO the destination
                        # folder; files on the same site get their full target path.
                        $target = if ($isFolder -or $SrcData.SiteUrl -ine $DstData.SiteUrl) { $DstData.FolderPath } else { Join-CloudPath $DstData.FolderPath $it.Name }
                        Invoke-CloudCopyMove -SourceUrl $srcUrl -TargetUrl $target -Move:($act -eq 'M') -Name $it.Name
                        Write-Ok ("{0}: {1}" -f $(if ($act -eq 'M') { 'Moved' } else { 'Copied' }), $it.Name)
                    }
                    $stats.Success++
                    if (-not $isFolder) { $stats.SuccessBytes += [long](Get-CloudProp $it 'Length' 0) }
                } catch [System.OperationCanceledException] {
                    $stats.Cancelled = $true; break
                } catch {
                    $msg = $_.Exception.Message
                    if ($msg -match '(?i)checked out') { $co = Get-CheckoutInfo $msg; $checkedOut.Add("$($it.Name): '$($co.File)' is checked out by $($co.User)") }
                    Write-Log "FAILED: $($it.Name) - $msg" 'ERROR'
                    Add-JobFailure -Stats $stats -File $it.Name -Path $srcUrl -Reason $msg
                }
            }
        } finally { $sw.Stop(); Stop-JobUi }
        $okLabel = switch ($act) { 'C' { 'Copied (items)' } 'M' { 'Moved (items)' } default { 'Recycled (items)' } }
        Show-TransferSummary -Title "CLOUD BATCH $verb" -Stats $stats -Elapsed $sw.Elapsed -Color Yellow -OkLabel $okLabel
        if ($checkedOut.Count -gt 0) {
            $lines = @($checkedOut | ForEach-Object { "- $_" }) + @('-', 'Ask the user to check the file in (or discard the checkout), then run this again.')
            Write-Panel -Title 'BLOCKED BY CHECKED-OUT FILES' -Lines $lines -Color Yellow
        }
        if ($act -eq 'D') { Write-Hint "Recycle Bin: $(Get-SiteRecycleBinUrl $SrcData.SiteUrl)" }
        Register-JobResult -Operation "Cloud batch $($verb.ToLower())" -Source $SrcData.FolderPath -Destination $(if ($DstData) { $DstData.FolderPath } else { '' }) -Stats $stats -Elapsed $sw.Elapsed
        Show-Pause -Message 'Press any key to return to the folder list...'
    }
}

# Verify picked cloud items against the same names in the destination folder. All source
# items are scanned first, then the destination (one site switch, even across sites).
function Invoke-CloudBatchVerify {
    param([hashtable]$SrcData, [hashtable]$DstData, [object[]]$Items)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $srcMap = @{}; $dstMap = @{}
    $unread = New-Object System.Collections.Generic.List[hashtable]
    $cancelled = $false
    Start-JobUi
    try {
        Connect-CloudSite -SiteUrl $SrcData.SiteUrl
        foreach ($it in $Items) {
            if ($script:JobCancelled) { $cancelled = $true; break }
            if (Test-IsCloudFolder $it) {
                $inv = Get-CloudInventory -SiteUrl $SrcData.SiteUrl -FolderPath (Join-CloudPath $SrcData.FolderPath $it.Name) -Tag "Source: $($it.Name)"
                foreach ($k in (ConvertTo-InventoryMap $inv $it.Name).GetEnumerator()) { $srcMap[$k.Key] = $k.Value }
                foreach ($e in $inv.Errors) { $unread.Add($e) }
                if ($inv.Cancelled) { $cancelled = $true }
            } else { $srcMap[$it.Name.ToLowerInvariant()] = @{ Length = [long](Get-CloudProp $it 'Length' 0); Path = $it.Name } }
        }
        if (-not $cancelled) {
            Connect-CloudSite -SiteUrl $DstData.SiteUrl
            $dstRel = Get-SiteRelativePath -SiteUrl $DstData.SiteUrl -ServerRelativePath $DstData.FolderPath
            $dstTop = @{}
            try { foreach ($x in (Get-CloudFolderItems -SiteRelativeUrl $dstRel)) { $dstTop[$x.Name.ToLowerInvariant()] = $x } }
            catch { $unread.Add(@{ File = '(destination)'; Path = $DstData.FolderPath; Reason = "Destination could not be listed: $($_.Exception.Message)" }) }
            foreach ($it in $Items) {
                if ($script:JobCancelled) { $cancelled = $true; break }
                $key = $it.Name.ToLowerInvariant()
                if (-not $dstTop.ContainsKey($key)) { continue }
                $d = $dstTop[$key]
                if (Test-IsCloudFolder $d) {
                    $inv = Get-CloudInventory -SiteUrl $DstData.SiteUrl -FolderPath (Join-CloudPath $DstData.FolderPath $d.Name) -Tag "Destination: $($d.Name)"
                    foreach ($k in (ConvertTo-InventoryMap $inv $it.Name).GetEnumerator()) { $dstMap[$k.Key] = $k.Value }
                    foreach ($e in $inv.Errors) { $unread.Add($e) }
                    if ($inv.Cancelled) { $cancelled = $true; break }
                } else { $dstMap[$key] = @{ Length = [long](Get-CloudProp $d 'Length' 0); Path = $d.Name } }
            }
        }
    } catch {
        Write-Bad "Verification could not run: $($_.Exception.Message)"
        return
    } finally { $sw.Stop(); Stop-JobUi }
    # Only compare the destination entries that belong to the selected items.
    $r = Compare-CloudInventories -SrcMap $srcMap -DstMap $dstMap
    $r.Cancelled = $cancelled
    foreach ($e in $unread) { $r.Unreadable++; $r.Differences.Add(@{ Status = 'UNREADABLE'; Path = $e.Path; Detail = $e.Reason }) }
    if ($r.Differences.Count -gt 0) { Write-LogBlock -Lines @($r.Differences | ForEach-Object { "VERIFY {0}: {1} ({2})" -f $_.Status, $_.Path, $_.Detail }) -Level 'WARN' }
    Show-VerifyResult -Title 'CLOUD BATCH VERIFY' -Result $r -Elapsed $sw.Elapsed -Note 'Compared every file name and exact size in bytes.'
    $vs = New-JobStats; $vs.Success = $r.Perfect; $vs.Failed = $r.Mismatch + $r.MissingInDst + $r.Unreadable; $vs.Cancelled = $r.Cancelled
    Register-JobResult -Operation 'Cloud batch verify' -Source $SrcData.FolderPath -Destination $DstData.FolderPath -Stats $vs -Elapsed $sw.Elapsed
    Show-Pause -Message 'Press any key to return to the folder list...'
}

# ======================================================================================
# SECTION 11 - DOWNLOAD ENGINES (SharePoint / OneDrive -> PC)
# ======================================================================================

# Office files can legitimately come down slightly bigger or smaller than SharePoint
# reports (SharePoint writes library metadata into the document on download). For these
# types a SMALL difference is accepted and reported; anything else is an error.
function Test-OfficeMetadataFile {
    param([string]$Name)
    $ext = [System.IO.Path]::GetExtension("$Name").ToLowerInvariant()
    return ($ext -in @('.docx', '.docm', '.dotx', '.dotm', '.xlsx', '.xlsm', '.xltx', '.xltm', '.pptx', '.pptm', '.potx', '.potm', '.doc', '.dot', '.xls', '.xlt', '.ppt', '.pot', '.vsdx', '.vsdm'))
}

function Test-AcceptableSizeDifference {
    param([string]$Name, [long]$CloudSize, [long]$LocalSize, [string]$LocalPath = '')
    if ($LocalSize -eq $CloudSize) { return $true }
    if (-not (Test-OfficeMetadataFile $Name) -or $LocalSize -le 0) { return $false }
    $limit = [Math]::Max(256KB, [long]($CloudSize * 0.05))
    if ([Math]::Abs($LocalSize - $CloudSize) -gt $limit) { return $false }
    # Modern Office files are ZIP packages - check the signature as a sanity test.
    $ext = [System.IO.Path]::GetExtension($Name).ToLowerInvariant()
    if ($LocalPath -and $ext -match '^\.(doc|xls|ppt|pot|dot|xlt|vsd)[xm]$') {
        try {
            $fs = [System.IO.File]::OpenRead($LocalPath)
            try { $b1 = $fs.ReadByte(); $b2 = $fs.ReadByte() } finally { $fs.Dispose() }
            if ($b1 -ne 0x50 -or $b2 -ne 0x4B) { return $false }
        } catch { return $false }
    }
    return $true
}

function ConvertTo-UtcTime {
    param($Value)
    if (-not $Value) { return $null }
    try {
        $d = [datetime]$Value
        if ($d.Kind -eq [System.DateTimeKind]::Unspecified) { $d = [datetime]::SpecifyKind($d, [System.DateTimeKind]::Utc) }
        return $d.ToUniversalTime()
    } catch { return $null }
}

# CSOM binary-stream fallback for files the normal download cannot fetch.
function Save-CloudFileStream {
    param([string]$ServerRelativeUrl, [string]$Dest, $UniqueId = $null)
    $ctx = Get-PnPContext
    $file = $null
    if ($UniqueId) { try { $file = $ctx.Web.GetFileById([guid]"$UniqueId") } catch { $file = $null } }
    if (-not $file) {
        try { $file = $ctx.Web.GetFileByServerRelativePath([Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($ServerRelativeUrl)) }
        catch { $file = $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl) }
    }
    $stream = $file.OpenBinaryStream()
    $ctx.ExecuteQuery()
    if ($null -eq $stream.Value) { throw 'The binary download stream was empty.' }
    $fs = $null
    try {
        $fs = [System.IO.File]::Create($Dest)
        $stream.Value.CopyTo($fs)
    } finally {
        if ($fs) { $fs.Dispose() }
        try { $stream.Value.Dispose() } catch { }
    }
}

# Downloads one file to a temp name in the buffer, checks the size, then moves it into
# place in one step (an existing file is only replaced once the new one is complete).
function Perform-SafeDownload {
    param(
        [string]$ServerRelativeUrl, [string]$LongDestPath, [string]$TempDir,
        [long]$ExpectedSize = -1, [string]$Name = '', $UniqueId = $null
    )
    try { if (-not [System.IO.Directory]::Exists($TempDir)) { $null = [System.IO.Directory]::CreateDirectory($TempDir) } } catch { }
    $sdTmpName = "rhfm_{0}.tmp" -f [guid]::NewGuid().ToString('N')
    $sdTmpPath = Join-Path $TempDir $sdTmpName
    try {
        $null = Invoke-WithRetry -OperationName "Download $Name" -Quiet -ScriptBlock {
            if ([System.IO.File]::Exists($sdTmpPath)) { [System.IO.File]::Delete($sdTmpPath) }
            try {
                Get-PnPFile -ServerRelativeUrl $ServerRelativeUrl -Path $TempDir -FileName $sdTmpName -AsFile -Force -ErrorAction Stop
            } catch {
                if ($_.Exception.Message -match '(?i)not found|does not exist|\(40[134]\)|access (is )?denied') { throw }
                Write-Log "Standard download failed for '$Name' ($($_.Exception.Message)) - trying the binary-stream fallback." 'WARN' -FileOnly
                Save-CloudFileStream -ServerRelativeUrl $ServerRelativeUrl -Dest $sdTmpPath -UniqueId $UniqueId
            }
            if (-not [System.IO.File]::Exists($sdTmpPath)) { throw 'The download did not produce a file.' }
            if ($ExpectedSize -ge 0) {
                $sdGot = (New-Object System.IO.FileInfo($sdTmpPath)).Length
                if (-not (Test-AcceptableSizeDifference -Name $Name -CloudSize $ExpectedSize -LocalSize $sdGot -LocalPath $sdTmpPath)) {
                    throw ("Size check failed: received {0} bytes but SharePoint reports {1} bytes." -f $sdGot, $ExpectedSize)
                }
            }
        }
        $dir = [System.IO.Path]::GetDirectoryName($LongDestPath)
        if (-not [System.IO.Directory]::Exists($dir)) { $null = [System.IO.Directory]::CreateDirectory($dir) }
        Move-FileReplace -Source (Get-LongPath $sdTmpPath) -Dest $LongDestPath
        return @{ Success = $true; Error = $null; Size = (New-Object System.IO.FileInfo($LongDestPath)).Length; Cancelled = $false }
    } catch {
        try { if ([System.IO.File]::Exists($sdTmpPath)) { [System.IO.File]::Delete($sdTmpPath) } } catch { }
        return @{ Success = $false; Error = $_.Exception.Message; Size = 0; Cancelled = ($_.Exception -is [System.OperationCanceledException] -or $script:JobCancelled) }
    }
}

# Local names for a cloud listing, assigned in sorted order so they are identical on
# every run (and identical to what a full-folder download of the parent would produce).
function Get-LocalNameMap {
    param([object[]]$Items, $Used = $null)
    if ($null -eq $Used) { $Used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase) }
    $map = @{}
    foreach ($it in @($Items | Sort-Object -Property Name)) { $map[$it.Name] = Get-UniqueLocalName -Name (Get-SafeWindowsName $it.Name) -Used $Used }
    return $map
}

# Downloads (or skips) one cloud file and records the outcome in -Stats. With -Manifest
# every file that ends up on disk is recorded for archive verification.
function Invoke-DownloadFile {
    param(
        $Item, [string]$ServerRelativeUrl, [string]$CloudRel, [string]$LocalPath,
        [string]$BufferDest, [hashtable]$Stats, [switch]$Manifest
    )
    $len  = [long](Get-CloudProp $Item 'Length' 0)
    $mod  = ConvertTo-UtcTime (Get-CloudProp $Item 'TimeLastModified')
    $long = Get-LongPath $LocalPath
    if ([System.IO.File]::Exists($long)) {
        $li = New-Object System.IO.FileInfo($long)
        # Downloads are stamped with the SharePoint date, so a file from an earlier run of
        # this tool matches on date AND size. A different file that merely has the same
        # size is not trusted.
        $sameTime = (-not $mod) -or (Test-SameFileTime $li.LastWriteTimeUtc $mod)
        if ($sameTime -and (Test-AcceptableSizeDifference -Name $Item.Name -CloudSize $len -LocalSize $li.Length -LocalPath $long)) {
            $Stats.Skipped++; $Stats.SkippedBytes += $li.Length
            if ($Manifest) { $Stats.Manifest.Add(@{ CloudRel = $CloudRel; Local = $long; CloudSize = $len; LocalSize = $li.Length }) }
            Write-Log "Already downloaded, skipped: $CloudRel" 'INFO' -ConsoleOnly -Detail
            return
        }
        if ($Manifest) {
            # Archive safety: never overwrite a file that is already in the archive folder (it
            # may be the only copy of something archived earlier). Download a side copy and
            # keep the existing file only if it is byte-for-byte the same (this also resumes
            # archives started with older versions, which did not keep SharePoint dates).
            $check = $long + '.rhfmcheck'
            $rc = Perform-SafeDownload -ServerRelativeUrl $ServerRelativeUrl -LongDestPath $check -TempDir $BufferDest -ExpectedSize $len -Name $Item.Name -UniqueId (Get-CloudProp $Item 'UniqueId')
            if (-not $rc.Success) {
                if ($rc.Cancelled) { $Stats.Cancelled = $true; return }
                Add-JobFailure -Stats $Stats -File $Item.Name -Path $ServerRelativeUrl -Reason $rc.Error
                return
            }
            $same = $false
            try { $same = ((New-Object System.IO.FileInfo($check)).Length -eq $li.Length) -and ((Get-FileSha256 $check) -eq (Get-FileSha256 $long)) } catch { $same = $false }
            try { [System.IO.File]::Delete($check) } catch { }
            if ($same) {
                if ($mod) { try { [System.IO.File]::SetLastWriteTimeUtc($long, $mod) } catch { } }
                $Stats.Skipped++; $Stats.SkippedBytes += $li.Length
                $Stats.Manifest.Add(@{ CloudRel = $CloudRel; Local = $long; CloudSize = $len; LocalSize = $li.Length })
                Write-Log "Already in the archive (content checked), skipped: $CloudRel" 'INFO' -ConsoleOnly -Detail
                return
            }
            Add-JobFailure -Stats $Stats -File $Item.Name -Path $LocalPath -Reason 'A DIFFERENT file with this name is already in the archive folder and was not overwritten. Move it away (or use an empty folder) and run again.'
            Write-Log "NOT OVERWRITTEN (archive): $LocalPath already exists with different content." 'ERROR'
            return
        }
    }
    if ($len -ge (Get-LargeFileBytes)) { Write-Host ("  Downloading large file {0} ({1}) - please wait..." -f $Item.Name, (Format-FileSize $len)) -ForegroundColor Yellow }
    $r = Perform-SafeDownload -ServerRelativeUrl $ServerRelativeUrl -LongDestPath $long -TempDir $BufferDest -ExpectedSize $len -Name $Item.Name -UniqueId (Get-CloudProp $Item 'UniqueId')
    if ($r.Success) {
        if ($mod) { try { [System.IO.File]::SetLastWriteTimeUtc($long, $mod) } catch { } }
        $Stats.Success++; $Stats.SuccessBytes += $r.Size; $script:SessionBytes += $r.Size
        if ($Stats.Names.Count -lt 200000) { $Stats.Names.Add($Item.Name) }
        if ($r.Size -ne $len) {
            $Stats.Accepted++
            Write-Log ("Office file size differs from SharePoint (metadata): {0} - cloud {1} bytes, received {2} bytes" -f $CloudRel, $len, $r.Size) 'WARN' -FileOnly
        }
        if ($Manifest) { $Stats.Manifest.Add(@{ CloudRel = $CloudRel; Local = $long; CloudSize = $len; LocalSize = $r.Size }) }
        Write-Log ("Downloaded {0} ({1})" -f $CloudRel, (Format-FileSize $r.Size)) 'SUCCESS' -ConsoleOnly -Detail
    } elseif ($r.Cancelled) {
        $Stats.Cancelled = $true
    } else {
        Write-Log "FAILED: $CloudRel - $($r.Error)" 'ERROR'
        Add-JobFailure -Stats $Stats -File $Item.Name -Path $ServerRelativeUrl -Reason $r.Error
    }
}

# Option 13 (and the engine behind every archive): downloads a cloud folder - or a single
# file - into a local folder. Requires an open connection to the site.
function Perform-DirectDownload {
    param(
        [hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest,
        [switch]$Nested, [switch]$Manifest, [switch]$NoRegister, [string]$Title = 'DOWNLOAD'
    )
    $stats = New-JobStats
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $barId = if ($Nested) { 2 } else { 1 }
    $leaf = ($CloudData.FolderPath.TrimEnd('/') -split '/')[-1]
    $localRoot = Resolve-LocalPath $LocalDest
    Start-JobUi
    try {
        if (-not $Nested) {
            Write-Log "DOWNLOAD STARTED: $($CloudData.FolderPath) -> $(Format-PathWithStorageTag $localRoot)" 'INFO'
            Show-DestinationFreeSpace -DestPath $localRoot
        } else { Write-Log "DOWNLOAD: $($CloudData.FolderPath) -> $localRoot" 'INFO' -FileOnly }
        $rootLong = Get-LongPath $localRoot
        if (-not [System.IO.Directory]::Exists($rootLong)) { $null = [System.IO.Directory]::CreateDirectory($rootLong) }
        $base = Get-SiteRelativePath -SiteUrl $CloudData.SiteUrl -ServerRelativePath $CloudData.FolderPath
        $single = Get-CloudFileIfFile -SiteRelativeUrl $base
        if ($single) {
            Write-Hint 'The address points to a single file.'
            $local = Join-LocalPath $localRoot (Get-SafeWindowsName $single.Name)
            Invoke-DownloadFile -Item $single -ServerRelativeUrl $CloudData.FolderPath -CloudRel $single.Name -LocalPath $local -BufferDest $BufferDest -Stats $stats -Manifest:$Manifest
        } else {
            $queue = New-Object System.Collections.Generic.Queue[hashtable]
            $queue.Enqueue(@{ Site = $base; Rel = ''; Local = $localRoot; Srv = $CloudData.FolderPath.TrimEnd('/') })
            $tick = [System.Diagnostics.Stopwatch]::StartNew()
            while ($queue.Count -gt 0) {
                if ($stats.Cancelled) { break }
                $cur = $queue.Dequeue()
                if ((Check-Interrupt -CurrentFolder "/$($cur.Site)") -eq 'CANCEL') { $stats.Cancelled = $true; break }
                Start-Sleep -Milliseconds 150   # gentle pacing between folders (anti-throttling)
                try { $items = Get-CloudFolderItems -SiteRelativeUrl $cur.Site }
                catch {
                    if ($_.Exception -is [System.OperationCanceledException]) { $stats.Cancelled = $true; break }
                    $stats.ListingErrors++
                    Add-JobFailure -Stats $stats -File $(if ($cur.Rel) { $cur.Rel } else { $leaf }) -Path $cur.Srv -Reason "Folder could not be listed: $($_.Exception.Message)"
                    Write-Log "FOLDER NOT LISTED: $($cur.Srv) - $($_.Exception.Message)" 'ERROR'
                    continue
                }
                # Sorted so name-clash suffixes are the same on every run (safe resume).
                $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($it in @($items | Sort-Object -Property Name)) {
                    if ($stats.Cancelled) { break }
                    if ((Check-Interrupt -CurrentFile $it.Name -CurrentFolder $cur.Srv) -eq 'CANCEL') { $stats.Cancelled = $true; break }
                    $safe = Get-SafeWindowsName $it.Name
                    $uniq = Get-UniqueLocalName -Name $safe -Used $used
                    if ($uniq -cne $it.Name) { Write-Log "Saved '$($it.Name)' as '$uniq' (Windows naming rules / name clash)." 'INFO' -FileOnly }
                    $rel = if ($cur.Rel) { "$($cur.Rel)/$($it.Name)" } else { $it.Name }
                    $localPath = Join-LocalPath $cur.Local $uniq
                    $srv = Get-CloudProp $it 'ServerRelativeUrl' (Join-CloudPath $cur.Srv $it.Name)
                    if (Test-IsCloudFolder $it) {
                        try { $null = [System.IO.Directory]::CreateDirectory((Get-LongPath $localPath)) }
                        catch { Add-JobFailure -Stats $stats -File $it.Name -Path $localPath -Reason "Local folder could not be created: $($_.Exception.Message)"; continue }
                        $queue.Enqueue(@{ Site = (Join-CloudPath $cur.Site $it.Name); Rel = $rel; Local = $localPath; Srv = $srv })
                    } else {
                        Invoke-DownloadFile -Item $it -ServerRelativeUrl $srv -CloudRel $rel -LocalPath $localPath -BufferDest $BufferDest -Stats $stats -Manifest:$Manifest
                    }
                    if ($tick.ElapsedMilliseconds -ge 500) {
                        $act = if ($Nested) { "Downloading $leaf" } else { "Download: $leaf" }
                        Update-Progress -Id $barId -Activity $act -Status ("{0} downloaded ({1}), {2} skipped, {3} failed" -f (Format-Count $stats.Success), (Format-FileSize $stats.SuccessBytes), (Format-Count $stats.Skipped), (Format-Count $stats.Failed)) -Detail ("/{0}   P = pause, C = cancel" -f $rel)
                        $tick.Restart()
                    }
                }
            }
        }
        Update-Progress -Id $barId -Activity 'Download' -Completed
    } catch {
        Write-Log "DOWNLOAD ERROR: $($_.Exception.Message)" 'ERROR'
        Add-JobFailure -Stats $stats -File $leaf -Path $CloudData.FolderPath -Reason $_.Exception.Message
    } finally { $sw.Stop(); Stop-JobUi }

    if ($Nested) {
        Write-Log ("DOWNLOAD {0}: {1} downloaded, {2} skipped, {3} failed, {4}" -f $leaf, $stats.Success, $stats.Skipped, $stats.Failed, (Format-FileSize $stats.SuccessBytes)) $(if ($stats.Failed) { 'WARN' } else { 'INFO' }) -FileOnly
    } else {
        $tag = Get-PathStorageType $localRoot
        Show-TransferSummary -Title $Title -Stats $stats -Elapsed $sw.Elapsed -Target $leaf -Storage $(if ($tag) { "Cloud -> $tag" } else { '' }) -Color Magenta -OkLabel 'Downloaded'
        if (-not $NoRegister) { Register-JobResult -Operation 'Download' -Source $CloudData.FolderPath -Destination $localRoot -Stats $stats -Elapsed $sw.Elapsed }
    }
    return $stats
}

# Option 14: pick items from a cloud folder and download them. Requires a connection.
function Perform-CloudToLocalBatchDownload {
    param([hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest)
    $base = Get-SiteRelativePath -SiteUrl $CloudData.SiteUrl -ServerRelativePath $CloudData.FolderPath
    if (Get-CloudFileIfFile -SiteRelativeUrl $base) {
        Write-Info 'The address points to a single file - downloading it directly.'
        $null = Perform-DirectDownload -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest
        return
    }
    $localRoot = Resolve-LocalPath $LocalDest
    while ($true) {
        $script:JobCancelled = $false
        try { Connect-CloudSite -SiteUrl $CloudData.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; return }
        Write-Panel -Title 'BATCH DOWNLOAD' -Color Magenta -Lines @(
            @{ Label = 'From (cloud)'; Value = $CloudData.FolderPath }
            @{ Label = 'To (this PC)'; Value = $localRoot }
            @{ Label = 'Temp buffer';  Value = $BufferDest }
        )
        Write-Hint 'Reading the folder...'
        try { $items = Get-CloudIndexItems -SiteRelativeUrl $base } catch { Write-Bad "The folder could not be listed: $($_.Exception.Message)"; return }
        if ($items.Count -eq 0) { Write-Caution 'This folder is empty.'; return }
        $rows = ConvertTo-CloudIndexRows -Items $items
        Show-ItemIndex -Title "CONTENTS OF $(($CloudData.FolderPath -split '/')[-1])" -Rows $rows -Color Magenta
        $sel = Read-IndexSelection -Count $rows.Count -Verb 'download'
        if ($sel.Action -eq 'Quit') { return }
        if ($sel.Action -eq 'Refresh') { continue }
        if ($sel.Action -eq 'Search') { Invoke-CloudSearchLoop -CloudData $CloudData; continue }
        $chosen = @($sel.Indices | ForEach-Object { $items[$_] })
        Show-DestinationFreeSpace -DestPath $localRoot
        $batch = New-JobStats
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $nameMap = Get-LocalNameMap -Items $items
        Start-JobUi
        try {
            $rootLong = Get-LongPath $localRoot
            if (-not [System.IO.Directory]::Exists($rootLong)) { $null = [System.IO.Directory]::CreateDirectory($rootLong) }
            $n = 0
            foreach ($it in $chosen) {
                if ($script:JobCancelled) { $batch.Cancelled = $true; break }
                $n++
                Update-Progress -Id 1 -Activity 'Batch download' -Status ("Item {0} of {1}: {2}" -f $n, $chosen.Count, $it.Name) -Percent (($n - 1) * 100.0 / $chosen.Count) -Detail 'P = pause, C = cancel'
                Write-BatchItemHeader -Index $n -Total $chosen.Count -Name $it.Name -Verb 'DOWNLOAD' -Color Magenta
                $uniq = $nameMap[$it.Name]
                $local = Join-LocalPath $localRoot $uniq
                $srv = Get-CloudProp $it 'ServerRelativeUrl' (Join-CloudPath $CloudData.FolderPath $it.Name)
                if (Test-IsCloudFolder $it) {
                    $res = Perform-DirectDownload -CloudData @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $srv } -LocalDest $local -BufferDest $BufferDest -Nested
                    Merge-JobStats -Into $batch -From $res
                    Write-Host ("  {0} downloaded, {1} skipped, {2} failed" -f $res.Success, $res.Skipped, $res.Failed) -ForegroundColor $(if ($res.Failed) { 'Yellow' } else { 'DarkGray' })
                } else {
                    Invoke-DownloadFile -Item $it -ServerRelativeUrl $srv -CloudRel $it.Name -LocalPath $local -BufferDest $BufferDest -Stats $batch
                }
                if ($script:JobCancelled) { $batch.Cancelled = $true }
            }
        } finally { $sw.Stop(); Stop-JobUi }
        $tag = Get-PathStorageType $localRoot
        Show-TransferSummary -Title 'BATCH DOWNLOAD' -Stats $batch -Elapsed $sw.Elapsed -Storage $(if ($tag) { "Cloud -> $tag" } else { '' }) -Color Magenta -OkLabel 'Downloaded'
        Register-JobResult -Operation 'Batch download' -Source $CloudData.FolderPath -Destination $localRoot -Stats $batch -Elapsed $sw.Elapsed
        Show-Pause -Message 'Press any key to return to the folder list...'
    }
}

# ======================================================================================
# SECTION 12 - UPLOAD ENGINES (PC -> SharePoint / OneDrive)
# ======================================================================================

# Creates the whole folder path in SharePoint if it is missing (Resolve-PnPFolder).
function Ensure-CloudFolder {
    param([string]$SiteRelativePath)
    if ([string]::IsNullOrWhiteSpace($SiteRelativePath)) { return $true }
    try {
        $null = Invoke-WithRetry -OperationName "Create folder /$SiteRelativePath" -Quiet -ScriptBlock {
            $null = Resolve-PnPFolder -SiteRelativePath $SiteRelativePath -ErrorAction Stop
        }
        return $true
    } catch {
        Write-Log "Could not create or open cloud folder /${SiteRelativePath}: $($_.Exception.Message)" 'WARN' -FileOnly
        return $false
    }
}

# Name -> size of the files already in one cloud folder (one listing per folder instead
# of one request per file). Returns an empty map for a folder that does not exist yet and
# $null when the folder could not be read (callers then check file by file).
function Get-CloudFileMap {
    param([string]$SiteRelFolder, [hashtable]$Cache)
    if ($Cache.ContainsKey($SiteRelFolder)) { return $Cache[$SiteRelFolder] }
    $map = @{}
    try {
        foreach ($it in (Get-CloudFolderItems -SiteRelativeUrl $SiteRelFolder -IncludeSystem)) {
            if (-not (Test-IsCloudFolder $it)) { $map[$it.Name.ToLowerInvariant()] = [long](Get-CloudProp $it 'Length' 0) }
        }
    } catch {
        if ($_.Exception.Message -notmatch '(?i)not found|does not exist|\(404\)') { $map = $null }
    }
    $Cache[$SiteRelFolder] = $map
    return $map
}

$script:UploadJunk = @('Thumbs.db', 'desktop.ini', '.DS_Store', 'ehthumbs.db')

# Uploads (or skips) one local file into a cloud folder and records the outcome in -Stats.
function Invoke-UploadFile {
    param(
        [System.IO.FileInfo]$File, [string]$CloudFolder, [string]$RelDisplay,
        [hashtable]$Stats, [hashtable]$FolderCache, [hashtable]$MapCache
    )
    $ufName = $File.Name
    if ($ufName -like '~$*' -or $script:UploadJunk -contains $ufName) {
        $Stats.Skipped++; $Stats.SkippedBytes += $File.Length
        Write-Log "Skipped temporary / system file: $RelDisplay" 'INFO' -ConsoleOnly -Detail
        return
    }
    $ufCloudName = Get-SafeWindowsName $ufName
    $ufFull = Join-CloudPath $CloudFolder $ufCloudName
    if ($ufFull.Length -ge $CloudPathLimit) {
        $Stats.TooLong++
        Write-Log ("SKIPPED (cloud path {0} chars, limit {1}): {2}" -f $ufFull.Length, $CloudPathLimit, $RelDisplay) 'WARN' -Detail
        Write-LongPathReport -Source (Remove-LongPrefix $File.FullName) -Dest $ufFull -Length $ufFull.Length
        return
    }
    if (-not $FolderCache.ContainsKey($CloudFolder)) { $FolderCache[$CloudFolder] = Ensure-CloudFolder -SiteRelativePath $CloudFolder }
    $map = Get-CloudFileMap -SiteRelFolder $CloudFolder -Cache $MapCache
    $exists = $false
    if ($null -ne $map) {
        $key = $ufCloudName.ToLowerInvariant()
        if ($map.ContainsKey($key) -and [long]$map[$key] -eq $File.Length) { $exists = $true }
    } else {
        try { $ex = Get-PnPFile -Url $ufFull -ErrorAction Stop; if ($ex -and [long](Get-CloudProp $ex 'Length' -1) -eq $File.Length) { $exists = $true } } catch { }
    }
    if ($exists) {
        $Stats.Skipped++; $Stats.SkippedBytes += $File.Length
        Write-Log "Already in SharePoint, skipped: $RelDisplay" 'INFO' -ConsoleOnly -Detail
        return
    }
    if ($File.Length -ge (Get-LargeFileBytes)) { Write-Host ("  Uploading large file {0} ({1}) - please wait..." -f $ufName, (Format-FileSize $File.Length)) -ForegroundColor Yellow }
    try {
        $null = Invoke-WithRetry -OperationName "Upload $ufName" -Quiet -ScriptBlock {
            $null = Add-PnPFile -Path $File.FullName -Folder $CloudFolder -NewFileName $ufCloudName -ErrorAction Stop
        }
        $Stats.Success++; $Stats.SuccessBytes += $File.Length; $script:SessionBytes += $File.Length
        if ($Stats.Names.Count -lt 200000) { $Stats.Names.Add($ufName) }
        if ($null -ne $map) { $map[$ufCloudName.ToLowerInvariant()] = $File.Length }
        Write-Log ("Uploaded {0} ({1})" -f $RelDisplay, (Format-FileSize $File.Length)) 'SUCCESS' -ConsoleOnly -Detail
    } catch [System.OperationCanceledException] {
        $Stats.Cancelled = $true
    } catch {
        if ($script:JobCancelled) { $Stats.Cancelled = $true; return }
        Write-Log "FAILED: $RelDisplay - $($_.Exception.Message)" 'ERROR'
        Add-JobFailure -Stats $Stats -File $ufName -Path $RelDisplay -Reason $_.Exception.Message
    }
}

# Option 15: uploads a local folder (all sub-folders and files) or a single file into a
# SharePoint folder. Requires an open connection to the destination site.
function Perform-DirectUpload {
    param([hashtable]$CloudDst, [string]$LocalSrc, [switch]$Nested)
    $stats = New-JobStats
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $barId = if ($Nested) { 2 } else { 1 }
    $srcFull = Resolve-LocalPath $LocalSrc
    $leaf = Split-Path -Path $srcFull -Leaf
    $dstRoot = Get-SiteRelativePath -SiteUrl $CloudDst.SiteUrl -ServerRelativePath $CloudDst.FolderPath
    $folderCache = @{}; $mapCache = @{}
    Start-JobUi
    try {
        if (-not $Nested) { Write-Log "UPLOAD STARTED: $(Format-PathWithStorageTag $srcFull) -> $($CloudDst.FolderPath)" 'INFO' }
        else { Write-Log "UPLOAD: $srcFull -> $($CloudDst.FolderPath)" 'INFO' -FileOnly }
        $long = Get-LongPath $srcFull
        if ([System.IO.File]::Exists($long)) {
            Invoke-UploadFile -File (New-Object System.IO.FileInfo($long)) -CloudFolder $dstRoot -RelDisplay $leaf -Stats $stats -FolderCache $folderCache -MapCache $mapCache
        } else {
            $tree = Get-LocalTree -Path $srcFull -Label 'Scanning source'
            if ($tree.Cancelled) { $stats.Cancelled = $true }
            foreach ($e in $tree.Errors) { $stats.Unreadable++; Add-JobFailure -Stats $stats -File $e.File -Path $e.Path -Reason $e.Reason }
            $total = $tree.Files.Count; $i = 0; [long]$seen = 0
            $tick = [System.Diagnostics.Stopwatch]::StartNew()
            if (-not $Nested -and $total -gt 0) { Write-Hint ("Uploading {0} files ({1})..." -f (Format-Count $total), (Format-FileSize $tree.Bytes)) }
            foreach ($f in $tree.Files) {
                if ($stats.Cancelled) { break }
                if ((Check-Interrupt -CurrentFile $f.Name -CurrentFolder (Remove-LongPrefix $f.DirectoryName)) -eq 'CANCEL') { $stats.Cancelled = $true; break }
                $i++; $seen += $f.Length
                if ($tick.ElapsedMilliseconds -ge 500 -or $i -eq $total) {
                    $eta = Get-ETA -Current $seen -Total $tree.Bytes -Elapsed $sw.Elapsed
                    Update-Progress -Id $barId -Activity "Upload: $leaf" -Status ("File {0} of {1}  -  {2} of {3}" -f (Format-Count $i), (Format-Count $total), (Format-FileSize $seen), (Format-FileSize $tree.Bytes)) -Percent ($i * 100.0 / [Math]::Max(1, $total)) -Detail ($(if ($eta) { "About $eta left  |  " } else { '' }) + 'P = pause, C = cancel')
                    $tick.Restart()
                }
                $rel = Get-ChildRelativePath $tree.Root $f.FullName
                $relDir = [System.IO.Path]::GetDirectoryName($rel)
                $cloudFolder = $dstRoot
                if (-not [string]::IsNullOrEmpty($relDir)) {
                    $relCloud = ((($relDir -replace '\\', '/') -split '/') | ForEach-Object { Get-SafeWindowsName $_ }) -join '/'
                    $cloudFolder = Join-CloudPath $dstRoot $relCloud
                }
                Invoke-UploadFile -File $f -CloudFolder $cloudFolder -RelDisplay $rel -Stats $stats -FolderCache $folderCache -MapCache $mapCache
            }
            # Keep empty folders too (folders that received files already exist).
            if (-not $stats.Cancelled) {
                foreach ($d in $tree.Dirs) {
                    $rel = (Get-ChildRelativePath $tree.Root $d.FullName) -replace '\\', '/'
                    $relCloud = (($rel -split '/') | ForEach-Object { Get-SafeWindowsName $_ }) -join '/'
                    $cf = Join-CloudPath $dstRoot $relCloud
                    if (-not $folderCache.ContainsKey($cf) -and $cf.Length -lt $CloudPathLimit) { $folderCache[$cf] = Ensure-CloudFolder -SiteRelativePath $cf }
                }
            }
        }
        Update-Progress -Id $barId -Activity 'Upload' -Completed
    } catch {
        Write-Log "UPLOAD ERROR: $($_.Exception.Message)" 'ERROR'
        Add-JobFailure -Stats $stats -File $leaf -Path $srcFull -Reason $_.Exception.Message
    } finally { $sw.Stop(); Stop-JobUi }

    if ($Nested) {
        Write-Log ("UPLOAD {0}: {1} uploaded, {2} skipped, {3} failed, {4}" -f $leaf, $stats.Success, $stats.Skipped, $stats.Failed, (Format-FileSize $stats.SuccessBytes)) $(if ($stats.Failed) { 'WARN' } else { 'INFO' }) -FileOnly
    } else {
        $tag = Get-PathStorageType $srcFull
        Show-TransferSummary -Title 'UPLOAD' -Stats $stats -Elapsed $sw.Elapsed -Target $leaf -Storage $(if ($tag) { "$tag -> Cloud" } else { '' }) -Color Blue -OkLabel 'Uploaded'
        Register-JobResult -Operation 'Upload' -Source $srcFull -Destination $CloudDst.FolderPath -Stats $stats -Elapsed $sw.Elapsed
    }
    return $stats
}

# Option 16: pick items from a local folder and upload them. Requires a connection.
function Perform-LocalToCloudBatchUpload {
    param([string]$LocalSrc, [hashtable]$CloudDst)
    $srcFull = Resolve-LocalPath $LocalSrc
    while ($true) {
        try { Connect-CloudSite -SiteUrl $CloudDst.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; return }
        try {
            $di = New-Object System.IO.DirectoryInfo((Get-LongPath $srcFull))
            $items = @($di.GetFileSystemInfos() | Sort-Object @{ Expression = { -not ($_ -is [System.IO.DirectoryInfo]) } }, @{ Expression = { $_.Name } })
        } catch { Write-Bad "Could not read '$srcFull': $($_.Exception.Message)"; return }
        if ($items.Count -eq 0) { Write-Caution 'This folder is empty.'; return }
        Write-Panel -Title 'BATCH UPLOAD' -Color Blue -Lines @(
            @{ Label = 'From (this PC)'; Value = $srcFull }
            @{ Label = 'To (cloud)';     Value = $CloudDst.FolderPath }
        )
        Write-Hint 'Measuring folder sizes...'
        $rows = Get-LocalIndexRows -Items $items
        Show-ItemIndex -Title "CONTENTS OF $(Split-Path -Path $srcFull -Leaf)" -Rows $rows -Color Blue -InfoHeader 'SIZE | FILES'
        $sel = Read-IndexSelection -Count $rows.Count -Verb 'upload'
        if ($sel.Action -eq 'Quit') { return }
        if ($sel.Action -eq 'Refresh') { continue }
        if ($sel.Action -eq 'Search') { Invoke-LocalSearchLoop -SearchPath $srcFull; continue }
        $chosen = @($sel.Indices | ForEach-Object { $rows[$_] })
        $batch = New-JobStats
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $dstRoot = Get-SiteRelativePath -SiteUrl $CloudDst.SiteUrl -ServerRelativePath $CloudDst.FolderPath
        $folderCache = @{}; $mapCache = @{}
        Start-JobUi
        try {
            $n = 0
            foreach ($c in $chosen) {
                if ($script:JobCancelled) { $batch.Cancelled = $true; break }
                $n++
                Update-Progress -Id 1 -Activity 'Batch upload' -Status ("Item {0} of {1}: {2}" -f $n, $chosen.Count, $c.Name) -Percent (($n - 1) * 100.0 / $chosen.Count) -Detail 'P = pause, C = cancel'
                Write-BatchItemHeader -Index $n -Total $chosen.Count -Name $c.Name -Verb 'UPLOAD' -Color Blue
                if ($c.Kind -eq 'FOLDER') {
                    $sub = @{ SiteUrl = $CloudDst.SiteUrl; FolderPath = (Join-CloudPath $CloudDst.FolderPath (Get-SafeWindowsName $c.Name)) }
                    $res = Perform-DirectUpload -CloudDst $sub -LocalSrc $c.Item.FullName -Nested
                    Merge-JobStats -Into $batch -From $res
                    Write-Host ("  {0} uploaded, {1} skipped, {2} failed" -f $res.Success, $res.Skipped, $res.Failed) -ForegroundColor $(if ($res.Failed) { 'Yellow' } else { 'DarkGray' })
                } else {
                    Invoke-UploadFile -File $c.Item -CloudFolder $dstRoot -RelDisplay $c.Name -Stats $batch -FolderCache $folderCache -MapCache $mapCache
                }
                if ($script:JobCancelled) { $batch.Cancelled = $true }
            }
        } finally { $sw.Stop(); Stop-JobUi }
        $tag = Get-PathStorageType $srcFull
        Show-TransferSummary -Title 'BATCH UPLOAD' -Stats $batch -Elapsed $sw.Elapsed -Storage $(if ($tag) { "$tag -> Cloud" } else { '' }) -Color Blue -OkLabel 'Uploaded'
        Register-JobResult -Operation 'Batch upload' -Source $srcFull -Destination $CloudDst.FolderPath -Stats $batch -Elapsed $sw.Elapsed
        Show-Pause -Message 'Press any key to return to the folder list...'
    }
}

# ======================================================================================
# SECTION 13 - ARCHIVE ENGINES (download -> verify on disk -> recycle in SharePoint)
# --------------------------------------------------------------------------------------
# Safety rules (all must hold before anything is deleted from SharePoint):
#   * the download finished, was not cancelled and had no failures
#   * every cloud folder could be listed (during the download AND the verification)
#   * a fresh cloud inventory matches the files ON DISK: every cloud file has a local copy
#     with the same size (Office files with SharePoint metadata differences are accepted
#     and reported), and no two cloud files were saved to the same local file
#   * the job is still not cancelled at the moment of deletion
# Deleted items always go to the SharePoint Recycle Bin (recoverable for ~93 days).
# ======================================================================================

function Test-ArchiveResult {
    param([hashtable]$CloudData, [hashtable]$DlStats, [switch]$Quiet)
    $v = @{
        Passed = $false; CloudFiles = 0; CloudBytes = [long]0; LocalFiles = 0; LocalBytes = [long]0
        Missing = 0; WrongSize = 0; Accepted = 0; IsFile = $false
        Reasons = (New-Object System.Collections.Generic.List[string])
        Problems = (New-Object System.Collections.Generic.List[string])
    }
    if ($DlStats.Cancelled)          { $v.Reasons.Add('The download was cancelled.'); return $v }
    if ($DlStats.Failed -gt 0)       { $v.Reasons.Add("$($DlStats.Failed) item(s) failed to download.") }
    if ($DlStats.ListingErrors -gt 0){ $v.Reasons.Add("$($DlStats.ListingErrors) cloud folder(s) could not be listed during the download.") }
    $inv = Get-CloudInventory -SiteUrl $CloudData.SiteUrl -FolderPath $CloudData.FolderPath -Tag 'Re-reading cloud for verification' -Quiet:$Quiet
    $v.IsFile = $inv.IsFile
    if ($inv.Cancelled) { $v.Reasons.Add('Verification was cancelled.'); return $v }
    if ($inv.Errors.Count -gt 0) { $v.Reasons.Add("$($inv.Errors.Count) cloud folder(s) could not be listed during verification.") }
    $byRel = @{}
    $seenLocal = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $DlStats.Manifest) {
        $byRel[$m.CloudRel.ToLowerInvariant()] = $m
        if (-not $seenLocal.Add($m.Local)) { $v.Reasons.Add("Two cloud files were saved to the same local file: $(Remove-LongPrefix $m.Local)") }
    }
    foreach ($f in $inv.Files) {
        $v.CloudFiles++; $v.CloudBytes += [long]$f.Length
        $m = $byRel[$f.RelPath.ToLowerInvariant()]
        if (-not $m -or -not [System.IO.File]::Exists($m.Local)) {
            $v.Missing++
            if ($v.Problems.Count -lt 200) { $v.Problems.Add("MISSING ON DISK: $($f.RelPath)") }
            continue
        }
        $len = (New-Object System.IO.FileInfo($m.Local)).Length
        $v.LocalFiles++; $v.LocalBytes += $len
        if ($len -ne [long]$f.Length) {
            if (Test-AcceptableSizeDifference -Name $f.Name -CloudSize ([long]$f.Length) -LocalSize $len -LocalPath $m.Local) { $v.Accepted++ }
            else {
                $v.WrongSize++
                if ($v.Problems.Count -lt 200) { $v.Problems.Add(("WRONG SIZE: {0} (cloud {1} bytes, disk {2} bytes)" -f $f.RelPath, $f.Length, $len)) }
            }
        }
    }
    if ($v.Missing -gt 0)   { $v.Reasons.Add("$($v.Missing) cloud file(s) are not on disk.") }
    if ($v.WrongSize -gt 0) { $v.Reasons.Add("$($v.WrongSize) file(s) on disk have a different size from the cloud copy.") }
    if ($v.Problems.Count -gt 0) { Write-LogBlock -Lines $v.Problems.ToArray() -Level 'ERROR' }
    $v.TopNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $inv.Files) { $null = $v.TopNames.Add(($f.RelPath -split '/')[0]) }
    foreach ($d in $inv.Folders) { $null = $v.TopNames.Add(($d -split '/')[0]) }
    $v.Passed = ($v.Reasons.Count -eq 0)
    return $v
}

function Show-ArchiveVerifyPanel {
    param([hashtable]$V, [string]$Title = 'ARCHIVE VERIFICATION')
    $rows = @(
        @{ Label = 'Cloud';        Value = ("{0} files  |  {1}" -f (Format-Count $V.CloudFiles), (Format-FileSize $V.CloudBytes)) }
        @{ Label = 'On this disk'; Value = ("{0} files  |  {1}" -f (Format-Count $V.LocalFiles), (Format-FileSize $V.LocalBytes)) }
        @{ Label = 'Missing on disk'; Value = (Format-Count $V.Missing); Level = $(if ($V.Missing) { 'ERROR' } else { 'SUCCESS' }) }
        @{ Label = 'Wrong size';      Value = (Format-Count $V.WrongSize); Level = $(if ($V.WrongSize) { 'ERROR' } else { 'SUCCESS' }) }
    )
    if ($V.Accepted) { $rows += @{ Label = 'Office metadata differences'; Value = "$($V.Accepted) (accepted)"; Level = 'WARN' } }
    if ($V.Passed) { $rows += @{ Label = 'Result'; Value = 'PASSED - every cloud file is safely on disk'; Level = 'SUCCESS' } }
    else {
        $rows += @{ Label = 'Result'; Value = 'FAILED - the cloud copy will NOT be deleted'; Level = 'ERROR' }
        foreach ($r in $V.Reasons) { $rows += @{ Text = "- $r"; Color = 'Red' } }
    }
    Write-SummaryBox -Title $Title -Lines $rows -BoxColor $(if ($V.Passed) { 'Green' } else { 'Red' })
}

# Recycles a verified archive source. Library roots are never removed - their top-level
# items are recycled one by one instead. Returns @{ Deleted; Failed; CheckedOut; Errors }.
function Remove-ArchiveSource {
    param([hashtable]$CloudData, [bool]$IsFile, $VerifiedTopNames = $null)
    $out = @{ Deleted = 0; Failed = 0; CheckedOut = (New-Object System.Collections.Generic.List[string]); Errors = (New-Object System.Collections.Generic.List[string]) }
    $rel = Get-SiteRelativePath -SiteUrl $CloudData.SiteUrl -ServerRelativePath $CloudData.FolderPath
    $targets = @()
    if ($IsFile) { $targets = @(@{ Srv = $CloudData.FolderPath; Folder = $false; Name = ($CloudData.FolderPath -split '/')[-1] }) }
    elseif (Test-IsLibraryRootPath -SiteUrl $CloudData.SiteUrl -SiteRelativePath $rel) {
        Write-Hint 'This is a library root: the library is kept and the verified items inside it are recycled.'
        foreach ($c in (Get-CloudFolderItems -SiteRelativeUrl $rel)) {
            if ($null -ne $VerifiedTopNames -and -not $VerifiedTopNames.Contains($c.Name)) {
                Write-Caution "'$($c.Name)' appeared after the verification - it was not archived, so it is left in place."
                continue
            }
            $targets += @{ Srv = (Get-CloudProp $c 'ServerRelativeUrl' (Join-CloudPath $CloudData.FolderPath $c.Name)); Folder = (Test-IsCloudFolder $c); Name = $c.Name }
        }
    } else { $targets = @(@{ Srv = $CloudData.FolderPath; Folder = $true; Name = ($CloudData.FolderPath -split '/')[-1] }) }
    foreach ($t in $targets) {
        if ($script:JobCancelled) { $out.Errors.Add('Cancelled before every item was deleted.'); break }
        try {
            Remove-CloudItem -SiteUrl $CloudData.SiteUrl -ServerRelativeUrl $t.Srv -IsFolder:$t.Folder
            $out.Deleted++
            Write-Log "Recycled (archived): $($t.Srv)" 'SUCCESS' -Detail
        } catch {
            $msg = $_.Exception.Message
            $out.Failed++
            if ($msg -match '(?i)checked out') { $co = Get-CheckoutInfo $msg; $out.CheckedOut.Add("$($t.Name): '$($co.File)' is checked out by $($co.User)") }
            else { $out.Errors.Add("$($t.Name): $msg") }
            Write-Log "DELETE FAILED (archive): $($t.Srv) - $msg" 'ERROR'
        }
    }
    return $out
}

function Show-CheckoutHelp {
    param([object]$Items)
    $list = @($Items)
    if ($list.Count -eq 0) { return }
    $lines = @($list | ForEach-Object { "- $_" }) + @('-', 'These were downloaded and verified but could not be recycled. Ask the owner to check the files in (or discard the checkout in the library), then run the archive again - files already on disk are skipped and only the delete is retried.')
    Write-Panel -Title 'DELETION BLOCKED BY CHECKED-OUT FILES' -Lines $lines -Color Yellow
}

# Archive one cloud folder or file: download -> verify on disk -> recycle.
# Returns @{ Outcome; Files; Bytes; Deleted; Verified }.
function Invoke-ArchiveOne {
    param([hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest, [switch]$Compact)
    $out = @{ Outcome = 'NOT STARTED'; Files = 0; Bytes = [long]0; Deleted = $false; Verified = $false; Level = 'ERROR'; Stage = 'Download' }
    if ($Compact) { Write-Step '1/3' 'Downloading...' } else { Write-Rule -Title 'STEP 1 OF 3 - DOWNLOAD' -Color Magenta }
    $dl = Perform-DirectDownload -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest -Manifest -NoRegister -Nested:$Compact -Title 'ARCHIVE DOWNLOAD'
    if ($Compact) { Write-Host ("        {0} downloaded, {1} already present, {2} failed" -f $dl.Success, $dl.Skipped, $dl.Failed) -ForegroundColor DarkGray }
    if ($dl.Cancelled -or $script:JobCancelled) { $out.Outcome = 'CANCELLED - cloud untouched'; $out.Level = 'WARN'; $out.Stage = 'Cancelled'; return $out }
    $out.Stage = if ($dl.Failed -gt 0 -or $dl.ListingErrors -gt 0) { 'Download' } else { 'Verify' }
    if ($Compact) { Write-Step '2/3' 'Verifying the files on disk against SharePoint...' -Color Yellow } else { Write-Rule -Title 'STEP 2 OF 3 - VERIFY' -Color Yellow }
    $v = Test-ArchiveResult -CloudData $CloudData -DlStats $dl -Quiet:$Compact
    $out.Files = $v.LocalFiles; $out.Bytes = $v.LocalBytes
    if ($Compact) {
        if ($v.Passed) { Write-Ok ("Verified: {0} files, {1}" -f $v.LocalFiles, (Format-FileSize $v.LocalBytes)) }
        else { foreach ($r in $v.Reasons) { Write-Bad $r } }
    } else { Show-ArchiveVerifyPanel -V $v }
    if (-not $v.Passed) { $out.Outcome = 'NOT DELETED - ' + ($v.Reasons -join ' '); return $out }
    $out.Verified = $true
    if ($script:JobCancelled) { $out.Outcome = 'CANCELLED before delete - cloud untouched'; $out.Level = 'WARN'; $out.Stage = 'Cancelled'; return $out }
    $out.Stage = 'Delete'
    if ($Compact) { Write-Step '3/3' 'Recycling the cloud copy...' -Color Cyan } else { Write-Rule -Title 'STEP 3 OF 3 - RECYCLE CLOUD COPY' -Color Cyan }
    try { $del = Remove-ArchiveSource -CloudData $CloudData -IsFile $v.IsFile -VerifiedTopNames $v.TopNames }
    catch { $del = @{ Deleted = 0; Failed = 1; CheckedOut = @(); Errors = @($_.Exception.Message) } }
    Show-CheckoutHelp -Items $del.CheckedOut
    foreach ($e in $del.Errors) { Write-Bad $e }
    if ($del.Failed -eq 0 -and @($del.Errors).Count -eq 0) {
        $out.Deleted = $true; $out.Level = 'SUCCESS'; $out.Stage = 'Done'
        $out.Outcome = ("ARCHIVED + RECYCLED - {0} files, {1}" -f $v.LocalFiles, (Format-FileSize $v.LocalBytes))
        Write-Ok 'Cloud copy sent to the SharePoint Recycle Bin.'
    } else {
        $out.Level = 'WARN'
        $out.Outcome = ("ARCHIVED, but {0} item(s) could not be recycled" -f $del.Failed)
    }
    return $out
}

# Option 17.
function Perform-ArchiveDownload {
    param([hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "ARCHIVE STARTED: $($CloudData.FolderPath) -> $LocalDest" 'INFO'
    Start-JobUi
    try { $res = Invoke-ArchiveOne -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest }
    finally { $sw.Stop(); Stop-JobUi }
    $rows = @(
        @{ Label = 'Cloud source'; Value = $CloudData.FolderPath }
        @{ Label = 'Saved to';     Value = $LocalDest }
        @{ Label = 'Outcome';      Value = $res.Outcome; Level = $res.Level }
        @{ Label = 'Time taken';   Value = (Format-Elapsed $sw.Elapsed) }
    )
    if ($res.Deleted) { $rows += @{ Label = 'Recycle Bin'; Value = (Get-SiteRecycleBinUrl $CloudData.SiteUrl) } }
    Write-SummaryBox -Title 'ARCHIVE SUMMARY' -Lines $rows -BoxColor $(switch ($res.Level) { 'SUCCESS' { 'Green' } 'WARN' { 'Yellow' } default { 'Red' } })
    $st = New-JobStats; $st.Success = $res.Files; $st.SuccessBytes = $res.Bytes
    if ($res.Level -eq 'ERROR') { $st.Failed = 1 }
    Register-JobResult -Operation 'Archive' -Source $CloudData.FolderPath -Destination $LocalDest -Stats $st -Elapsed $sw.Elapsed -Outcome $res.Outcome
}

# Option 18 (interactive) and the per-library engine of the leaver options (-AutoSelectAll).
# Each picked item is downloaded, verified on disk and recycled on its own.
# Returns @{ Archived; Deleted; DlFailed; VerifyFailed; DeleteFailed; Bytes; Cancelled; CheckedOut }.
function Perform-BatchArchiveDownload {
    param([hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest, [switch]$AutoSelectAll, $UsedNames = $null)
    $totals = @{ Archived = 0; Deleted = 0; DlFailed = 0; VerifyFailed = 0; DeleteFailed = 0; Bytes = [long]0; Cancelled = $false; Items = 0; CheckedOut = (New-Object System.Collections.Generic.List[string]); Rows = (New-Object System.Collections.Generic.List[hashtable]) }
    $base = Get-SiteRelativePath -SiteUrl $CloudData.SiteUrl -ServerRelativePath $CloudData.FolderPath
    if (-not $AutoSelectAll -and (Get-CloudFileIfFile -SiteRelativeUrl $base)) {
        Write-Info 'The address points to a single file.'
        if (Confirm-Start -Title 'READY TO ARCHIVE THIS FILE' -Color Red -Danger -DangerText 'Download, verify and then recycle the cloud copy?' -Rows @(@{ Label = 'Cloud file'; Value = $CloudData.FolderPath }, @{ Label = 'Save to'; Value = $LocalDest })) {
            Perform-ArchiveDownload -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest
        } else { Write-Hint 'Nothing was changed.' }
        return $totals
    }
    while ($true) {
        try { $items = Get-CloudIndexItems -SiteRelativeUrl $base }
        catch {
            Write-Bad "The folder could not be listed: $($_.Exception.Message)"
            $totals.DlFailed++
            $totals.Rows.Add(@{ Label = ($CloudData.FolderPath -split '/')[-1]; Status = "NOT LISTED - $($_.Exception.Message)"; Level = 'ERROR' })
            return $totals
        }
        if ($items.Count -eq 0) {
            if (-not $AutoSelectAll) { Write-Caution 'This folder is empty - nothing to archive.' } else { Write-Hint 'Library is empty - nothing to archive.' }
            return $totals
        }
        if ($AutoSelectAll) {
            $chosen = $items
            Write-Hint ("Archiving all {0} item(s) in /{1}" -f $items.Count, $base)
        } else {
            Write-Panel -Title 'BATCH ARCHIVE' -Color Yellow -Lines @(
                @{ Label = 'From (cloud)'; Value = $CloudData.FolderPath }
                @{ Label = 'To (this PC)'; Value = $LocalDest }
                'Each picked item is downloaded, verified on disk and only then sent to the SharePoint Recycle Bin.'
            )
            $rows = ConvertTo-CloudIndexRows -Items $items
            Show-ItemIndex -Title "CONTENTS OF $(($CloudData.FolderPath -split '/')[-1])" -Rows $rows -Color Yellow
            $sel = Read-IndexSelection -Count $rows.Count -Verb 'archive'
            if ($sel.Action -eq 'Quit') { return $totals }
            if ($sel.Action -eq 'Refresh') { continue }
            if ($sel.Action -eq 'Search') { Invoke-CloudSearchLoop -CloudData $CloudData; continue }
            $chosen = @($sel.Indices | ForEach-Object { $items[$_] })
            if (-not (Read-Confirmation -Message ("Archive {0} item(s)? Each one is downloaded, verified, then recycled in SharePoint (restorable for ~93 days)." -f $chosen.Count))) { Write-Hint 'Nothing was changed.'; continue }
        }
        $nameMap = Get-LocalNameMap -Items $items -Used $UsedNames
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Start-JobUi
        try {
            $rootLong = Get-LongPath (Resolve-LocalPath $LocalDest)
            if (-not [System.IO.Directory]::Exists($rootLong)) { $null = [System.IO.Directory]::CreateDirectory($rootLong) }
            $n = 0
            foreach ($it in $chosen) {
                if ((Check-Interrupt -CurrentFile $it.Name -CurrentFolder $CloudData.FolderPath) -eq 'CANCEL') { $totals.Cancelled = $true; break }
                $n++; $totals.Items++
                Update-Progress -Id 1 -Activity "Archiving /$base" -Status ("Item {0} of {1}: {2}" -f $n, $chosen.Count, $it.Name) -Percent (($n - 1) * 100.0 / $chosen.Count) -Detail 'P = pause, C = cancel'
                Write-BatchItemHeader -Index $n -Total $chosen.Count -Name $it.Name -Verb 'ARCHIVE' -Color Yellow
                $srv = Get-CloudProp $it 'ServerRelativeUrl' (Join-CloudPath $CloudData.FolderPath $it.Name)
                $local = Join-LocalPath (Resolve-LocalPath $LocalDest) $nameMap[$it.Name]
                $itemData = @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $srv }
                if (Test-IsCloudFolder $it) {
                    $res = Invoke-ArchiveOne -CloudData $itemData -LocalDest $local -BufferDest $BufferDest -Compact
                } else {
                    # A single file: download it into the destination folder, then verify / recycle.
                    Write-Step '1/3' 'Downloading...'
                    $dl = New-JobStats
                    Invoke-DownloadFile -Item $it -ServerRelativeUrl $srv -CloudRel $it.Name -LocalPath $local -BufferDest $BufferDest -Stats $dl -Manifest
                    $res = @{ Outcome = ''; Files = 0; Bytes = [long]0; Deleted = $false; Verified = $false; Level = 'ERROR'; Stage = 'Download' }
                    if ($dl.Cancelled -or $script:JobCancelled) { $res.Outcome = 'CANCELLED - cloud untouched'; $res.Level = 'WARN'; $res.Stage = 'Cancelled' }
                    else {
                        Write-Step '2/3' 'Verifying...' -Color Yellow
                        $v = Test-ArchiveResult -CloudData $itemData -DlStats $dl -Quiet
                        $res.Files = $v.LocalFiles; $res.Bytes = $v.LocalBytes
                        if (-not $v.Passed) { foreach ($r in $v.Reasons) { Write-Bad $r }; $res.Outcome = 'NOT DELETED - ' + ($v.Reasons -join ' '); if ($dl.Failed -eq 0) { $res.Stage = 'Verify' } }
                        elseif ($script:JobCancelled) { $res.Verified = $true; $res.Outcome = 'CANCELLED before delete - cloud untouched'; $res.Level = 'WARN'; $res.Stage = 'Cancelled' }
                        else {
                            $res.Verified = $true; $res.Stage = 'Delete'
                            Write-Ok ("Verified ({0})" -f (Format-FileSize $v.LocalBytes))
                            Write-Step '3/3' 'Recycling the cloud copy...' -Color Cyan
                            try {
                                Remove-CloudItem -SiteUrl $CloudData.SiteUrl -ServerRelativeUrl $srv
                                $res.Deleted = $true; $res.Level = 'SUCCESS'; $res.Stage = 'Done'; $res.Outcome = "ARCHIVED + RECYCLED - $(Format-FileSize $v.LocalBytes)"
                                Write-Ok 'Sent to the SharePoint Recycle Bin.'
                            } catch {
                                $msg = $_.Exception.Message
                                if ($msg -match '(?i)checked out') { $co = Get-CheckoutInfo $msg; $totals.CheckedOut.Add("$($it.Name): checked out by $($co.User)") }
                                Write-Bad "Could not recycle: $msg"
                                $res.Level = 'WARN'; $res.Outcome = "ARCHIVED, but could not be recycled: $msg"
                            }
                        }
                    }
                }
                $totals.Bytes += $res.Bytes
                if ($res.Verified) { $totals.Archived++ }
                switch ($res.Stage) {
                    'Done'     { $totals.Deleted++ }
                    'Download' { $totals.DlFailed++ }
                    'Verify'   { $totals.VerifyFailed++ }
                    'Delete'   { $totals.DeleteFailed++ }
                }
                $totals.Rows.Add(@{ Label = $it.Name; Status = $res.Outcome; Level = $res.Level })
                if ($script:JobCancelled) { $totals.Cancelled = $true; break }
            }
        } finally { $sw.Stop(); Stop-JobUi }

        $rows = @(
            @{ Label = 'Archived (verified)'; Value = ("{0} item(s)  |  {1}" -f $totals.Archived, (Format-FileSize $totals.Bytes)); Level = 'SUCCESS' }
            @{ Label = 'Recycled in SharePoint'; Value = "$($totals.Deleted)"; Level = $(if ($totals.Deleted -eq $totals.Archived) { 'SUCCESS' } else { 'WARN' }) }
            @{ Label = 'Download problems'; Value = "$($totals.DlFailed)"; Level = $(if ($totals.DlFailed) { 'ERROR' } else { 'SUCCESS' }) }
            @{ Label = 'Failed verification'; Value = "$($totals.VerifyFailed)"; Level = $(if ($totals.VerifyFailed) { 'ERROR' } else { 'SUCCESS' }) }
            @{ Label = 'Could not recycle'; Value = "$($totals.DeleteFailed)"; Level = $(if ($totals.DeleteFailed) { 'WARN' } else { 'SUCCESS' }) }
            @{ Label = 'Time taken'; Value = (Format-Elapsed $sw.Elapsed) }
        )
        if ($totals.Cancelled) { $rows += @{ Label = 'Stopped'; Value = 'Cancelled by you - remaining items were not touched'; Level = 'WARN' } }
        Write-SummaryBox -Title ("BATCH ARCHIVE - /{0}" -f $base) -Lines $rows -BoxColor $(if ($totals.DlFailed -or $totals.VerifyFailed) { 'Red' } elseif ($totals.DeleteFailed -or $totals.Cancelled) { 'Yellow' } else { 'Green' })
        Show-CheckoutHelp -Items $totals.CheckedOut
        if ($AutoSelectAll) { return $totals }
        $st = New-JobStats; $st.Success = $totals.Archived; $st.SuccessBytes = $totals.Bytes; $st.Failed = $totals.DlFailed + $totals.VerifyFailed + $totals.DeleteFailed; $st.Cancelled = $totals.Cancelled
        Register-JobResult -Operation 'Batch archive' -Source $CloudData.FolderPath -Destination $LocalDest -Stats $st -Elapsed $sw.Elapsed
        if ($totals.Deleted -gt 0) { Write-Hint "Recycle Bin: $(Get-SiteRecycleBinUrl $CloudData.SiteUrl)" }
        Show-Pause -Message 'Press any key to return to the folder list...'
        $totals = @{ Archived = 0; Deleted = 0; DlFailed = 0; VerifyFailed = 0; DeleteFailed = 0; Bytes = [long]0; Cancelled = $false; Items = 0; CheckedOut = (New-Object System.Collections.Generic.List[string]); Rows = (New-Object System.Collections.Generic.List[hashtable]) }
    }
}

# Option 19: archive several cloud sources (each into its own sub-folder), unattended.
function Perform-MultiSourceArchive {
    param([object[]]$Sources, [string]$LocalDest, [string]$BufferDest)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[hashtable]
    $ok = 0; $bad = 0; $bytes = [long]0
    Write-Log "MULTI-SOURCE ARCHIVE STARTED: $($Sources.Count) source(s) -> $LocalDest" 'INFO'
    Start-JobUi
    try {
        for ($i = 0; $i -lt $Sources.Count; $i++) {
            $src = $Sources[$i]
            if ((Check-Interrupt) -eq 'CANCEL') { for ($j = $i; $j -lt $Sources.Count; $j++) { $rows.Add(@{ Label = $Sources[$j].Label; Status = 'NOT STARTED (cancelled)'; Level = 'WARN' }) }; break }
            Update-Progress -Id 1 -Activity 'Multi-source archive' -Status ("Source {0} of {1}: {2}" -f ($i + 1), $Sources.Count, $src.Label) -Percent ($i * 100.0 / $Sources.Count) -Detail 'P = pause, C = cancel'
            Write-Host ""
            Write-Rule -Title ("SOURCE {0} OF {1}: {2}" -f ($i + 1), $Sources.Count, $src.Label) -Color Yellow -Double
            Write-Hint "Cloud : $($src.CloudData.FolderPath)"
            $dest = Join-LocalPath (Resolve-LocalPath $LocalDest) $src.SubFolder
            Write-Hint "Saving: $dest"
            try { Connect-CloudSite -SiteUrl $src.CloudData.SiteUrl }
            catch { $rows.Add(@{ Label = $src.Label; Status = "CONNECTION FAILED - $($_.Exception.Message)"; Level = 'ERROR' }); $bad++; continue }
            $res = Invoke-ArchiveOne -CloudData $src.CloudData -LocalDest $dest -BufferDest $BufferDest -Compact
            $rows.Add(@{ Label = $src.Label; Status = $res.Outcome; Level = $res.Level })
            if ($res.Deleted) { $ok++; $bytes += $res.Bytes } else { $bad++ }
            if ($script:JobCancelled) { for ($j = $i + 1; $j -lt $Sources.Count; $j++) { $rows.Add(@{ Label = $Sources[$j].Label; Status = 'NOT STARTED (cancelled)'; Level = 'WARN' }) }; break }
        }
    } finally { $sw.Stop(); Stop-JobUi }
    Show-OutcomeTable -Title 'MULTI-SOURCE ARCHIVE - RESULTS' -Rows $rows.ToArray() -Summary @(
        @{ Label = 'Archived + recycled'; Value = ("{0} of {1}  |  {2}" -f $ok, $Sources.Count, (Format-FileSize $bytes)); Level = $(if ($ok -eq $Sources.Count) { 'SUCCESS' } else { 'WARN' }) }
        @{ Label = 'Need attention'; Value = "$bad"; Level = $(if ($bad) { 'ERROR' } else { 'SUCCESS' }) }
        @{ Label = 'Time taken'; Value = (Format-Elapsed $sw.Elapsed) }
    )
    $st = New-JobStats; $st.Success = $ok; $st.Failed = $bad; $st.SuccessBytes = $bytes; $st.Cancelled = $script:JobCancelled
    Register-JobResult -Operation 'Multi-source archive' -Source "$($Sources.Count) sources" -Destination $LocalDest -Stats $st -Elapsed $sw.Elapsed
}

function Show-OutcomeTable {
    param([string]$Title, [object[]]$Rows, [object[]]$Summary = @())
    $lines = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Summary) { $lines.Add($s) }
    if ($Summary.Count -gt 0 -and $Rows.Count -gt 0) { $lines.Add('-') }
    foreach ($r in $Rows) { $lines.Add(@{ Label = $r.Label; Value = $r.Status; Level = $r.Level }) }
    Write-SummaryBox -Title $Title -Lines $lines.ToArray() -BoxColor Cyan
}

# ======================================================================================
# SECTION 14 - LEAVER ONEDRIVE TOOLS
# ======================================================================================

function ConvertTo-TenantName {
    param([string]$Text)
    $t = "$Text".Trim().ToLowerInvariant() -replace '^https?://', ''
    $t = ($t -split '/')[0]
    $t = $t -replace '\.sharepoint\.com$', '' -replace '\.onmicrosoft\.com$', ''
    $t = $t -replace '-(my|admin)$', ''
    return $t
}

# john.smith@contoso.com -> https://contoso-my.sharepoint.com/personal/john_smith_contoso_com
# (every character other than a-z, 0-9 and _ becomes "_", as SharePoint does).
function ConvertTo-OneDriveSiteUrl {
    param([string]$TenantName, [string]$Upn)
    $safe = ($Upn.Trim().ToLowerInvariant()) -replace '[^a-z0-9_]', '_'
    return "https://$(ConvertTo-TenantName $TenantName)-my.sharepoint.com/personal/$safe"
}

function Get-LeaverRecycleBinUrl {
    param([string]$LeaverUrl)
    $m = [regex]::Match($LeaverUrl, '(?i)(https://[^/]+/personal/[^/?# ]+)')
    if ($m.Success) { return "$($m.Groups[1].Value.TrimEnd('/'))/_layouts/15/RecycleBin.aspx" }
    return ''
}

# Shows recycle-bin links and offers to copy / open them.
function Write-RecycleBinLinks {
    param([object[]]$Links)
    $list = @($Links | Where-Object { $_ -and $_.Url })
    if ($list.Count -eq 0) { return }
    $lines = New-Object System.Collections.Generic.List[object]
    $lines.Add('Deleted items stay in each Recycle Bin for about 93 days. Open a link to review or empty it (select all > Delete).')
    $lines.Add('-')
    foreach ($l in $list) {
        if ($list.Count -gt 1 -and $l.Label) { $lines.Add(@{ Text = $l.Label; Color = 'White' }) }
        $lines.Add(@{ Text = $l.Url; Color = 'Cyan' })
    }
    Write-Panel -Title ("RECYCLE BIN LINK{0}" -f $(if ($list.Count -gt 1) { 'S' } else { '' })) -Lines $lines.ToArray() -Color Cyan
    Write-LogBlock -Lines @($list | ForEach-Object { "RECYCLE BIN: $($_.Url)" })
    $a = Read-Choice -Prompt 'C = copy to clipboard   O = open in browser   B = both   ENTER = skip' -Valid @('C', 'O', 'B')
    if ($a -in @('C', 'B')) { if (Set-ClipboardText (($list | ForEach-Object { $_.Url }) -join "`r`n")) { Write-Ok "Copied $($list.Count) link(s)." } }
    if ($a -in @('O', 'B')) {
        foreach ($l in $list) { try { Start-Process $l.Url -ErrorAction Stop; if ($list.Count -gt 1) { Start-Sleep -Milliseconds 800 } } catch { Write-Bad "Could not open $($l.Url)" } }
        Write-Ok 'Opened - check your browser.'
    }
}

# Archives every user document library of a connected personal (OneDrive) site.
# The Documents library goes straight into -LocalDest; other libraries get their own
# sub-folder (renamed if that name is already used by a Documents folder).
function Perform-LeaverOneDriveArchive {
    param([string]$LeaverSiteUrl, [string]$LocalDest, [string]$BufferDest)
    $out = @{ Libraries = 0; LibOk = 0; LibProblems = 0; Archived = 0; Deleted = 0; Bytes = [long]0; Cancelled = $false; Rows = (New-Object System.Collections.Generic.List[hashtable]); Error = '' }
    $system = @('Form Templates', 'Site Pages', 'Style Library', 'Converted Forms', 'Preservation Hold Library', 'Master Page Gallery', 'Composed Looks', 'Theme Gallery', 'Site Collection Documents', 'Site Collection Images', 'Site Assets')
    try {
        $lists = Invoke-WithRetry -OperationName 'Read libraries' -Quiet -ScriptBlock { Get-PnPList -ErrorAction Stop }
    } catch { $out.Error = "Libraries could not be read: $($_.Exception.Message)"; Write-Bad $out.Error; return $out }
    $libs = @($lists | Where-Object { $_.BaseType -eq 1 -and -not $_.Hidden -and $system -notcontains $_.Title } |
        Sort-Object @{ Expression = { $_.Title -ne 'Documents' } }, @{ Expression = { $_.Title } })
    $out.Libraries = $libs.Count
    if ($libs.Count -eq 0) { Write-Caution 'No document libraries found - the OneDrive may already be empty.'; return $out }
    $rows = @($libs | ForEach-Object { @{ Label = $_.Title; Value = ("{0:N0} items" -f [long](Get-CloudProp $_ 'ItemCount' 0)) } })
    Write-Panel -Title 'DOCUMENT LIBRARIES' -Lines $rows -Color DarkYellow
    $localRoot = Resolve-LocalPath $LocalDest
    foreach ($lib in $libs) {
        if ((Check-Interrupt) -eq 'CANCEL' -or $script:JobCancelled) { $out.Cancelled = $true; $out.Rows.Add(@{ Label = $lib.Title; Status = 'NOT STARTED (cancelled)'; Level = 'WARN' }); continue }
        $root = Get-CloudProp $lib 'RootFolder'
        $libUrl = Get-CloudProp $root 'ServerRelativeUrl'
        if (-not $libUrl) { try { $libUrl = (Get-PnPProperty -ClientObject $lib -Property RootFolder).ServerRelativeUrl } catch { } }
        if (-not $libUrl) { $out.LibProblems++; $out.Rows.Add(@{ Label = $lib.Title; Status = 'Library address could not be read'; Level = 'ERROR' }); continue }
        # Documents goes straight into the destination. Other libraries get a folder whose
        # name cannot clash with a Documents folder - on this run or any later re-run.
        if ($lib.Title -eq 'Documents') { $libLocal = $localRoot }
        else { $libLocal = Join-LocalPath $localRoot ((Get-SafeWindowsName $lib.Title) + ' (library)') }
        Write-Host ""
        Write-Rule -Title ("LIBRARY: {0}  ->  {1}" -f $lib.Title, $libLocal) -Color DarkYellow -Double
        try { if (-not [System.IO.Directory]::Exists((Get-LongPath $libLocal))) { $null = [System.IO.Directory]::CreateDirectory((Get-LongPath $libLocal)) } } catch { }
        $res = Perform-BatchArchiveDownload -CloudData @{ SiteUrl = $LeaverSiteUrl; FolderPath = $libUrl } -LocalDest $libLocal -BufferDest $BufferDest -AutoSelectAll
        $out.Archived += $res.Archived; $out.Deleted += $res.Deleted; $out.Bytes += $res.Bytes
        $problems = $res.DlFailed + $res.VerifyFailed + $res.DeleteFailed
        if ($res.Cancelled) { $out.Cancelled = $true }
        if ($problems -eq 0 -and -not $res.Cancelled) { $out.LibOk++; $out.Rows.Add(@{ Label = $lib.Title; Status = ("OK - {0} item(s) archived and recycled ({1})" -f $res.Deleted, (Format-FileSize $res.Bytes)); Level = 'SUCCESS' }) }
        else {
            $out.LibProblems++
            $st = if ($res.Cancelled) { 'CANCELLED part-way' } else { "{0} download / {1} verify / {2} recycle problem(s)" -f $res.DlFailed, $res.VerifyFailed, $res.DeleteFailed }
            $out.Rows.Add(@{ Label = $lib.Title; Status = $st; Level = $(if ($res.Cancelled) { 'WARN' } else { 'ERROR' }) })
        }
        if ($script:JobCancelled) { $out.Cancelled = $true }
    }
    return $out
}

# Option 20.
function Invoke-LeaverArchiveSingle {
    param([string]$SiteUrl, [string]$LocalDest, [string]$BufferDest)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Start-JobUi
    $res = $null
    try {
        try {
            Connect-CloudSite -SiteUrl $SiteUrl -Force
            $null = Invoke-WithRetry -OperationName 'Open OneDrive' -Quiet -ScriptBlock { Get-PnPList -ErrorAction Stop | Select-Object -First 1 }
            Write-Ok 'Connected to the OneDrive.'
        } catch {
            Write-Bad "Could not open this OneDrive: $($_.Exception.Message)"
            Write-Hint 'Possible causes: no access for your account, the OneDrive was deleted, or the leaver has no licence (use option 23).'
            return
        }
        $res = Perform-LeaverOneDriveArchive -LeaverSiteUrl $SiteUrl -LocalDest $LocalDest -BufferDest $BufferDest
    } finally { $sw.Stop(); Stop-JobUi }
    $summary = @(
        @{ Label = 'OneDrive'; Value = ($SiteUrl -replace '^https://', '') }
        @{ Label = 'Libraries OK'; Value = ("{0} of {1}" -f $res.LibOk, $res.Libraries); Level = $(if ($res.LibProblems -or $res.Error) { 'WARN' } else { 'SUCCESS' }) }
        @{ Label = 'Items archived + recycled'; Value = ("{0}  |  {1}" -f $res.Deleted, (Format-FileSize $res.Bytes)) }
        @{ Label = 'Time taken'; Value = (Format-Elapsed $sw.Elapsed) }
    )
    if ($res.Error) { $summary += @{ Label = 'Error'; Value = $res.Error; Level = 'ERROR' } }
    Show-OutcomeTable -Title 'LEAVER ONEDRIVE ARCHIVE - RESULTS' -Rows $res.Rows.ToArray() -Summary $summary
    $st = New-JobStats; $st.Success = $res.Deleted; $st.SuccessBytes = $res.Bytes; $st.Failed = $res.LibProblems + [int][bool]$res.Error; $st.Cancelled = $res.Cancelled
    Register-JobResult -Operation 'Leaver archive' -Source $SiteUrl -Destination $LocalDest -Stats $st -Elapsed $sw.Elapsed
    if ($res.Deleted -gt 0) { Write-RecycleBinLinks -Links @(@{ Label = $SiteUrl; Url = (Get-LeaverRecycleBinUrl $SiteUrl) }) }
}

# Option 21: several leavers, unattended.
function Perform-MultiSourceLeaverArchive {
    param([object[]]$Leavers, [string]$LocalDest, [string]$BufferDest)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[hashtable]
    $links = New-Object System.Collections.Generic.List[hashtable]
    $ok = 0; $bad = 0; [long]$bytes = 0
    Start-JobUi
    try {
        for ($i = 0; $i -lt $Leavers.Count; $i++) {
            $lv = $Leavers[$i]
            if ((Check-Interrupt) -eq 'CANCEL' -or $script:JobCancelled) { for ($j = $i; $j -lt $Leavers.Count; $j++) { $rows.Add(@{ Label = $Leavers[$j].Label; Status = 'NOT STARTED (cancelled)'; Level = 'WARN' }) }; break }
            Update-Progress -Id 1 -Activity 'Multi-leaver archive' -Status ("Leaver {0} of {1}: {2}" -f ($i + 1), $Leavers.Count, $lv.Label) -Percent ($i * 100.0 / $Leavers.Count) -Detail 'P = pause, C = cancel'
            Write-Host ""
            Write-Rule -Title ("LEAVER {0} OF {1}: {2}" -f ($i + 1), $Leavers.Count, $lv.Label) -Color DarkYellow -Double
            $dest = Join-LocalPath (Resolve-LocalPath $LocalDest) $lv.Label
            Write-Hint "OneDrive: $($lv.SiteUrl)"
            Write-Hint "Saving  : $dest"
            try {
                Connect-CloudSite -SiteUrl $lv.SiteUrl -Force
                $null = Invoke-WithRetry -OperationName 'Open OneDrive' -Quiet -ScriptBlock { Get-PnPList -ErrorAction Stop | Select-Object -First 1 }
            } catch { $bad++; $rows.Add(@{ Label = $lv.Label; Status = "CONNECTION FAILED - $($_.Exception.Message)"; Level = 'ERROR' }); continue }
            $r = Perform-LeaverOneDriveArchive -LeaverSiteUrl $lv.SiteUrl -LocalDest $dest -BufferDest $BufferDest
            $bytes += $r.Bytes
            if ($r.Deleted -gt 0) { $links.Add(@{ Label = $lv.Label; Url = (Get-LeaverRecycleBinUrl $lv.SiteUrl) }) }
            if ($r.Error) { $bad++; $rows.Add(@{ Label = $lv.Label; Status = $r.Error; Level = 'ERROR' }) }
            elseif ($r.Cancelled) { $bad++; $rows.Add(@{ Label = $lv.Label; Status = 'CANCELLED part-way'; Level = 'WARN' }) }
            elseif ($r.LibProblems -gt 0) { $bad++; $rows.Add(@{ Label = $lv.Label; Status = ("{0} of {1} libraries had problems - see above" -f $r.LibProblems, $r.Libraries); Level = 'ERROR' }) }
            elseif ($r.Libraries -eq 0) { $ok++; $rows.Add(@{ Label = $lv.Label; Status = 'Nothing to archive (no libraries / already empty)'; Level = 'SUCCESS' }) }
            else { $ok++; $rows.Add(@{ Label = $lv.Label; Status = ("ARCHIVED - {0} item(s), {1}" -f $r.Deleted, (Format-FileSize $r.Bytes)); Level = 'SUCCESS' }) }
        }
    } finally { $sw.Stop(); Stop-JobUi }
    Show-OutcomeTable -Title 'MULTI-LEAVER ARCHIVE - RESULTS' -Rows $rows.ToArray() -Summary @(
        @{ Label = 'Leavers fully archived'; Value = ("{0} of {1}" -f $ok, $Leavers.Count); Level = $(if ($ok -eq $Leavers.Count) { 'SUCCESS' } else { 'WARN' }) }
        @{ Label = 'Data archived'; Value = (Format-FileSize $bytes) }
        @{ Label = 'Time taken'; Value = (Format-Elapsed $sw.Elapsed) }
    )
    $st = New-JobStats; $st.Success = $ok; $st.Failed = $bad; $st.SuccessBytes = $bytes; $st.Cancelled = $script:JobCancelled
    Register-JobResult -Operation 'Multi-leaver archive' -Source "$($Leavers.Count) leavers" -Destination $LocalDest -Stats $st -Elapsed $sw.Elapsed
    if ($links.Count -gt 0) { Write-RecycleBinLinks -Links $links.ToArray() }
}

# ======================================================================================
# SECTION 15 - MICROSOFT GRAPH (option 23) - pure REST, works on PowerShell 5.1
# ======================================================================================

function Save-MgTokenCache {
    param([string]$RefreshToken, [string]$AccountName)
    try {
        $enc = ConvertFrom-SecureString (ConvertTo-SecureString $RefreshToken -AsPlainText -Force)
        $json = ConvertTo-Json -InputObject ([pscustomobject]@{ Token = $enc; Account = $AccountName; At = (Get-Date -Format 'o') })
        [System.IO.File]::WriteAllText($script:MgTokenCacheFile, $json)
    } catch { Write-Log "Graph sign-in could not be saved: $($_.Exception.Message)" 'WARN' -FileOnly }
}

function Load-MgTokenCache {
    try {
        if (-not (Test-Path -LiteralPath $script:MgTokenCacheFile)) { return $null }
        $raw = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($script:MgTokenCacheFile))
        $dec = (New-Object System.Net.NetworkCredential('', (ConvertTo-SecureString $raw.Token))).Password
        return @{ RefreshToken = $dec; AccountName = $raw.Account }
    } catch { Write-Log "Saved Graph sign-in could not be read: $($_.Exception.Message)" 'WARN' -FileOnly; return $null }
}

function Get-JwtAccountName {
    param([string]$Token)
    try {
        $p = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
        $p += '=' * ((4 - ($p.Length % 4)) % 4)
        $c = ConvertFrom-Json -InputObject ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)))
        foreach ($n in 'upn', 'preferred_username', 'unique_name', 'email') { if ($c.$n) { return [string]$c.$n } }
    } catch { }
    return 'unknown account'
}

function Get-RestErrorInfo {
    param($ErrorRecord)
    $info = @{ Code = ''; Description = $ErrorRecord.Exception.Message; Status = 0 }
    try { if ($ErrorRecord.Exception.Response) { $info.Status = [int]$ErrorRecord.Exception.Response.StatusCode } } catch { }
    try {
        $body = $ErrorRecord.ErrorDetails.Message
        if ($body) {
            $j = ConvertFrom-Json -InputObject $body
            if ($j.error -is [string]) { $info.Code = $j.error; if ($j.error_description) { $info.Description = $j.error_description } }
            elseif ($j.error) { $info.Code = [string]$j.error.code; if ($j.error.message) { $info.Description = [string]$j.error.message } }
        }
    } catch { }
    return $info
}

function Update-MgGraphToken {
    if (-not $script:MgRefreshToken) { return $false }
    try {
        $body = "client_id=$($script:MgClientId)&grant_type=refresh_token&refresh_token=$([uri]::EscapeDataString($script:MgRefreshToken))&scope=$([uri]::EscapeDataString($script:MgScopes))"
        $resp = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/token' -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $script:MgAccessToken = $resp.access_token
        if ($resp.refresh_token) { $script:MgRefreshToken = $resp.refresh_token }
        $script:MgTokenExpiry = (Get-Date).AddSeconds([int]$resp.expires_in)
        $script:MgGraphConnected = $true
        $script:MgAccount = Get-JwtAccountName $resp.access_token
        return $true
    } catch {
        Write-Log "Graph token refresh failed: $((Get-RestErrorInfo $_).Description)" 'WARN' -FileOnly
        $script:MgGraphConnected = $false
        return $false
    }
}

# Authenticated Graph call with token refresh and throttling (429/503) handling.
function Invoke-MgRestMethod {
    param([string]$Method, [string]$Endpoint, $Body = $null)
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        if ($script:MgTokenExpiry -and (Get-Date) -gt $script:MgTokenExpiry.AddMinutes(-5)) {
            if (-not (Update-MgGraphToken)) { throw 'The Microsoft Graph sign-in has expired. Run option 23 again to sign in.' }
        }
        $p = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0/$Endpoint"; Headers = @{ Authorization = "Bearer $($script:MgAccessToken)" }; ErrorAction = 'Stop' }
        if ($null -ne $Body) { $p.Body = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Body -Depth 5 -Compress)); $p.ContentType = 'application/json; charset=utf-8' }
        try { return (Invoke-RestMethod @p) }
        catch {
            $e = Get-RestErrorInfo $_
            if (($e.Status -eq 429 -or $e.Status -eq 503 -or $e.Status -eq 504) -and $attempt -lt 4) {
                $wait = 5 * $attempt
                try { $ra = $_.Exception.Response.Headers['Retry-After']; if ($ra) { $wait = [Math]::Min(60, [int]$ra) } } catch { }
                Start-InterruptibleSleep -Seconds $wait -Label 'Microsoft Graph is busy - waiting'
                continue
            }
            if ($e.Status -eq 401 -and $attempt -lt 4 -and (Update-MgGraphToken)) { continue }
            throw $e.Description
        }
    }
}

# Signs in to Microsoft Graph: saved sign-in first, then the device-code flow.
function Ensure-MgGraphForLicensing {
    if ($script:MgGraphConnected -and $script:MgTokenExpiry -and (Get-Date) -lt $script:MgTokenExpiry.AddMinutes(-5)) { return $true }
    if ($script:MgGraphConnected -and (Update-MgGraphToken)) { return $true }
    Write-Rule -Title 'MICROSOFT GRAPH SIGN-IN (licence group changes)' -Color Cyan
    $cached = Load-MgTokenCache
    if ($cached) {
        $script:MgRefreshToken = $cached.RefreshToken
        if (Update-MgGraphToken) {
            if (Read-YesNo -Question "Continue as $($script:MgAccount) (saved sign-in)?" -Default $true) {
                Save-MgTokenCache -RefreshToken $script:MgRefreshToken -AccountName $script:MgAccount
                return $true
            }
            $script:MgGraphConnected = $false; $script:MgRefreshToken = ''; $script:MgAccessToken = ''
            try { Remove-Item -LiteralPath $script:MgTokenCacheFile -Force -ErrorAction Stop } catch { }
            Write-Hint 'OK - sign in with the account you want to use.'
        } else { Write-Hint 'The saved sign-in has expired - please sign in again.' }
    }
    try {
        $dc = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode' -Body "client_id=$($script:MgClientId)&scope=$([uri]::EscapeDataString($script:MgScopes))" -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    } catch {
        Write-Bad "Could not start the sign-in: $((Get-RestErrorInfo $_).Description)"
        return $false
    }
    Write-Panel -Title 'SIGN IN' -Color Cyan -Lines @(
        @{ Label = 'Open';       Value = [string]$dc.verification_uri; Level = 'SUCCESS' }
        @{ Label = 'Enter code'; Value = [string]$dc.user_code; Level = 'SUCCESS' }
        'Sign in with an account that can manage the licence group. This window continues by itself.'
    )
    if (Set-ClipboardText ([string]$dc.user_code)) { Write-Hint 'The code is already on your clipboard.' }
    $interval = [Math]::Max(5, [int]$dc.interval)
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $body = "client_id=$($script:MgClientId)&grant_type=urn:ietf:params:oauth:grant-type:device_code&device_code=$([uri]::EscapeDataString($dc.device_code))"
    $token = $null
    while ((Get-Date) -lt $deadline) {
        Start-InterruptibleSleep -Seconds $interval -Label 'Waiting for you to sign in'
        if ($script:JobCancelled) { $script:JobCancelled = $false; Write-Caution 'Sign-in cancelled.'; return $false }
        try {
            $token = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/token' -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            break
        } catch {
            $e = Get-RestErrorInfo $_
            switch ($e.Code) {
                'authorization_pending' { continue }
                'slow_down' { $interval += 5; continue }
                'expired_token' { Write-Bad 'The sign-in code expired. Run option 23 again.'; return $false }
                { $_ -in @('authorization_declined', 'access_denied') } { Write-Bad 'The sign-in was declined.'; return $false }
                '' { Write-Log "Device-code poll network error: $($e.Description)" 'WARN' -FileOnly; continue }
                default { Write-Bad "Sign-in failed: $($e.Description)"; return $false }
            }
        }
    }
    if (-not $token -or -not $token.access_token) { Write-Bad 'Sign-in timed out. Run option 23 again.'; return $false }
    $script:MgAccessToken = $token.access_token
    $script:MgRefreshToken = $token.refresh_token
    $script:MgTokenExpiry = (Get-Date).AddSeconds([int]$token.expires_in)
    $script:MgGraphConnected = $true
    $script:MgAccount = Get-JwtAccountName $token.access_token
    Write-Ok "Signed in as $($script:MgAccount). The sign-in is saved (encrypted) for next time."
    Write-Log "Microsoft Graph: signed in as $($script:MgAccount)." 'INFO' -FileOnly
    if ($script:MgRefreshToken) { Save-MgTokenCache -RefreshToken $script:MgRefreshToken -AccountName $script:MgAccount }
    return $true
}

# Finds a user by UPN, falling back to the mail address.
function Resolve-MgUser {
    param([string]$Email)
    try { return (Invoke-MgRestMethod -Method GET -Endpoint "users/$([uri]::EscapeDataString($Email))?`$select=id,displayName,userPrincipalName,mail") }
    catch {
        $q = $Email -replace "'", "''"
        $f = [uri]::EscapeDataString("userPrincipalName eq '$q' or mail eq '$q'")
        $r = Invoke-MgRestMethod -Method GET -Endpoint "users?`$filter=$f&`$select=id,displayName,userPrincipalName,mail"
        $u = @($r.value)
        if ($u.Count -ge 1) { return $u[0] }
        throw "No user with the address '$Email' was found."
    }
}

# $true / $false = member or not, $null = could not be determined.
function Test-MgGroupMember {
    param([string]$GroupId, [string]$UserId)
    try {
        $r = Invoke-MgRestMethod -Method POST -Endpoint "users/$UserId/checkMemberGroups" -Body @{ groupIds = @($GroupId) }
        return (@($r.value) -contains $GroupId)
    } catch { return $null }
}

function Remove-MgGroupMember {
    param([string]$GroupId, [string]$UserId)
    $null = Invoke-MgRestMethod -Method DELETE -Endpoint "groups/$GroupId/members/$UserId/`$ref"
}

# Removes a user from the licence group, retrying for a few minutes and confirming the
# result - it must not give up because of a brief network or throttling hiccup, and it
# runs even after the job was cancelled. Returns '' on success or the last error text.
function Remove-MgGroupMemberSafe {
    param([string]$GroupId, [string]$UserId)
    $waits = @(0, 5, 10, 20, 30, 60, 90)
    $last = ''
    $wasCancelled = $script:JobCancelled
    $script:JobCancelled = $false
    try {
        foreach ($w in $waits) {
            if ($w -gt 0) { Write-Hint "Retrying the licence removal in ${w}s..."; Start-Sleep -Seconds $w }
            try { Remove-MgGroupMember -GroupId $GroupId -UserId $UserId }
            catch { $last = "$($_.Exception.Message)" }
            $member = Test-MgGroupMember -GroupId $GroupId -UserId $UserId
            if ($member -eq $false) { return '' }
            if ($null -eq $member -and $last -match '(?i)does not exist|not found|ResourceNotFound') { return '' }
            if (-not $last) { $last = 'The user still shows as a member of the group.' }
        }
        return $last
    } finally { if ($wasCancelled) { $script:JobCancelled = $true } }
}

# Returns a licence-assignment error message for this group, or ''.
function Get-MgLicenseError {
    param([string]$UserId, [string]$GroupId)
    try {
        $u = Invoke-MgRestMethod -Method GET -Endpoint "users/$UserId`?`$select=licenseAssignmentStates"
        foreach ($s in @($u.licenseAssignmentStates)) {
            if ($s.assignedByGroup -eq $GroupId -and $s.state -eq 'Error') { return "Licence assignment error: $($s.error)" }
        }
    } catch { }
    return ''
}

# Option 23 engine. One leaver at a time: add to licence group -> wait for OneDrive ->
# archive -> remove from group (also when something fails or you cancel).
function Perform-LicensedLeaverArchive {
    param([object[]]$Leavers, [string]$LocalDest, [string]$BufferDest, [string]$GroupId, [string]$GroupName)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = New-Object System.Collections.Generic.List[hashtable]
    $links = New-Object System.Collections.Generic.List[hashtable]
    $ok = 0; $bad = 0; [long]$bytes = 0
    Write-Log "LICENSED LEAVER ARCHIVE: $($Leavers.Count) leaver(s), group '$GroupName'" 'INFO'
    Start-JobUi
    try {
        for ($i = 0; $i -lt $Leavers.Count; $i++) {
            $lv = $Leavers[$i]
            if ((Check-Interrupt) -eq 'CANCEL' -or $script:JobCancelled) { for ($j = $i; $j -lt $Leavers.Count; $j++) { $rows.Add(@{ Label = $Leavers[$j].FolderName; Status = 'NOT STARTED (cancelled)'; Level = 'WARN' }) }; break }
            $dest = Join-LocalPath (Resolve-LocalPath $LocalDest) $lv.FolderName
            Write-Host ""
            Write-Rule -Title ("LEAVER {0} OF {1}: {2}" -f ($i + 1), $Leavers.Count, $lv.FolderName) -Color Cyan -Double
            Write-Hint "Email   : $($lv.Email)"
            Write-Hint "OneDrive: $($lv.SiteUrl)"
            Write-Hint "Saving  : $dest"
            Update-Progress -Id 1 -Activity 'Unlicensed leaver archive' -Status ("Leaver {0} of {1}: {2}" -f ($i + 1), $Leavers.Count, $lv.FolderName) -Percent ($i * 100.0 / $Leavers.Count)

            Write-Step '1/5' 'Finding the user...'
            try { $user = Resolve-MgUser -Email $lv.Email; Write-Ok "Found $($user.displayName)" }
            catch { $bad++; $rows.Add(@{ Label = $lv.FolderName; Status = "SKIPPED - $($_.Exception.Message)"; Level = 'ERROR' }); continue }

            Write-Step '2/5' "Adding to licence group '$GroupName'..."
            $addedByUs = $false
            $before = Test-MgGroupMember -GroupId $GroupId -UserId $user.id
            if ($before -eq $true) {
                Write-Hint 'Already in the group before this run - it will be left in the group afterwards.'
            } else {
                $addErr = ''
                try { $null = Invoke-MgRestMethod -Method POST -Endpoint "groups/$GroupId/members/`$ref" -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" } }
                catch { $addErr = "$($_.Exception.Message)" }
                # Decide by checking membership, so a retried or timed-out request can never
                # leave someone in the group without this run knowing to remove them again.
                $after = Test-MgGroupMember -GroupId $GroupId -UserId $user.id
                if ($after -eq $true -or (-not $addErr -and $null -eq $after)) {
                    $addedByUs = $true
                    Write-Ok 'Added - the licence is being assigned.'
                    Write-Log "Added $($user.displayName) to licence group $GroupName." 'INFO' -FileOnly
                } else {
                    $bad++
                    $rows.Add(@{ Label = $lv.FolderName; Status = "FAILED - could not add to the licence group: $addErr"; Level = 'ERROR' })
                    continue
                }
            }

            $status = ''; $level = 'ERROR'
            try {
                Write-Step '3/5' 'Waiting for the OneDrive to open (checked every 30 s, up to 30 min)...'
                $poll = [System.Diagnostics.Stopwatch]::StartNew()
                $open = $false; $first = $true; $licErr = ''; $lastPollErr = ''
                while ($poll.Elapsed.TotalMinutes -lt 30 -and -not $script:JobCancelled) {
                    if (-not $first) {
                        $licErr = Get-MgLicenseError -UserId $user.id -GroupId $GroupId
                        if ($licErr) { break }
                        $why = if ($lastPollErr) { ' - ' + (Get-FitText -Text $lastPollErr -Width 40) } else { '' }
                        Start-InterruptibleSleep -Seconds 30 -Label ("Not open yet ({0} waited){1}" -f (Format-Elapsed $poll.Elapsed), $why)
                        if ($script:JobCancelled) { break }
                    }
                    $first = $false
                    try {
                        Connect-CloudSite -SiteUrl $lv.SiteUrl -Force
                        $null = Get-PnPList -ErrorAction Stop | Select-Object -First 1
                        $open = $true; break
                    } catch {
                        $lastPollErr = "$($_.Exception.Message)"
                        Write-Log "OneDrive not open yet ($([int]$poll.Elapsed.TotalSeconds)s): $lastPollErr" 'INFO' -FileOnly
                    }
                }
                if ($script:JobCancelled) { $status = 'CANCELLED while waiting for the licence'; $level = 'WARN' }
                elseif ($licErr) { $status = "FAILED - $licErr"; Write-Bad $licErr }
                elseif (-not $open) { $status = 'FAILED - the OneDrive did not open within 30 minutes' + $(if ($lastPollErr) { " (last error: $lastPollErr)" } else { '' }); Write-Bad $status }
                else {
                    Write-Ok ("OneDrive is open (after {0})." -f (Format-Elapsed $poll.Elapsed))
                    Write-Step '4/5' 'Archiving...'
                    $r = Perform-LeaverOneDriveArchive -LeaverSiteUrl $lv.SiteUrl -LocalDest $dest -BufferDest $BufferDest
                    $bytes += $r.Bytes
                    if ($r.Deleted -gt 0) { $links.Add(@{ Label = $lv.FolderName; Url = (Get-LeaverRecycleBinUrl $lv.SiteUrl) }) }
                    if ($r.Error) { $status = $r.Error }
                    elseif ($r.Cancelled) { $status = 'CANCELLED part-way'; $level = 'WARN' }
                    elseif ($r.LibProblems -gt 0) { $status = ("{0} of {1} libraries had problems - see above" -f $r.LibProblems, $r.Libraries) }
                    else { $status = ("ARCHIVED - {0} item(s), {1}" -f $r.Deleted, (Format-FileSize $r.Bytes)); $level = 'SUCCESS' }
                }
            } finally {
                Write-Step '5/5' 'Freeing the licence...'
                if ($addedByUs) {
                    $remErr = Remove-MgGroupMemberSafe -GroupId $GroupId -UserId $user.id
                    if (-not $remErr) {
                        Write-Ok 'Removed from the licence group - the licence is free for the next leaver.'
                        Write-Log "Removed $($user.displayName) from licence group $GroupName." 'INFO' -FileOnly
                    } else {
                        Write-Bad "Could not remove $($user.displayName) from the licence group: $remErr"
                        Write-Log "WARNING: $($user.displayName) is still in licence group $GroupName : $remErr" 'WARN'
                        $status += ' (STILL IN LICENCE GROUP - remove manually)'; $level = 'ERROR'
                    }
                } else { Write-Hint 'Not removed: the user was already in the group before this run.' }
            }
            if ($level -eq 'SUCCESS') { $ok++ } else { $bad++ }
            $rows.Add(@{ Label = $lv.FolderName; Status = $status; Level = $level })
        }
    } finally { $sw.Stop(); Stop-JobUi }
    Show-OutcomeTable -Title 'UNLICENSED LEAVER ARCHIVE - RESULTS' -Rows $rows.ToArray() -Summary @(
        @{ Label = 'Leavers archived'; Value = ("{0} of {1}" -f $ok, $Leavers.Count); Level = $(if ($ok -eq $Leavers.Count) { 'SUCCESS' } else { 'WARN' }) }
        @{ Label = 'Licence group'; Value = $GroupName }
        @{ Label = 'Data archived'; Value = (Format-FileSize $bytes) }
        @{ Label = 'Time taken'; Value = (Format-Elapsed $sw.Elapsed) }
    )
    $st = New-JobStats; $st.Success = $ok; $st.Failed = $bad; $st.SuccessBytes = $bytes; $st.Cancelled = $script:JobCancelled
    Register-JobResult -Operation 'Unlicensed leaver archive' -Source "$($Leavers.Count) leavers" -Destination $LocalDest -Stats $st -Elapsed $sw.Elapsed
    if ($links.Count -gt 0) { Write-RecycleBinLinks -Links $links.ToArray() }
}

# ======================================================================================
# SECTION 16 - OPTION SCREENS (one function per menu option)
# ======================================================================================

$script:GuideTipShown = $false

function Show-OpIntro {
    param([string]$Key, [string]$Title, [string]$Color = 'Cyan', [string[]]$What = @(), [string[]]$Notes = @(), [string]$Danger = '')
    Write-Banner -Title ("[{0}]  {1}" -f $Key, $Title) -Color $Color
    if ($script:Settings.ShowGuides) {
        $lines = New-Object System.Collections.Generic.List[object]
        foreach ($w in $What) { $lines.Add(@{ Text = $w; Color = 'White' }) }
        if ($Notes.Count -gt 0) { $lines.Add('-'); foreach ($n in $Notes) { $lines.Add(@{ Text = $n; Color = 'Gray' }) } }
        if ($Danger) { $lines.Add('-'); $lines.Add(@{ Text = $Danger; Color = 'Yellow' }) }
        if ($lines.Count -gt 0) { Write-Panel -Lines $lines.ToArray() -Color $Color }
        if (-not $script:GuideTipShown) { Write-Hint 'Know all this already? Turn these explanations off in [S] Settings.'; $script:GuideTipShown = $true }
    } elseif ($Danger) { Write-Caution $Danger }
}

# Shows what is about to happen and asks to start. -Danger requires typing YES.
function Confirm-Start {
    param([string]$Title, [object[]]$Rows, [string]$Color = 'Cyan', [switch]$Danger, [string]$DangerText = '')
    Write-Host ""
    Write-Panel -Title $Title -Lines $Rows -Color $Color
    if ($Danger) { return (Read-Confirmation -Message $(if ($DangerText) { $DangerText } else { 'Start now?' }) -Color Yellow) }
    return (Read-YesNo -Question 'Start now?' -Default $true)
}

function Read-RunAgain {
    param([string]$What)
    Write-Host ""
    $a = (Read-Line -Prompt ("  ENTER = {0}   |   M = main menu > " -f $What) -LogLabel 'Run again').Trim()
    return -not ($a -match '^(?i)(m|q|menu)$')
}

function Invoke-OpCloudCopyMove {
    param([switch]$Move)
    $verb = if ($Move) { 'move' } else { 'copy' }
    Show-OpIntro -Key $(if ($Move) { '2' } else { '1' }) -Title "CLOUD $($verb.ToUpper()) - server-side, inside SharePoint / OneDrive" -Color Yellow -What @(
        "Microsoft's servers $verb the folder for you. Nothing is downloaded to this PC, so even huge folders use no local disk space."
        "The source folder is placed INSIDE the destination folder (e.g. 'Reports' copied into 'Archive' becomes 'Archive/Reports')."
    ) -Notes @(
        'Both addresses must be in the same Microsoft 365 tenant (different sites are fine).'
        'Between tenants use [13] Download then [15] Upload.'
        'How to get an address: open the folder in your browser and copy the address bar.'
    )
    while ($true) {
        $src = Read-CloudFolder -Prompt "SharePoint / OneDrive folder to $verb (paste its address):" -LogLabel 'Source URL'
        if (-not $src) { return }
        $dst = Read-CloudFolder -Prompt "Destination folder it should go INTO:" -LogLabel 'Destination URL'
        if (-not $dst) { return }
        if (Confirm-Start -Title "READY TO $($verb.ToUpper())" -Color Yellow -Rows @(
                @{ Label = 'From'; Value = $src.FolderPath }, @{ Label = 'Into'; Value = $dst.FolderPath }
                @{ Label = 'Result'; Value = (Join-CloudPath $dst.FolderPath (($src.FolderPath -split '/')[-1])) })) {
            Perform-CloudTransfer -SrcData $src -DstData $dst -IsMove:$Move
        }
        if (-not (Read-RunAgain "another cloud $verb")) { return }
    }
}

function Invoke-OpCloudBatch {
    Show-OpIntro -Key '3' -Title 'CLOUD BATCH - pick items to copy, move, delete or verify' -Color Yellow -What @(
        'Shows the contents of a SharePoint / OneDrive folder as a numbered list. Pick items (3 | 2-5 | 1,3,7 | A) and choose what to do with them.'
    ) -Notes @('The destination is only needed for copy, move and verify - you will be asked when it is needed.', 'Deleted items go to the SharePoint Recycle Bin.')
    while ($true) {
        $src = Read-CloudFolder -Prompt 'SharePoint / OneDrive folder to work in (paste its address):' -LogLabel 'Source URL'
        if (-not $src) { return }
        $dst = Read-CloudFolder -Prompt 'Destination folder (for copy / move / verify):' -LogLabel 'Destination URL' -AllowEmpty -EmptyHint 'decide later'
        if ($null -eq $dst) { return }
        Perform-CloudBatch -SrcData $src -DstData $dst
        if (-not (Read-RunAgain 'another cloud batch')) { return }
    }
}

function Invoke-OpCloudWipe {
    Show-OpIntro -Key '4' -Title 'CLOUD WIPE - send a folder to the SharePoint Recycle Bin' -Color Yellow -What @(
        'Deletes a SharePoint / OneDrive folder (or file) and everything in it. Items go to the Recycle Bin and can be restored for about 93 days.'
    ) -Danger 'You will see what is about to be deleted and must type YES to confirm.'
    while ($true) {
        $t = Read-CloudFolder -Prompt 'Address of the folder or file to delete:' -LogLabel 'Wipe URL'
        if (-not $t) { return }
        Perform-CloudWipe -TargetData $t
        if (-not (Read-RunAgain 'delete another')) { return }
    }
}

function Invoke-OpCloudVerify {
    Show-OpIntro -Key '5' -Title 'CLOUD VERIFY - prove two cloud folders match' -Color Yellow -What @(
        'Compares every file name and exact size in two SharePoint / OneDrive folders without downloading anything.'
    ) -Notes @('MISSING = in the source but not the destination.  DIFFERENT = same name, different size.  EXTRA = only in the destination.')
    while ($true) {
        $src = Read-CloudFolder -Prompt 'Original (source) folder address:' -LogLabel 'Source URL'
        if (-not $src) { return }
        $dst = Read-CloudFolder -Prompt 'Copy (destination) folder address to check:' -LogLabel 'Destination URL'
        if (-not $dst) { return }
        Perform-CloudVerify -SrcData $src -DstData $dst
        if (-not (Read-RunAgain 'verify another pair')) { return }
    }
}

function Invoke-OpCloudSearch {
    Show-OpIntro -Key '6' -Title 'CLOUD SEARCH - find files and folders in SharePoint / OneDrive' -Color Yellow -What @(
        'Searches every sub-folder for names that match. Type a word (invoice), an extension (*.pdf) or a pattern (*2024*report*).'
    ) -Notes @('Results can be opened in the browser, have their link copied, or be exported to CSV.')
    while ($true) {
        $src = Read-CloudFolder -Prompt 'Folder to search in (paste its address):' -LogLabel 'Cloud search URL'
        if (-not $src) { return }
        try { Connect-CloudSite -SiteUrl $src.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; continue }
        Invoke-CloudSearchLoop -CloudData $src
        if (-not (Read-RunAgain 'search another folder')) { return }
    }
}

function Invoke-OpLocalCopyMove {
    param([switch]$Move)
    $verb = if ($Move) { 'move' } else { 'copy' }
    Show-OpIntro -Key $(if ($Move) { '8' } else { '7' }) -Title "LOCAL $($verb.ToUpper()) - between folders on this PC or the network" -Color Green -What @(
        "Smart $verb of a folder (or a single file): files that are already identical at the destination are skipped, so an interrupted job simply resumes."
        'Works with local drives, USB drives, mapped drives and network paths (\\server\share\folder).'
    ) -Notes @(
        'Moving between two SharePoint tenants? Sync both libraries with OneDrive and move between the two synced folders.'
        'Free space is checked before the job starts. Dates are preserved.'
    ) -Danger $(if ($Move) { 'A move removes each source file after it has been copied, and removes the emptied source folders.' } else { '' })
    while ($true) {
        $src = Read-LocalFolder -Prompt "Folder (or file) to $verb FROM:" -LogLabel 'Source path' -MustExist -AllowFile
        if (-not $src) { return }
        $dst = Read-LocalFolder -Prompt "Folder to $verb INTO (it is created if missing):" -LogLabel 'Destination path' -CreateIfMissing
        if (-not $dst) { return }
        $ov = Test-PathOverlap -Source $src -Destination $dst
        if ($ov -and -not [System.IO.File]::Exists((Get-LongPath $src))) { Write-Bad "$ov Please choose different folders."; continue }
        if (Confirm-Start -Title "READY TO $($verb.ToUpper())" -Color Green -Rows @(@{ Label = 'From'; Value = (Format-PathWithStorageTag $src) }, @{ Label = 'Into'; Value = (Format-PathWithStorageTag $dst) })) {
            $null = Perform-LocalTransfer -Src $src -Dst $dst -IsMove:$Move
        }
        if (-not (Read-RunAgain "another local $verb")) { return }
    }
}

function Invoke-OpLocalBatch {
    Show-OpIntro -Key '9' -Title 'LOCAL BATCH - pick items to copy, move, delete or verify' -Color Green -What @(
        'Shows the top level of a folder as a numbered list with sizes. Pick items (3 | 2-5 | 1,3,7 | A) and choose copy, move, delete or verify.'
    ) -Notes @('Delete here is permanent (no Recycle Bin) and always asks you to type YES.')
    while ($true) {
        $src = Read-LocalFolder -Prompt 'Folder whose contents you want to work with:' -LogLabel 'Source path' -MustExist
        if (-not $src) { return }
        $dst = Read-LocalFolder -Prompt 'Destination folder (for copy / move / verify):' -LogLabel 'Destination path' -AllowEmpty -EmptyHint 'decide later' -CreateIfMissing
        if ($null -eq $dst) { return }
        Perform-LocalBatch -Src $src -Dst $dst
        if (-not (Read-RunAgain 'another local batch')) { return }
    }
}

function Invoke-OpLocalWipe {
    Show-OpIntro -Key '10' -Title 'LOCAL WIPE - permanently delete a folder' -Color Green -What @(
        'Deletes a folder and everything in it from this PC or a network share.'
        'Recommended: copy [7], verify [11], and only then wipe the source.'
    ) -Danger 'This bypasses the Windows Recycle Bin - it cannot be undone. Drive roots, share roots and system / profile folders are refused.'
    while ($true) {
        $t = Read-LocalFolder -Prompt 'Folder to delete permanently:' -LogLabel 'Wipe path' -MustExist
        if (-not $t) { return }
        Perform-LocalWipe -TargetFolder $t
        if (-not (Read-RunAgain 'wipe another folder')) { return }
    }
}

function Invoke-OpLocalVerify {
    Show-OpIntro -Key '11' -Title 'LOCAL VERIFY - prove two folders match' -Color Green -What @(
        'Compares every file in two folders: present in both, same size and same date. Optionally also compares SHA-256 checksums (slower, proves the content).'
    ) -Notes @('MISSING = in the source but not the destination.  DIFFERENT = name matches but size, date or content differs.  EXTRA = only in the destination.')
    while ($true) {
        $src = Read-LocalFolder -Prompt 'Original (source) folder:' -LogLabel 'Source path' -MustExist
        if (-not $src) { return }
        $dst = Read-LocalFolder -Prompt 'Copy (destination) folder to check:' -LogLabel 'Destination path' -MustExist
        if (-not $dst) { return }
        $ov = Test-PathOverlap -Source $src -Destination $dst
        if ($ov) { Write-Bad "$ov Comparing a folder with itself proves nothing - choose the copy to check."; continue }
        $hash = Read-YesNo -Question 'Also compare checksums (slower, reads every byte)?' -Default $false
        $null = Perform-LocalVerify -Src $src -Dst $dst -Hash:$hash
        if (-not (Read-RunAgain 'verify another pair')) { return }
    }
}

function Invoke-OpLocalSearch {
    Show-OpIntro -Key '12' -Title 'LOCAL SEARCH - find files and folders on this PC or the network' -Color Green -What @(
        'Searches every sub-folder for names that match. Type a word (invoice), an extension (*.pdf) or a pattern (*2024*report*).'
    ) -Notes @('Folders that cannot be read are skipped and counted. Results can be shown in Explorer, copied or exported to CSV.')
    while ($true) {
        $p = Read-LocalFolder -Prompt 'Folder to search in:' -LogLabel 'Search path' -MustExist
        if (-not $p) { return }
        Invoke-LocalSearchLoop -SearchPath $p
        if (-not (Read-RunAgain 'search another folder')) { return }
    }
}

function Invoke-OpDownload {
    param([switch]$Batch)
    Show-OpIntro -Key $(if ($Batch) { '14' } else { '13' }) -Title $(if ($Batch) { 'BATCH DOWNLOAD - pick SharePoint items to download' } else { 'DOWNLOAD - a whole SharePoint / OneDrive folder to this PC' }) -Color Magenta -What @(
        $(if ($Batch) { 'Shows a SharePoint / OneDrive folder as a numbered list; the items you pick are downloaded with all their contents.' } else { 'Downloads a folder (all sub-folders and files) or a single file straight to a local or network folder - no OneDrive sync needed.' })
        'Already-downloaded files are skipped, so you can simply run it again to resume. Each file is size-checked and keeps its SharePoint date.'
    ) -Notes @('Names that Windows does not allow are adjusted automatically (the log lists every rename).')
    while ($true) {
        $src = Read-CloudFolder -Prompt 'SharePoint / OneDrive folder (or file) to download:' -LogLabel 'Source URL'
        if (-not $src) { return }
        $dst = Read-LocalFolder -Prompt 'Save into this folder on the PC / network (created if missing):' -LogLabel 'Destination path' -CreateIfMissing
        if (-not $dst) { return }
        $buf = Get-BufferFolder
        if (-not $buf) { return }
        try { Connect-CloudSite -SiteUrl $src.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; Write-Hint 'Check the address, your internet connection and that you signed in with the right account.'; continue }
        if ($Batch) { Perform-CloudToLocalBatchDownload -CloudData $src -LocalDest $dst -BufferDest $buf }
        else { $null = Perform-DirectDownload -CloudData $src -LocalDest $dst -BufferDest $buf }
        if (-not (Read-RunAgain 'download something else')) { return }
    }
}

function Invoke-OpUpload {
    param([switch]$Batch)
    Show-OpIntro -Key $(if ($Batch) { '16' } else { '15' }) -Title $(if ($Batch) { 'BATCH UPLOAD - pick local items to upload' } else { 'UPLOAD - a whole local folder to SharePoint / OneDrive' }) -Color Blue -What @(
        $(if ($Batch) { 'Shows a local folder as a numbered list; the items you pick are uploaded with all their contents.' } else { 'Uploads a folder (all sub-folders and files) or a single file into a SharePoint / OneDrive folder. Missing folders are created.' })
        'Files already in SharePoint with the same size are skipped, so you can run it again to resume.'
    ) -Notes @(
        'Between two sites or tenants: [13] Download from the first, then upload here to the second.'
        "Office lock files (~`$...), Thumbs.db and desktop.ini are skipped. Paths over $CloudPathLimit characters are skipped and listed in the LongPaths log."
    )
    while ($true) {
        $src = Read-LocalFolder -Prompt 'Local folder (or file) to upload:' -LogLabel 'Upload source path' -MustExist -AllowFile
        if (-not $src) { return }
        $dst = Read-CloudFolder -Prompt 'SharePoint / OneDrive folder to upload INTO:' -LogLabel 'Upload destination URL'
        if (-not $dst) { return }
        try { Connect-CloudSite -SiteUrl $dst.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; continue }
        if ($Batch) {
            if ([System.IO.File]::Exists((Get-LongPath $src))) { $null = Perform-DirectUpload -CloudDst $dst -LocalSrc $src }
            else { Perform-LocalToCloudBatchUpload -LocalSrc $src -CloudDst $dst }
        } else { $null = Perform-DirectUpload -CloudDst $dst -LocalSrc $src }
        if (-not (Read-RunAgain 'upload something else')) { return }
    }
}

function Invoke-OpArchive {
    param([switch]$Batch)
    Show-OpIntro -Key $(if ($Batch) { '18' } else { '17' }) -Title $(if ($Batch) { 'BATCH ARCHIVE - pick cloud items to archive and recycle' } else { 'ARCHIVE - download, verify, then recycle a cloud folder' }) -Color Red -What @(
        '1. DOWNLOAD everything to a local / network folder.   2. VERIFY every file on disk against a fresh SharePoint listing.   3. RECYCLE the cloud copy.'
        'Nothing is deleted unless step 2 passes completely. Re-running skips files already on disk, so an interrupted archive resumes.'
    ) -Notes @('Deleted items go to the SharePoint Recycle Bin (about 93 days). A library root is never deleted - its contents are.') -Danger 'You confirm once before anything starts; after that it runs unattended (press C to stop safely).'
    while ($true) {
        $src = Read-CloudFolder -Prompt 'SharePoint / OneDrive folder (or file) to archive:' -LogLabel 'Archive source URL'
        if (-not $src) { return }
        $dst = Read-LocalFolder -Prompt 'Save the archive into this folder (created if missing):' -LogLabel 'Archive destination path' -CreateIfMissing
        if (-not $dst) { return }
        $buf = Get-BufferFolder
        if (-not $buf) { return }
        try { Connect-CloudSite -SiteUrl $src.SiteUrl } catch { Write-Bad "Could not connect: $($_.Exception.Message)"; continue }
        if ($Batch) { $null = Perform-BatchArchiveDownload -CloudData $src -LocalDest $dst -BufferDest $buf }
        elseif (Confirm-Start -Title 'READY TO ARCHIVE' -Color Red -Danger -DangerText 'Download, verify and then recycle the cloud copy?' -Rows @(
                @{ Label = 'Cloud source'; Value = $src.FolderPath }, @{ Label = 'Save to'; Value = $dst }, 'If verification passes, the cloud copy goes to the SharePoint Recycle Bin.')) {
            Perform-ArchiveDownload -CloudData $src -LocalDest $dst -BufferDest $buf
        } else { Write-Hint 'Nothing was changed.' }
        if (-not (Read-RunAgain 'archive something else')) { return }
    }
}

# Builds a short, readable default label from a cloud address.
function Get-DefaultLabel {
    param([hashtable]$CloudData, [int]$Number)
    $segs = @($CloudData.FolderPath -split '/' | Where-Object { $_ })
    $name = if ($segs.Count -gt 0) { $segs[-1] } else { "Source$Number" }
    if ($name -in @('Documents', 'Shared Documents') -or $name -like '*.aspx') {
        $m = [regex]::Match($CloudData.FolderPath, '(?i)/personal/([^/]+)')
        if ($m.Success) {
            $parts = $m.Groups[1].Value -split '_'
            $name = if ($parts.Count -ge 3) { $parts[0..($parts.Count - 3)] -join '_' } else { $m.Groups[1].Value }
        }
    }
    return (Get-SafeWindowsName $name)
}

function Invoke-OpMultiArchive {
    Show-OpIntro -Key '19' -Title 'MULTI-SOURCE ARCHIVE - archive several cloud folders in one run' -Color Red -What @(
        'Paste the addresses of several SharePoint / OneDrive folders. Each one is archived into its own sub-folder (named by you) under one destination: download -> verify -> recycle.'
        'Sources on different sites are handled automatically (you may be asked to sign in once per site).'
    ) -Danger 'You confirm the plan once; then it runs unattended. A source that fails verification is never deleted.'
    $dst = Read-LocalFolder -Prompt 'Destination root folder (one sub-folder per source is created):' -LogLabel 'Multi-archive destination' -CreateIfMissing
    if (-not $dst) { return }
    $buf = Get-BufferFolder
    if (-not $buf) { return }
    $sources = New-Object System.Collections.Generic.List[hashtable]
    $labels = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Write-Host ""
    Write-Hint 'Paste one address per prompt. Press ENTER on an empty line when you are done, or Q to cancel.'
    while ($true) {
        $raw = ConvertTo-CleanPath (Read-Line -Prompt ("  Source {0} address > " -f ($sources.Count + 1)) -LogLabel 'Multi-archive source')
        if ($raw -match '^(?i)q$') { Write-Hint 'Cancelled.'; return }
        if (-not $raw) { break }
        $p = Parse-CloudUrl -Url $raw
        if (-not $p) { continue }
        if ($p.Note) { Write-Caution $p.Note }
        $def = Get-UniqueLocalName -Name (Get-DefaultLabel -CloudData $p -Number ($sources.Count + 1)) -Used (New-Object 'System.Collections.Generic.HashSet[string]' ($labels, [System.StringComparer]::OrdinalIgnoreCase))
        $lab = (Read-Line -Prompt "  Folder name for it (ENTER = '$def') > " -LogLabel 'Multi-archive label').Trim()
        $lab = if ($lab) { Get-SafeWindowsName $lab } else { $def }
        if (-not $labels.Add($lab)) { Write-Caution "'$lab' is already used by another source - pick a different name."; continue }
        $sources.Add(@{ CloudData = $p; SubFolder = $lab; Label = $lab })
        Write-Ok "Added '$lab'  <-  $($p.FolderPath)"
    }
    if ($sources.Count -eq 0) { Write-Hint 'No sources entered.'; return }
    $rows = @(@{ Label = 'Destination root'; Value = $dst }, '-') + @($sources | ForEach-Object { @{ Label = $_.Label; Value = $_.CloudData.FolderPath } })
    if (-not (Confirm-Start -Title ("ARCHIVE PLAN - {0} SOURCE(S)" -f $sources.Count) -Rows $rows -Color Red -Danger -DangerText 'Archive all of these (download, verify, recycle)?')) { Write-Hint 'Nothing was changed.'; return }
    Perform-MultiSourceArchive -Sources $sources.ToArray() -LocalDest $dst -BufferDest $buf
    Show-Pause
}

$script:LastTenant = ''
function Read-TenantName {
    if ($script:LastTenant) {
        $a = (Read-Line -Prompt "  Tenant name (ENTER = '$($script:LastTenant)', Q = cancel) > " -LogLabel 'Tenant').Trim()
        if (-not $a) { return $script:LastTenant }
    } else {
        Write-Hint "Your tenant name is the first part of your SharePoint address, e.g. 'contoso' from contoso.sharepoint.com."
        $a = (Read-Line -Prompt '  Tenant name (Q = cancel) > ' -LogLabel 'Tenant').Trim()
    }
    if (-not $a -or $a -match '^(?i)q$') { return '' }
    $t = ConvertTo-TenantName $a
    if ($t -notmatch '^[a-z0-9-]+$') { Write-Bad "'$a' does not look like a tenant name."; return '' }
    $script:LastTenant = $t
    return $t
}

# Turns a UPN or OneDrive address into a personal-site URL ('' when it cannot). Only
# OneDrive (/personal/) sites are accepted - a team site must never be archived by mistake.
function Resolve-LeaverSite {
    param([string]$Text)
    $t = ConvertTo-CleanPath $Text
    if ($t -match '^(?i)https?://') {
        $m = [regex]::Match($t, '(?i)(https://[^/]+/personal/[^/?#\s]+)')
        if ($m.Success) { return $m.Groups[1].Value.TrimEnd('/') }
        Write-Bad 'That address is not a OneDrive (it has no /personal/ part). Team sites cannot be archived with the leaver options.'
        return ''
    }
    if ($t -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $tenant = Read-TenantName
        if (-not $tenant) { Write-Hint 'A tenant name is needed to build the OneDrive address.'; return '' }
        return (ConvertTo-OneDriveSiteUrl -TenantName $tenant -Upn $t)
    }
    Write-Bad "That is not an email address or a OneDrive address."
    return ''
}

function Get-LeaverDefaultLabel {
    param([string]$SiteUrl, [int]$Number)
    $m = [regex]::Match($SiteUrl, '(?i)/personal/([^/]+)')
    if (-not $m.Success) { return "Leaver$Number" }
    $parts = $m.Groups[1].Value -split '_'
    $n = if ($parts.Count -ge 3) { $parts[0..($parts.Count - 3)] -join '_' } else { $m.Groups[1].Value }
    return (Get-SafeWindowsName $n)
}

function Invoke-OpLeaver {
    Show-OpIntro -Key '20' -Title "LEAVER ONEDRIVE ARCHIVE - archive one leaver's OneDrive" -Color DarkYellow -What @(
        "Connects to a leaver's OneDrive, then archives every document library: download -> verify -> recycle, item by item."
        'Documents library content lands directly in your destination; other libraries get their own sub-folder.'
    ) -Notes @("Enter the leaver's email (the OneDrive address is built for you) or paste any OneDrive address.", 'If their licence was removed and the OneDrive is frozen, use [23] instead.') -Danger 'You confirm once before it starts; after that it runs unattended (press C to stop safely).'
    while ($true) {
        $in = Read-PathInput -Prompt "Leaver's email address OR their OneDrive address:" -Type Any -LogLabel 'Leaver UPN or URL' -FreeText
        if ($in -eq 'Q') { return }
        $site = Resolve-LeaverSite $in
        if (-not $site) { continue }
        Write-Hint "OneDrive: $site"
        $dst = Read-LocalFolder -Prompt 'Save the archive into this folder (created if missing):' -LogLabel 'Leaver destination' -CreateIfMissing
        if (-not $dst) { return }
        $buf = Get-BufferFolder
        if (-not $buf) { return }
        if (Confirm-Start -Title 'READY TO ARCHIVE THIS ONEDRIVE' -Color DarkYellow -Danger -DangerText 'Archive every library and recycle the cloud copies that pass verification?' -Rows @(
                @{ Label = 'OneDrive'; Value = $site }, @{ Label = 'Save to'; Value = $dst })) {
            Invoke-LeaverArchiveSingle -SiteUrl $site -LocalDest $dst -BufferDest $buf
        } else { Write-Hint 'Nothing was changed.' }
        if (-not (Read-RunAgain 'archive another leaver')) { return }
    }
}

function Invoke-OpMultiLeaver {
    Show-OpIntro -Key '21' -Title 'MULTI-LEAVER ARCHIVE - queue several OneDrives and walk away' -Color DarkYellow -What @(
        'Enter several leavers (email or OneDrive address) with a folder name each, confirm the plan, and every OneDrive is archived in turn.'
        'A leaver that cannot be opened is skipped and reported; the others continue.'
    ) -Danger 'You confirm the plan once; then it runs unattended. Nothing that fails verification is deleted.'
    $dst = Read-LocalFolder -Prompt 'Destination root (one sub-folder per leaver):' -LogLabel 'Multi-leaver destination' -CreateIfMissing
    if (-not $dst) { return }
    $buf = Get-BufferFolder
    if (-not $buf) { return }
    $list = New-Object System.Collections.Generic.List[hashtable]
    $labels = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Write-Host ""
    Write-Hint 'One leaver per prompt (email or OneDrive address). ENTER on an empty line when done, Q to cancel.'
    while ($true) {
        $raw = (Read-Line -Prompt ("  Leaver {0} > " -f ($list.Count + 1)) -LogLabel 'Multi-leaver entry').Trim()
        if ($raw -match '^(?i)q$') { Write-Hint 'Cancelled.'; return }
        if (-not $raw) { break }
        $site = Resolve-LeaverSite $raw
        if (-not $site) { continue }
        $def = Get-UniqueLocalName -Name (Get-LeaverDefaultLabel -SiteUrl $site -Number ($list.Count + 1)) -Used (New-Object 'System.Collections.Generic.HashSet[string]' ($labels, [System.StringComparer]::OrdinalIgnoreCase))
        $lab = (Read-Line -Prompt "  Folder name (ENTER = '$def') > " -LogLabel 'Multi-leaver label').Trim()
        $lab = if ($lab) { Get-SafeWindowsName $lab } else { $def }
        if (-not $labels.Add($lab)) { Write-Caution "'$lab' is already used - pick a different name."; continue }
        $list.Add(@{ SiteUrl = $site; Label = $lab })
        Write-Ok "Added '$lab'  <-  $site"
    }
    if ($list.Count -eq 0) { Write-Hint 'No leavers entered.'; return }
    $rows = @(@{ Label = 'Destination root'; Value = $dst }, '-') + @($list | ForEach-Object { @{ Label = $_.Label; Value = ($_.SiteUrl -replace '^https://', '') } })
    if (-not (Confirm-Start -Title ("PLAN - {0} LEAVER(S)" -f $list.Count) -Rows $rows -Color DarkYellow -Danger -DangerText 'Start the unattended archive of all these OneDrives?')) { Write-Hint 'Nothing was changed.'; return }
    Perform-MultiSourceLeaverArchive -Leavers $list.ToArray() -LocalDest $dst -BufferDest $buf
    Show-Pause
}

function Invoke-OpRecycleLinks {
    Show-OpIntro -Key '22' -Title 'LEAVER RECYCLE BIN LINKS' -Color Cyan -What @(
        "Turns leaver emails or OneDrive addresses into links to each OneDrive's Recycle Bin, so you can review or empty them in the browser."
    ) -Notes @('In the browser: select all > Delete to empty a Recycle Bin permanently.')
    $links = New-Object System.Collections.Generic.List[hashtable]
    Write-Hint 'One per prompt (email or OneDrive address). ENTER on an empty line when done, Q to cancel.'
    while ($true) {
        $raw = (Read-Line -Prompt ("  Leaver {0} > " -f ($links.Count + 1)) -LogLabel 'Recycle bin entry').Trim()
        if ($raw -match '^(?i)q$') { return }
        if (-not $raw) { break }
        $site = Resolve-LeaverSite $raw
        if (-not $site) { continue }
        $rb = Get-LeaverRecycleBinUrl $site
        $links.Add(@{ Label = ($site -split '/')[-1]; Url = $rb })
        Write-Ok $rb
    }
    if ($links.Count -eq 0) { return }
    Write-RecycleBinLinks -Links $links.ToArray()
    Show-Pause
}

# Parses pasted leaver rows: Email <TAB> URL <TAB> Folder (or the 6-column usage report).
function ConvertFrom-LeaverRows {
    param([string[]]$Rows)
    $ok = New-Object System.Collections.Generic.List[hashtable]
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($row)) { continue }
        $cols = @($row -split "`t" | ForEach-Object { $_.Trim() })
        if ($cols.Count -lt 3) { $cols = @($row -split '[;,]' | ForEach-Object { $_.Trim() }) }
        if ($cols.Count -lt 3) { $errors.Add("Needs 3 columns (email, address, folder): $(Get-FitText -Text $row -Width 70)"); continue }
        $email = $cols[0]; $url = $cols[1]
        $folder = if ($cols.Count -ge 6) { $cols[5] } else { $cols[2] }
        if ($email -notmatch '@') { $errors.Add("Column 1 is not an email address: $email"); continue }
        $m = [regex]::Match($url, '(?i)(https://[^/]+/personal/([^/?#\s]+))')
        if (-not $m.Success) { $errors.Add("Column 2 is not a OneDrive (/personal/) address: $url"); continue }
        $url = $m.Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($folder)) { $errors.Add("Folder name is empty for $email"); continue }
        $expected = ($email.ToLowerInvariant()) -replace '[^a-z0-9_]', '_'
        if ($m.Groups[2].Value.ToLowerInvariant() -ne $expected) { $warnings.Add("Check: the OneDrive address for $email is /personal/$($m.Groups[2].Value) (expected /personal/$expected).") }
        $folder = Get-SafeWindowsName $folder
        if (-not $seen.Add($folder)) { $errors.Add("Folder name '$folder' is used twice - each leaver needs their own folder ($email skipped)."); continue }
        $ok.Add(@{ Email = $email; SiteUrl = $url; FolderName = $folder })
    }
    return @{ Leavers = $ok.ToArray(); Errors = $errors.ToArray(); Warnings = $warnings.ToArray() }
}

function Invoke-OpLicensedLeaver {
    Show-OpIntro -Key '23' -Title 'UNLICENSED LEAVER ARCHIVE - temporarily licence, archive, un-licence' -Color Cyan -What @(
        'For leavers whose licence was removed (their OneDrive is frozen). For each leaver, one at a time:'
        '  add to your licence group -> wait until the OneDrive opens -> archive (download, verify, recycle) -> remove from the group.'
        'Only one licence is used at any moment; the user is removed from the group even if something fails or you cancel.'
    ) -Notes @(
        'Needs: a security group with a SharePoint licence assigned, and rights to change its members (you sign in to Microsoft Graph - no extra module needed).'
        'Leaver list: paste rows "Email <TAB> OneDrive address <TAB> Folder name" (or the 6-column OneDrive usage report), or press V to take them from the clipboard.'
    )
    $dst = Read-LocalFolder -Prompt 'Step 1 of 4 - destination folder (one sub-folder per leaver):' -LogLabel 'Licensed destination' -CreateIfMissing
    if (-not $dst) { return }
    $buf = Get-BufferFolder
    if (-not $buf) { return }

    if (-not (Ensure-PnPModule)) { return }
    Write-Host ""; Write-Rule -Title 'Step 2 of 4 - sign in to Microsoft Graph' -Color Cyan
    if (-not (Ensure-MgGraphForLicensing)) { Write-Bad 'Microsoft Graph sign-in did not complete.'; return }

    Write-Host ""; Write-Rule -Title 'Step 3 of 4 - licence group' -Color Cyan
    $group = $null
    while (-not $group) {
        $grpInput = (Read-Line -Prompt '  Group display name or object ID (Q = cancel) > ' -LogLabel 'Licence group').Trim()
        if ($grpInput -match '^(?i)q$') { return }
        if (-not $grpInput) { Write-Caution 'Please type the group name (or Q to cancel).'; continue }
        try {
            if ($grpInput -match '^[0-9a-fA-F-]{36}$') { $group = Invoke-MgRestMethod -Method GET -Endpoint "groups/$grpInput`?`$select=id,displayName"; break }
            $f = [uri]::EscapeDataString("displayName eq '$($grpInput -replace "'", "''")'")
            $found = @((Invoke-MgRestMethod -Method GET -Endpoint "groups?`$filter=$f&`$select=id,displayName,description").value)
            if ($found.Count -eq 0) { Write-Bad "No group called '$grpInput' was found. Check the name in Entra ID (or paste the group's object ID)."; continue }
            if ($found.Count -eq 1) { $group = $found[0]; break }
            Write-Caution "$($found.Count) groups have that name - pick the right one:"
            for ($i = 0; $i -lt $found.Count; $i++) { Write-Host (Get-FitText -Text ("  {0}. {1}  ({2})  {3}" -f ($i + 1), $found[$i].displayName, $found[$i].id, $found[$i].description) -Width (Get-UiWidth)) -ForegroundColor Gray }
            $n = (Read-Line -Prompt '  Number > ' -LogLabel 'Group pick').Trim()
            if ($n -match '^\d{1,6}$' -and [int]$n -ge 1 -and [int]$n -le $found.Count) { $group = $found[[int]$n - 1] }
            else { Write-Caution 'Not one of the listed numbers - please enter the group again.' }
        } catch { Write-Bad "Group look-up failed: $($_.Exception.Message)" }
    }
    Write-Ok "Licence group: $($group.displayName) ($($group.id))"

    Write-Host ""; Write-Rule -Title 'Step 4 of 4 - leavers' -Color Cyan
    $parsed = $null
    while ($true) {
        $rows = New-Object System.Collections.Generic.List[string]
        Write-Hint 'Paste rows (ENTER on an empty line when done), V = read the clipboard, Q = cancel.'
        while ($true) {
            $r = Read-Line -Prompt ("  Row {0} > " -f ($rows.Count + 1)) -LogLabel 'Leaver row'
            if ($r.Trim() -match '^(?i)q$') { return }
            if ($r.Trim() -match '^(?i)v$') {
                try { $clip = @(Get-Clipboard -ErrorAction Stop) } catch { Write-Bad "Clipboard could not be read: $($_.Exception.Message)"; continue }
                $clipRows = @(($clip -join "`n") -split "`r?`n" | Where-Object { $_.Trim() })
                if ($clipRows.Count -eq 0) { Write-Caution 'The clipboard is empty.'; continue }
                Write-Info "$($clipRows.Count) row(s) on the clipboard:"
                foreach ($cr in ($clipRows | Select-Object -First 10)) { Write-Hint (Get-FitText -Text $cr -Width ((Get-UiWidth) - 10)) }
                if ($clipRows.Count -gt 10) { Write-Hint "...and $($clipRows.Count - 10) more" }
                if (Read-YesNo -Question 'Use these rows?' -Default $true) { foreach ($cr in $clipRows) { $rows.Add($cr) }; break }
                continue
            }
            if (-not $r.Trim()) { if ($rows.Count -gt 0) { break } else { Write-Caution 'Nothing entered yet.'; continue } }
            $rows.Add($r)
        }
        $parsed = ConvertFrom-LeaverRows -Rows $rows.ToArray()
        foreach ($e in $parsed.Errors) { Write-Caution $e }
        foreach ($w in $parsed.Warnings) { Write-Caution $w }
        if ($parsed.Leavers.Count -gt 0) { break }
        Write-Bad 'No valid rows - please try again (Email <TAB> OneDrive address <TAB> Folder name).'
    }
    $plan = @(
        @{ Label = 'Licence group'; Value = $group.displayName }
        @{ Label = 'Destination'; Value = $dst }
        '-'
    ) + @($parsed.Leavers | ForEach-Object { @{ Label = $_.FolderName; Value = "$($_.Email)  |  $($_.SiteUrl -replace '^https://', '')" } })
    if ($parsed.Warnings.Count -gt 0) { $plan += '-'; foreach ($w in $parsed.Warnings) { $plan += @{ Text = $w; Color = 'Yellow' } } }
    if (-not (Confirm-Start -Title ("PLAN - {0} LEAVER(S)" -f $parsed.Leavers.Count) -Rows $plan -Color Cyan -Danger -DangerText 'Start: licence -> archive -> un-licence, one leaver at a time?')) { Write-Hint 'Nothing was changed.'; return }
    Perform-LicensedLeaverArchive -Leavers $parsed.Leavers -LocalDest $dst -BufferDest $buf -GroupId $group.id -GroupName $group.displayName
    Show-Pause
}

# ======================================================================================
# SECTION 17 - UTILITY SCREENS (help, settings, history, guided mode)
# ======================================================================================

function Show-HelpScreen {
    Write-Banner -Title 'QUICK REFERENCE' -Color White
    Write-Panel -Title 'AT ANY PATH OR ADDRESS PROMPT' -Color Cyan -Lines @(
        @{ Label = 'Paste / type'; Value = 'C:\Folder   D:\Archive\2024   \\server\share\folder   https://...sharepoint.com/...' }
        @{ Label = 'Drag and drop'; Value = 'Drag a folder from File Explorer into this window - quotes are removed for you.' }
        @{ Label = 'F'; Value = 'Pick one of your saved favourites.' }
        @{ Label = 'R'; Value = 'Re-use a path from earlier in this session.' }
        @{ Label = 'B'; Value = 'Browse for a folder with the Windows folder picker.' }
        @{ Label = 'Q'; Value = 'Cancel and go back.' }
    )
    Write-Panel -Title 'WHILE A JOB IS RUNNING' -Color Yellow -Lines @(
        @{ Label = 'P'; Value = 'Pause (then any key resumes, C cancels).' }
        @{ Label = 'C or Ctrl+C'; Value = 'Cancel cleanly: the current file finishes, the summary is shown, nothing is half-done.' }
        'Hotkeys need the normal PowerShell console (the ISE cannot read single key presses).'
    )
    Write-Panel -Title 'PICKING ITEMS IN A LIST' -Color Green -Lines @(
        @{ Label = '3'; Value = 'item 3' }
        @{ Label = '2-5'; Value = 'items 2, 3, 4 and 5 (10-2 works too)' }
        @{ Label = '1,3,7   or   1 3 7'; Value = 'items 1, 3 and 7' }
        @{ Label = 'A'; Value = 'everything' }
        @{ Label = 'S / ENTER / Q'; Value = 'search / refresh the list / go back' }
    )
    Write-Panel -Title 'SEARCH PATTERNS' -Color Magenta -Lines @(
        @{ Label = 'invoice'; Value = 'any name containing "invoice"' }
        @{ Label = '*.pdf'; Value = 'all PDF files' }
        @{ Label = '*2024*report*'; Value = 'names containing 2024 followed by report' }
        @{ Label = 'Budget [Final].xlsx'; Value = 'exact name - square brackets are literal' }
    )
    Write-Panel -Title 'GOOD TO KNOW' -Color DarkCyan -Lines @(
        'Every job can be re-run safely: identical files are skipped, so interrupted jobs resume.'
        'Cloud copy / move ([1]-[3]) run on Microsoft''s servers. Between tenants: [13] download, then [15] upload.'
        'Archive options only recycle cloud data after every file has been checked on disk. Recycle Bin items stay for ~93 days.'
        "Everything is logged to $LogFile (open the folder with L on the main menu)."
    )
    Show-Pause
}

function Show-SettingsScreen {
    while ($true) {
        $s = $script:Settings
        $bufText = if ($s.DefaultBuffer) { $s.DefaultBuffer } else { "Windows temp folder ($([System.IO.Path]::GetTempPath().TrimEnd('\')))" }
        Write-Banner -Title 'SETTINGS' -Color Cyan -Subtitle "Saved in $($script:SettingsFile)"
        Write-Panel -Color Cyan -Lines @(
            @{ Label = '1  Display style'; Value = $(if ($s.Theme -eq 'Ascii') { 'Plain ASCII (+-|)' } else { 'Box drawing (recommended)' }) }
            @{ Label = '2  Explanations'; Value = $(if ($s.ShowGuides) { 'Shown when an option opens' } else { 'Hidden (header only)' }) }
            @{ Label = '3  Console detail'; Value = $(if ($s.ConsoleDetail -eq 'Compact') { 'Compact - errors and summaries only (faster)' } else { 'Normal - a line per file' }) }
            @{ Label = '4  Completion chime'; Value = $(if ($s.Sound) { 'On (jobs over 30 seconds)' } else { 'Off' }) }
            @{ Label = '5  Temp download buffer'; Value = $bufText }
            @{ Label = '6  Large-file progress'; Value = "Files of $($s.LargeFileMB) MB or more get their own progress bar" }
            @{ Label = '7  Reset to defaults'; Value = '' }
        )
        $c = (Read-Line -Prompt '  Number to change (ENTER = back) > ' -LogLabel 'Settings').Trim()
        switch ($c) {
            '1' { $s.Theme = $(if ($s.Theme -eq 'Ascii') { 'Unicode' } else { 'Ascii' }); Initialize-Theme -Style $s.Theme }
            '2' { $s.ShowGuides = -not $s.ShowGuides }
            '3' { $s.ConsoleDetail = $(if ($s.ConsoleDetail -eq 'Compact') { 'Normal' } else { 'Compact' }) }
            '4' { $s.Sound = -not $s.Sound; if ($s.Sound) { try { [Console]::Beep(784, 120); [Console]::Beep(1046, 180) } catch { } } }
            '5' {
                $b = Read-LocalFolder -Prompt 'Temp buffer folder (ENTER = Windows temp folder):' -LogLabel 'Default buffer' -AllowEmpty -EmptyHint 'use the Windows temp folder' -CreateIfMissing
                if ($null -ne $b) {
                    if ($b -and -not (Test-WritableFolder $b)) { Write-Bad 'That folder is not writable.' } else { $s.DefaultBuffer = $b }
                }
            }
            '6' {
                $v = (Read-Line -Prompt '  Size in MB (50 or more) > ' -LogLabel 'Large file MB').Trim()
                if ($v -match '^\d{1,6}$' -and [int]$v -ge 50) { $s.LargeFileMB = [int]$v } else { Write-Caution 'Please enter a whole number of 50 or more.' }
            }
            '7' { if (Read-YesNo -Question 'Reset every setting to its default?' -Default $false) { foreach ($k in $script:DefaultSettings.Keys) { $s[$k] = $script:DefaultSettings[$k] }; Initialize-Theme -Style $s.Theme } }
            '' { return }
            default { Write-Caution 'Type a number from 1 to 7, or press ENTER.' }
        }
        Save-AppSettings
    }
}

function Show-JobHistory {
    while ($true) {
        $hist = (Get-JobHistory)
        Write-Banner -Title 'JOB HISTORY' -Right "$($hist.Count) job(s) recorded" -Color Cyan
        if ($hist.Count -eq 0) { Write-Hint 'No jobs yet - finished jobs appear here.'; Show-Pause; return }
        [long]$bytes = 0; $files = 0; $secs = 0; $clean = 0
        foreach ($h in $hist) { $bytes += [long]$h.Bytes; $files += [int]$h.Ok; $secs += [int]$h.Seconds; if ($h.Outcome -eq 'Success') { $clean++ } }
        Write-Panel -Title 'ALL TIME' -Color Cyan -Lines @(
            @{ Label = 'Jobs'; Value = ("{0} ({1} clean)" -f $hist.Count, $clean) }
            @{ Label = 'Files handled'; Value = (Format-Count $files) }
            @{ Label = 'Data moved'; Value = (Format-FileSize $bytes); Level = 'SUCCESS' }
            @{ Label = 'Time spent working'; Value = (Format-Duration $secs) }
        )
        $W = Get-UiWidth
        $outW = if ($W -ge 90) { 23 } else { 12 }
        $opW = [Math]::Max(8, $W - 2 - 17 - ($outW + 1) - 20)
        Write-Host ("  {0} {1} {2} {3} {4}" -f 'WHEN'.PadRight(16), 'OPERATION'.PadRight($opW), 'OUTCOME'.PadRight($outW), 'FILES'.PadLeft(9), 'DATA'.PadLeft(9)) -ForegroundColor DarkGray
        foreach ($h in @($hist | Select-Object -Last 20)) {
            $col = switch -Regex ($h.Outcome) { '^Success' { 'Green' } '^Cancel' { 'Yellow' } default { 'Red' } }
            Write-Host ("  {0} " -f "$($h.When)".PadRight(16)) -ForegroundColor Gray -NoNewline
            Write-Host ((Get-FitText -Text "$($h.Operation)" -Width $opW).PadRight($opW) + ' ') -ForegroundColor White -NoNewline
            Write-Host ((Get-FitText -Text "$($h.Outcome)" -Width $outW).PadRight($outW) + ' ') -ForegroundColor $col -NoNewline
            Write-Host ((Format-Count ([long]$h.Ok)).PadLeft(9) + ' ' + (Format-FileSize ([long]$h.Bytes)).PadLeft(9)) -ForegroundColor Gray
        }
        $a = Read-Choice -Prompt 'E = export all to CSV   C = clear history   ENTER = back' -Valid @('E', 'C')
        if (-not $a) { return }
        if ($a -eq 'E') {
            $csv = Join-Path $LogFileDirectory ("JobHistory_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            try { $hist | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8; Write-Ok "Saved: $csv" } catch { Write-Bad $_.Exception.Message }
            Show-Pause -Message 'Press any key to continue...'
        }
        if ($a -eq 'C' -and (Read-Confirmation -Message 'Clear the whole job history? (Logs are not affected.)')) { Save-JobHistory -Entries @(); Write-Ok 'History cleared.' }
    }
}

# Two questions -> the right option.
function Start-GuidedMode {
    Write-Banner -Title 'GUIDED MODE' -Color Magenta -Subtitle 'Answer two quick questions and the right option opens for you.'
    Write-Host ""
    Write-Host "  Where are the files now?" -ForegroundColor White
    Write-Host "    1  In SharePoint or OneDrive (in the cloud)" -ForegroundColor Gray
    Write-Host "    2  On this PC, a USB drive or a network share" -ForegroundColor Gray
    Write-Host "    3  In a leaver's OneDrive" -ForegroundColor Gray
    $where = (Read-Line -Prompt '  1, 2 or 3 > ' -LogLabel 'Guided: where').Trim()
    $key = ''
    switch ($where) {
        '1' {
            Write-Host ""
            Write-Host "  What would you like to do?" -ForegroundColor White
            $opts = @(
                @('1', 'Copy them to another SharePoint / OneDrive folder', '1'), @('2', 'Move them to another SharePoint / OneDrive folder', '2'),
                @('3', 'Download them to this PC / network', '13'), @('4', 'Download them, then delete them from the cloud (archive)', '17'),
                @('5', 'Check that two cloud folders match', '5'), @('6', 'Find a file', '6'), @('7', 'Delete a folder (Recycle Bin)', '4')
            )
            foreach ($o in $opts) { Write-Host ("    {0}  {1}" -f $o[0], $o[1]) -ForegroundColor Gray }
            $a = (Read-Line -Prompt '  Number > ' -LogLabel 'Guided: what').Trim()
            foreach ($o in $opts) { if ($o[0] -eq $a) { $key = $o[2] } }
            if ($key -in @('1', '2', '13', '17') -and (Read-YesNo -Question 'Only some items from the folder (pick from a list)?' -Default $false)) {
                $key = switch ($key) { '1' { '3' } '2' { '3' } '13' { '14' } '17' { '18' } }
            }
        }
        '2' {
            Write-Host ""
            Write-Host "  What would you like to do?" -ForegroundColor White
            $opts = @(
                @('1', 'Copy them to another folder', '7'), @('2', 'Move them to another folder', '8'), @('3', 'Upload them to SharePoint / OneDrive', '15'),
                @('4', 'Check that two folders match', '11'), @('5', 'Find a file', '12'), @('6', 'Delete a folder permanently', '10')
            )
            foreach ($o in $opts) { Write-Host ("    {0}  {1}" -f $o[0], $o[1]) -ForegroundColor Gray }
            $a = (Read-Line -Prompt '  Number > ' -LogLabel 'Guided: what').Trim()
            foreach ($o in $opts) { if ($o[0] -eq $a) { $key = $o[2] } }
            if ($key -in @('7', '8', '15') -and (Read-YesNo -Question 'Only some items from the folder (pick from a list)?' -Default $false)) {
                $key = switch ($key) { '7' { '9' } '8' { '9' } '15' { '16' } }
            }
        }
        '3' {
            Write-Host ""
            $opts = @(
                @('1', 'Archive one leaver (account still licensed)', '20'), @('2', 'Archive several leavers in one run', '21'),
                @('3', 'Their licence was removed - OneDrive is frozen', '23'), @('4', 'Get Recycle Bin links after an archive', '22')
            )
            foreach ($o in $opts) { Write-Host ("    {0}  {1}" -f $o[0], $o[1]) -ForegroundColor Gray }
            $a = (Read-Line -Prompt '  Number > ' -LogLabel 'Guided: leaver').Trim()
            foreach ($o in $opts) { if ($o[0] -eq $a) { $key = $o[2] } }
        }
    }
    if (-not $key) { Write-Hint 'No problem - back to the main menu.'; return }
    $entry = $script:MenuItems | Where-Object { $_.Key -eq $key } | Select-Object -First 1
    Write-Ok ("Recommended: [{0}] {1}" -f $entry.Key, $entry.Text)
    if (Read-YesNo -Question 'Open it now?' -Default $true) { Invoke-MenuEntry -Entry $entry }
}

# ======================================================================================
# SECTION 18 - MAIN MENU
# ======================================================================================

$script:MenuItems = @(
    @{ Key = '1';  Group = 'CLOUD';    Text = 'Copy a folder';                Words = 'cloud copy sharepoint onedrive server side';        Run = { Invoke-OpCloudCopyMove } }
    @{ Key = '2';  Group = 'CLOUD';    Text = 'Move a folder';                Words = 'cloud move sharepoint onedrive server side';        Run = { Invoke-OpCloudCopyMove -Move } }
    @{ Key = '3';  Group = 'CLOUD';    Text = 'Batch: pick items';            Words = 'cloud batch pick select copy move delete verify';   Run = { Invoke-OpCloudBatch } }
    @{ Key = '4';  Group = 'CLOUD';    Text = 'Wipe a folder (Recycle Bin)';  Words = 'cloud wipe delete remove recycle';                  Run = { Invoke-OpCloudWipe } }
    @{ Key = '5';  Group = 'CLOUD';    Text = 'Verify two folders match';     Words = 'cloud verify compare check match';                  Run = { Invoke-OpCloudVerify } }
    @{ Key = '6';  Group = 'CLOUD';    Text = 'Search by name';               Words = 'cloud search find';                                 Run = { Invoke-OpCloudSearch } }
    @{ Key = '7';  Group = 'LOCAL';    Text = 'Copy a folder';                Words = 'local copy pc network unc drive';                   Run = { Invoke-OpLocalCopyMove } }
    @{ Key = '8';  Group = 'LOCAL';    Text = 'Move a folder';                Words = 'local move pc network unc drive';                   Run = { Invoke-OpLocalCopyMove -Move } }
    @{ Key = '9';  Group = 'LOCAL';    Text = 'Batch: pick items';            Words = 'local batch pick select copy move delete verify';   Run = { Invoke-OpLocalBatch } }
    @{ Key = '10'; Group = 'LOCAL';    Text = 'Wipe a folder (permanent!)';   Words = 'local wipe delete remove permanent';                Run = { Invoke-OpLocalWipe } }
    @{ Key = '11'; Group = 'LOCAL';    Text = 'Verify two folders match';     Words = 'local verify compare check match hash checksum';    Run = { Invoke-OpLocalVerify } }
    @{ Key = '12'; Group = 'LOCAL';    Text = 'Search by name';               Words = 'local search find';                                 Run = { Invoke-OpLocalSearch } }
    @{ Key = '13'; Group = 'TRANSFER'; Text = 'Download a folder';            Words = 'download cloud to pc save';                         Run = { Invoke-OpDownload } }
    @{ Key = '14'; Group = 'TRANSFER'; Text = 'Download: pick items';         Words = 'download batch pick select';                        Run = { Invoke-OpDownload -Batch } }
    @{ Key = '15'; Group = 'TRANSFER'; Text = 'Upload a folder';              Words = 'upload pc to cloud';                                Run = { Invoke-OpUpload } }
    @{ Key = '16'; Group = 'TRANSFER'; Text = 'Upload: pick items';           Words = 'upload batch pick select';                          Run = { Invoke-OpUpload -Batch } }
    @{ Key = '17'; Group = 'ARCHIVE';  Text = 'Archive one folder';           Words = 'archive download verify delete recycle reclaim';    Run = { Invoke-OpArchive } }
    @{ Key = '18'; Group = 'ARCHIVE';  Text = 'Archive: pick items';          Words = 'archive batch pick select';                         Run = { Invoke-OpArchive -Batch } }
    @{ Key = '19'; Group = 'ARCHIVE';  Text = 'Archive many folders';         Words = 'archive multi multiple many sources';               Run = { Invoke-OpMultiArchive } }
    @{ Key = '20'; Group = 'LEAVERS';  Text = "Archive a leaver's OneDrive";  Words = 'leaver onedrive archive user personal';             Run = { Invoke-OpLeaver } }
    @{ Key = '21'; Group = 'LEAVERS';  Text = 'Archive many leavers';         Words = 'leaver multi multiple many onedrive';               Run = { Invoke-OpMultiLeaver } }
    @{ Key = '22'; Group = 'LEAVERS';  Text = 'Recycle Bin links';            Words = 'leaver recycle bin links empty';                    Run = { Invoke-OpRecycleLinks } }
    @{ Key = '23'; Group = 'LEAVERS';  Text = 'Unlicensed leaver archive';    Words = 'leaver licence license unlicensed frozen group graph'; Run = { Invoke-OpLicensedLeaver } }
)

$script:MenuGroups = @{
    CLOUD    = @{ Title = 'CLOUD  (same tenant, server-side)'; Color = 'Yellow' }
    LOCAL    = @{ Title = 'LOCAL  (PC, USB, network share)';   Color = 'Green' }
    TRANSFER = @{ Title = 'DOWNLOAD / UPLOAD';                 Color = 'Magenta' }
    ARCHIVE  = @{ Title = 'ARCHIVE  (download, verify, recycle)'; Color = 'Red' }
    LEAVERS  = @{ Title = 'LEAVER ONEDRIVES';                  Color = 'DarkYellow' }
}

$script:Tips = @(
    'Drag a folder from File Explorer into this window to paste its path.'
    'Every job can simply be run again - identical files are skipped, so interrupted jobs resume.'
    'Press P to pause or C to cancel any running job. Ctrl+C cancels cleanly too.'
    'Save paths and SharePoint addresses you use often with F (Favourites).'
    'Not sure which option you need? Press W for guided mode.'
    'Type a word such as "download" or "verify" at the menu to find the right option.'
    'J shows your job history and how much data you have moved so far.'
    'Prefer a calmer screen? S > Console detail: Compact shows only errors and summaries.'
    'Archive options never delete anything from SharePoint until every file is checked on disk.'
)

function Invoke-MenuEntry {
    param([hashtable]$Entry)
    Write-Log "MENU: [$($Entry.Key)] $($Entry.Group) - $($Entry.Text)" 'INFO' -FileOnly
    try {
        & $Entry.Run
    } catch {
        Write-Host ""
        Write-Bad "Something unexpected went wrong: $($_.Exception.Message)"
        Write-Hint "The details are in the log. You are back at the main menu - nothing else was changed."
        Write-Log ("UNEXPECTED ERROR in [{0}]: {1}`r`n{2}" -f $Entry.Key, $_.Exception.Message, $_.ScriptStackTrace) 'ERROR' -FileOnly
        Show-Pause
    } finally {
        $script:JobDepth = 0
        $script:JobCancelled = $false
        Set-CtrlCAsInput $false
        Complete-AllProgress
    }
}

function Get-MenuColumnLines {
    param([string[]]$Groups)
    $lines = New-Object System.Collections.Generic.List[hashtable]
    foreach ($grp in $Groups) {
        if ($lines.Count -gt 0) { $lines.Add(@{ Kind = 'blank' }) }
        $lines.Add(@{ Kind = 'head'; Text = $script:MenuGroups[$grp].Title; Color = $script:MenuGroups[$grp].Color })
        foreach ($m in ($script:MenuItems | Where-Object { $_.Group -eq $grp })) { $lines.Add(@{ Kind = 'item'; Key = $m.Key; Text = $m.Text; Color = $script:MenuGroups[$grp].Color }) }
    }
    return ,$lines.ToArray()
}

function Write-MenuCell {
    param([hashtable]$Cell, [int]$Width)
    if (-not $Cell -or $Cell.Kind -eq 'blank') { Write-Host (' ' * $Width) -NoNewline; return }
    if ($Cell.Kind -eq 'head') { Write-Host ((Get-FitText -Text $Cell.Text -Width $Width).PadRight($Width)) -ForegroundColor $Cell.Color -NoNewline; return }
    $k = ("{0,3}" -f $Cell.Key)
    Write-Host ($k + '  ') -ForegroundColor Cyan -NoNewline
    Write-Host ((Get-FitText -Text $Cell.Text -Width ($Width - 5)).PadRight($Width - 5)) -ForegroundColor White -NoNewline
}

function Show-MainMenu {
    $W = Get-UiWidth
    $hour = (Get-Date).Hour
    $greet = if ($hour -lt 12) { 'Good morning' } elseif ($hour -lt 18) { 'Good afternoon' } else { 'Good evening' }
    Write-Banner -Title ("{0} {1}" -f $script:AppName.ToUpper(), $script:AppVersion) -Right "$greet, $ActiveUser" -Color Cyan
    $favCount = (Get-Favorites).Count
    $sess = if ($script:Session.Jobs -gt 0) { "{0} job(s), {1} moved" -f $script:Session.Jobs, (Format-FileSize $script:Session.Bytes) } else { 'no jobs yet' }
    $conn = if ($script:PnPSiteUrl) { "connected to $($script:PnPSiteUrl -replace '^https://', '')" } else { 'not connected' }
    Write-Hint ("This session: {0}   |   Favourites: {1}   |   Cloud: {2}" -f $sess, $favCount, $conn)
    Write-Host ""
    $left  = Get-MenuColumnLines -Groups @('CLOUD', 'TRANSFER', 'LEAVERS')
    $right = Get-MenuColumnLines -Groups @('LOCAL', 'ARCHIVE')
    if ($W -ge 78) {
        $colW = [int](($W - 4) / 2)
        $rows = [Math]::Max($left.Count, $right.Count)
        for ($i = 0; $i -lt $rows; $i++) {
            Write-Host '  ' -NoNewline
            Write-MenuCell -Cell $(if ($i -lt $left.Count) { $left[$i] } else { $null }) -Width $colW
            Write-Host ' ' -NoNewline
            Write-MenuCell -Cell $(if ($i -lt $right.Count) { $right[$i] } else { $null }) -Width ($colW - 1)
            Write-Host ''
        }
    } else {
        foreach ($c in (@($left) + @(@{ Kind = 'blank' }) + @($right))) { Write-Host '  ' -NoNewline; Write-MenuCell -Cell $c -Width ($W - 3); Write-Host '' }
    }
    Write-Host ""
    Write-Rule -Color DarkGray
    $tools = @(@('W', 'Guided'), @('F', 'Favourites'), @('J', 'History'), @('S', 'Settings'), @('H', 'Help'), @('L', 'Logs'), @('Q', 'Quit'))
    $rowsOut = New-Object System.Collections.Generic.List[object]
    $cur = New-Object System.Collections.Generic.List[object]; $len = 2
    foreach ($t in $tools) {
        $w = 2 + $t[1].Length
        $need = if ($cur.Count -gt 0) { 3 + $w } else { $w }
        if ($cur.Count -gt 0 -and ($len + $need) -gt $W) { $rowsOut.Add($cur.ToArray()); $cur = New-Object System.Collections.Generic.List[object]; $len = 2; $need = $w }
        $cur.Add($t); $len += $need
    }
    if ($cur.Count -gt 0) { $rowsOut.Add($cur.ToArray()) }
    foreach ($r in $rowsOut) {
        Write-Host '  ' -NoNewline
        for ($i = 0; $i -lt $r.Count; $i++) {
            if ($i -gt 0) { Write-Host '   ' -NoNewline }
            Write-Host $r[$i][0] -ForegroundColor Cyan -NoNewline
            Write-Host (' ' + $r[$i][1]) -ForegroundColor Gray -NoNewline
        }
        Write-Host ''
    }
    $tip = $script:Tips[(Get-Random -Minimum 0 -Maximum $script:Tips.Count)]
    Write-Host ("  Tip: " + (Get-FitText -Text $tip -Width ($W - 8))) -ForegroundColor DarkGray
}

# Finds options whose number, title or keywords match what was typed.
function Find-MenuEntries {
    param([string]$Text)
    $words = @($Text.ToLowerInvariant() -split '\s+' | Where-Object { $_ })
    return @($script:MenuItems | Where-Object {
            $hay = ("{0} {1} {2}" -f $_.Group, $_.Text, $_.Words).ToLowerInvariant()
            $all = $true
            foreach ($w in $words) { if ($hay -notmatch [regex]::Escape($w)) { $all = $false; break } }
            $all
        })
}

function Show-SessionSummary {
    $dur = (Get-Date) - $script:Session.Started
    Write-Host ""
    Write-SummaryBox -Title 'SESSION SUMMARY' -BoxColor Cyan -Lines @(
        @{ Label = 'Jobs run'; Value = "$($script:Session.Jobs)" }
        @{ Label = 'Files handled'; Value = (Format-Count $script:Session.Files) }
        @{ Label = 'Data moved'; Value = (Format-FileSize $script:Session.Bytes); Level = 'SUCCESS' }
        @{ Label = 'Failures'; Value = (Format-Count $script:Session.Failed); Level = $(if ($script:Session.Failed) { 'WARN' } else { 'SUCCESS' }) }
        @{ Label = 'Session length'; Value = (Format-Elapsed $dur) }
        @{ Label = 'Log file'; Value = $LogFile }
    )
}

function Start-MainMenu {
    Import-AppSettings
    Set-AppTitle
    Write-Log "===== SESSION START: $($script:AppName) $($script:AppVersion) | user $ActiveUser | PowerShell $($PSVersionTable.PSVersion) | host $($Host.Name) =====" 'INFO' -FileOnly
    if ($PSVersionTable.PSVersion.Major -gt 5) {
        Write-Caution "You are running PowerShell $($PSVersionTable.PSVersion). This tool is built and tested for Windows PowerShell 5.1."
    }
    $keyMap = @{}
    foreach ($m in $script:MenuItems) { $keyMap[$m.Key] = $m }
    try {
        while ($true) {
            Show-MainMenu
            $choice = (Read-Line -Prompt '  Choose an option > ' -LogLabel 'Main menu').Trim()
            if (-not $choice) { continue }
            $upper = $choice.ToUpperInvariant()
            if ($keyMap.ContainsKey($choice)) { Invoke-MenuEntry -Entry $keyMap[$choice]; continue }
            if ($upper -eq 'Q') {
                Show-SessionSummary
                Write-Host ""
                Write-Host "  Goodbye, $ActiveUser - see you next time!" -ForegroundColor Cyan
                Write-Host ""
                return
            }
            if ($upper -eq 'F') { Manage-Favorites; continue }
            if ($upper -eq 'J') { Show-JobHistory; continue }
            if ($upper -eq 'S') { Show-SettingsScreen; continue }
            if ($upper -eq 'H' -or $upper -eq '?') { Show-HelpScreen; continue }
            if ($upper -eq 'W') { Start-GuidedMode; continue }
            if ($upper -eq 'L') {
                if (Show-InExplorer -Path $LogFileDirectory -OpenFolder) { Write-Ok "Opened $LogFileDirectory" } else { Write-Hint "Logs are in: $LogFileDirectory" }
                continue
            }
            $hits = @(Find-MenuEntries -Text $choice)
            if ($hits.Count -eq 0) { Write-Caution "'$choice' is not an option. Type a number (1-23), W for guided mode, or H for help."; Start-Sleep -Milliseconds 900; continue }
            if ($hits.Count -eq 1) {
                if (Read-YesNo -Question ("Open [{0}] {1} - {2}?" -f $hits[0].Key, $script:MenuGroups[$hits[0].Group].Title, $hits[0].Text) -Default $true) { Invoke-MenuEntry -Entry $hits[0] }
                continue
            }
            Write-Host ""
            Write-Info "Options matching '$choice':"
            foreach ($h in $hits) { Write-Host ("  {0,3}  {1,-10} {2}" -f $h.Key, $h.Group, $h.Text) -ForegroundColor Gray }
            $pick = (Read-Line -Prompt '  Number (ENTER = back) > ' -LogLabel 'Menu search pick').Trim()
            if ($pick -and $keyMap.ContainsKey($pick)) { Invoke-MenuEntry -Entry $keyMap[$pick] }
        }
    } finally {
        Disconnect-CloudSite
        Set-CtrlCAsInput $false
        Complete-AllProgress
        Restore-AppTitle
        Write-Log "===== SESSION END: $($script:Session.Jobs) job(s), $(Format-FileSize $script:Session.Bytes) =====" 'INFO' -FileOnly
    }
}

# ======================================================================================
# GO
# ======================================================================================
if ($LibraryMode) {
    Import-AppSettings
    return
}
Start-MainMenu
