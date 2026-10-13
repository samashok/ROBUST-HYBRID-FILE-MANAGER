# Robust Hybrid File Manager 2.0

A friendly, menu-driven file manager for **SharePoint Online, OneDrive and local / network folders**,
built for **Windows PowerShell 5.1** (the version that ships with Windows 10 / 11).

Copy, move, download, upload, verify, search, wipe and archive files — including leavers' OneDrives —
from one screen, with pause / cancel, resume, full audit logging and clear summaries.

```
  ╔═══════════════════════════════════════════════════════════════════════════╗
  ║ ROBUST HYBRID FILE MANAGER 2.0                      Good afternoon, alex  ║
  ╚═══════════════════════════════════════════════════════════════════════════╝
       This session: 2 job(s), 4.31 GB moved   |   Favourites: 5   |   Cloud: not connected

  CLOUD  (same tenant, server-side)      LOCAL  (PC, USB, network share)
    1  Copy a folder                       7  Copy a folder
    2  Move a folder                       8  Move a folder
    3  Batch: pick items                   9  Batch: pick items
    4  Wipe a folder (Recycle Bin)        10  Wipe a folder (permanent!)
    5  Verify two folders match           11  Verify two folders match
    6  Search by name                     12  Search by name

  DOWNLOAD / UPLOAD                      ARCHIVE  (download, verify, recycle)
   13  Download a folder                  17  Archive one folder
   14  Download: pick items               18  Archive: pick items
   15  Upload a folder                    19  Archive many folders
   16  Upload: pick items

  LEAVER ONEDRIVES
   20  Archive a leaver's OneDrive
   21  Archive many leavers
   22  Recycle Bin links
   23  Unlicensed leaver archive
  ─────────────────────────────────────────────────────────────────────────────
  W Guided   F Favourites   J History   S Settings   H Help   L Logs   Q Quit
```

## Getting started

1. Download `Cross-Tenant-Cloud-FileManager.ps1`.
2. Right-click it and choose **Run with PowerShell**, or in a PowerShell window:
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser   # once
   .\Cross-Tenant-Cloud-FileManager.ps1
   ```
3. Type an option number, a word such as `download`, or **W** for guided mode.

**Requirements:** Windows PowerShell 5.1 and a Microsoft 365 account with access to the sites involved.
The **PnP.PowerShell 1.12.0** module is installed automatically (for your account only) the first
time you use a cloud option. Local-only options work without it.

**Sign-in** is the same as always: `Connect-PnPOnline -UseWebLogin` (a Microsoft sign-in window).
The connection is reused while you stay on the same site, so you sign in less often.
Option 23 signs in to Microsoft Graph with a device code; that sign-in is stored encrypted (DPAPI) for next time.

## Everyday tips

| Where | Key | What it does |
|---|---|---|
| Any path / address prompt | drag & drop | Drag a folder from File Explorer into the window (quotes are removed) |
| | `F` | Pick a saved favourite |
| | `R` | Re-use a path from earlier in this session |
| | `B` | Browse with the Windows folder picker |
| | `Q` | Cancel / go back |
| Lists of items | `3`, `2-5`, `1,3,7`, `1 3 7`, `A` | Pick items (ranges can be backwards) |
| | `S` / ENTER / `Q` | Search / refresh / back |
| While a job runs | `P` | Pause (any key resumes, `C` cancels) |
| | `C` or `Ctrl+C` | Cancel cleanly — the current file finishes and the summary is shown |

Every job can simply be **run again**: files that are already identical are skipped, so interrupted jobs resume.

## Safety

* **Archive options (17–23) never delete anything until every cloud file has been verified on disk**
  against a fresh SharePoint listing: correct size, no missing files, no name clashes, no folders that
  could not be listed, and not cancelled. Deleted items go to the SharePoint Recycle Bin (~93 days).
* Resumed archives compare dates as well as sizes, and an archive never overwrites a different file that is
  already in the archive folder. Only items that were verified are recycled.
* The root of a document library is never deleted — its contents are. Leaver options only accept OneDrive
  (`/personal/`) addresses.
* Every delete or archive asks once, before it starts (type `YES`).
* A local move onto itself — including the same folder reached through a mapped drive, a server alias or a
  junction — is refused; emptied source folders are removed
  only when they are truly empty; junctions / symbolic links are never followed or emptied.
* Local wipe refuses drive roots, share roots and system / profile folders.

## Files it keeps

Everything lives in `C:\Users\Public\Documents\RobustCloudWork` (change it with `-WorkDir`):

| File | Purpose |
|---|---|
| `Logs\HybridJobLog_<user>.log` | Full audit log (rotates at 5 MB) |
| `Logs\LongPaths_FailedToSync_<user>.txt` | Items skipped because their path is too long |
| `Favorites_<user>.json` | Your favourites (same format as earlier versions) |
| `Settings_<user>.json` | Display / behaviour settings (`S` in the menu) |
| `JobHistory_<user>.json` | The last 500 jobs (`J` in the menu) |
| `MgTokenCache_<user>.dat` | Encrypted Graph sign-in for option 23 |

## Command-line options

```powershell
.\Cross-Tenant-Cloud-FileManager.ps1 -WorkDir D:\FileManagerWork   # different work folder
.\Cross-Tenant-Cloud-FileManager.ps1 -Ascii                        # plain +-| borders for old fonts
.\Cross-Tenant-Cloud-FileManager.ps1 -CloudPathLimit 380           # stricter path-length limit
```

## Tests

The `tests` folder runs the real engines against temporary folders and an in-memory fake SharePoint
and Microsoft Graph. No tenant is touched:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-EngineTests.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-UiTests.ps1
```

## Version history

See the comment block at the top of the script. Versions 1–111 are in this repository's git history;
2.0 is a complete overhaul that keeps the option numbers, file formats and sign-in methods.
