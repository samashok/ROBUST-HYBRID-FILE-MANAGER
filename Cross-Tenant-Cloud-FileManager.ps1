<#
# ========================================================================================
#                             ROBUST HYBRID FILE MANAGER (v111)
#                    The Educational Masterclass & Telemetry Edition
#
# --- THE MASTER CHANGELOG ---
#   v1:  Initial script framework created for basic PnP PowerShell transfers.
#   v2:  Added Microsoft PnP module dependency check and auto-install logic.
#   v3:  Implemented basic Add and Remove loops for Cloud operations.
#   v4:  Fixed module import looping issue; stabilized connection caching.
#   v5:  Added Local Transfer engines for basic hard drive copying.
#   v6:  Replaced standard 'Copy-Item' with raw .NET [System.IO.File] for speed.
#   v7:  Built the 'Get-LongPath' kernel bypass to destroy the 260-character limit.
#   v8:  Unified the Main Menu to handle both Cloud and Local operations.
#   v9:  Added support for multi-line pasted inputs in the Batch selection menu.
#   v10: Added the 'Get-SafeWindowsName' sanitizer to fix illegal Mac characters.
#   v11: Introduced the Hybrid Engine to bridge Local drives and Cloud URLs.
#   v12: Built the 'Parse-CloudUrl' cleaner to automatically strip tracking tags.
#   v13: Enhanced the 'Write-Log' function with a rotating 5MB file archiver.
#   v14: Fixed array range operator (1..5) for strict PowerShell backward compatibility.
#   v15: Built the 'Check-Interrupt' hardware sensor for instant Pause/Cancel hotkeys.
#   v16: Hid advanced verification metrics behind a clean, simple UI.
#   v17: Added the 'Invoke-ForceDelete' engine to aggressively bypass RAM locks.
#   v18: Built the 'Remove-EmptyDirectories' sweeper to clean up after Move jobs.
#   v19: Fixed Graph API URL extraction bugs by adding deep AbsolutePath logic.
#   v20: Redesigned the Batch Menu UI for cleaner, wrapped text reading.
#   v21: Added the 'Get-LocalDirectoryMetrics' accountant for fast folder sizing.
#   v22: Added performance timers (Stopwatch) to all major operations.
#   v23: Removed test loops and prepared the core engines for production use.
#   v24: Reverted to stable WebLogin authentication to fix ISE EventSourceExceptions.
#   v25: Added the 3-try 'Write-Log' retry loop to bypass strict Antivirus file locks.
#   v26: Built the 'Get-CloudInventory' recursive scanner using memory-safe Queues.
#   v27: Fixed regex space-escaping bugs in the URL parser.
#   v28: Added the "Sticky Selection Loop" so typos in Batch don't crash the menu.
#   v29: Built the 'Perform-CloudVerify' engine for absolute data integrity checks.
#   v30: Added the 'Perform-LocalVerify' engine for hard drive integrity checks.
#   v31: CRITICAL FIX: Removed all backtick-escaped quotes to prevent parser corruption.
#   v32: Built the 'Perform-DirectDownload' engine using direct REST API calls.
#   v33: Implemented CSOM Binary Fallback for direct downloads that fail via REST.
#   v34: Added Exponential Backoff (5s, 10s, 15s) to prevent Microsoft API throttling.
#   v35: Removed -ConsoleOnly flag from operational logs to ensure permanent audit trailing.
#   v36: Added the '-SkipAuth' flag to prevent session drops during massive Batch verifies.
#   v37: Added Summary Acknowledgment logic to track perfectly synced files.
#   v38: Refined Batch menu hotkeys for more intuitive keyboard navigation.
#   v39: Built the 'Perform-CloudToLocalBatchDownload' engine.
#   v40: Fixed UI Scope Overrides so Batch progress bars show the active file names.
#   v41: Added "Textbook" comments (PowerShell 101) to every function.
#   v42: Restored missing file logs that were accidentally removed during optimization.
#   v43: I/O Optimization: Pushed high-volume "Skipped" logs to screen-only to save disk speed.
#   v44: Built the "Grand Batch Summary" to aggregate math at the end of massive jobs.
#   v45: Added Live File Size Telemetry directly into the console outputs.
#   v46: Built the Master Data Auditor to mathematically prove Source and Dest sizes match.
#   v47: Fixed the Elapsed Time reset bug by passing Inherited Start Times to worker functions.
#   v48: Added Payload tracking (Bytes Transferred) to the Local Copy/Move engines.
#   v49: Synchronized the Payload tracking outputs across all Download and Batch engines.
#   v50: Added Pro-Tip UI elements and dynamic list-refreshing via the ENTER key.
#   v51: Fixed UI Scope Override specifically for nested Direct Download progress bars.
#   v52: Added the 'Q' hotkey to back out of any menu gracefully.
#   v53: Allowed 'Q' to cleanly exit the entire script from the Main Menu.
#   v54: Multi-User Safe Logging: Injected $env:USERNAME into log files to stop write conflicts.
#   v55: Built 'Clear-KeyBuffer' with a Try/Catch to prevent crashes inside PowerShell ISE.
#   v56: Final Polish: Fully unrolled changelog and comprehensive PowerShell beginner lessons.
#   v57: Added file sizes to Batch menus (Protected Cloud API from recursive throttling).
#   v58: Enabled Deep Local Telemetry to actively display Folder sizes in Local Batch menus.
#   v59: Restored full permanent audit trailing by writing all file-level telemetry to the .log file.
#   v60: Built 'Bookend Auditing' to permanently log Paths, Menus, & Summaries without I/O lag.
#   v61: Added comprehensive, beginner-friendly inline documentation for cross-training.
#   v62: Implemented comprehensive User Input logging via custom -FileOnly switch.
#   v63: Deepened educational comments across all advanced execution engines for non-technical readability.
#   v64: Injected Target Name variables into completion logs and batch headers for precise audit trailing.
#   v65: Moved elapsed time directly into all Batch Grand Summary blocks and added Hours format to all timers.
#   v66: Restored rich per-item Payload Summary blocks to Local Transfers and all Batch Delete/Cloud actions.
#   v67: Added dynamic Local Directory Metrics summary blocks to the end of Local Batch operations.
#   v68: Reordered batch completion sequence so Grand Batch Summary is strictly presented last.
#   v69: Injected detailed error exception reasons directly into inline [ERROR] logs for real-time tracking.
#   v70: Routed all Summary/Metric/Failure reports through Write-Log to guarantee physical .log preservation and patched Parse-CloudUrl to prevent local-path crashes.
#   v71: Implemented Pre-Flight Storage Diagnostics with Smart Overwrite Calculation for Local Transfers, and Informational Space checks for Cloud Downloads.
#   v72: Enriched Cloud Batch menus with zero-latency Item Count and Last Modified metrics for folders.
#   v73: Added 'Perform-LocalSearch' engine utilizing a .NET Queue architecture to prevent Access Denied crashes.
#   v74: Trapped Search functions in an infinite loop until 'Q' is pressed. Added a dedicated Cloud Search engine equipped with an anti-throttling shield. Integrated the [S] Search action directly into all Batch menus.
#   v75: Upgraded Search engines to actively match Folders. Rebuilt Cloud Search logs to display absolute paths and injected an interactive Link Generation Menu post-search.
#   v76: Redesigned Search Menus to display Full Absolute Paths. Added CSV Export features to all Search tools. Added 'Open in Explorer' shortcut to Local Search. Added 'Open Logs' hotkey to Main Menu.
#   v77: Upgraded Search engines with continuous loops, live progress bars, full path displays, interactive launch menus, and patched Cloud Search batch handoffs along with PnP disconnect crashes.
#   v78: Upgraded Local Batch menus to display recursive Item Counts alongside folder sizes to match Cloud UI.
#   v79: Rebuilt Local Batch item count logic to strictly mirror Microsoft API's top-level architecture for exact 1:1 Cloud/Local parity checking.
#   v80: Built the Favorites & Bookmarks System. Users can now save named shortcuts for Cloud URLs and
#        Local Paths and recall them instantly with [F] at any input prompt. Added [F] Favorites Manager
#        to Main Menu for full add/rename/delete/list control. Added 'Read-PathInput' helper that wraps
#        every URL/path prompt with [F] picker and [Q] cancel, eliminating repetitive typing for
#        common operations. Favorites persist to disk as JSON in the WorkDir.
#   v81: CRITICAL BUG FIX: Rebuilt the Favorites system to resolve broken picker in all Batch/Download
#        menus. Root cause: v80 converted ConvertFrom-Json PSCustomObjects into hashtables, then piped
#        hashtable[] through Where-Object — in PowerShell 5, .Type on a hashtable in a pipeline can
#        resolve to .GetType() instead of the key, corrupting all property reads. Fix: Get-Favorites
#        now returns [PSCustomObject] array directly (no hashtable conversion). Save-Favorites uses
#        ConvertTo-Json -InputObject @(...) to guarantee correct JSON array serialization even for
#        single-item lists. Show-FavoritePicker uses a for-loop with index access instead of pipeline
#        filtering to eliminate the issue entirely. Manage-Favorites builds updated lists using
#        consistent PSCustomObjects to prevent type drift across save/load cycles.
#        NEW FEATURE: Quick-Save prompt in Read-PathInput — after typing any URL or path manually
#        that isn't already saved, the script offers to save it as a Favorite on the spot.
#   v82: Complete UX overhaul of all user-facing text. Rewrote the main menu with clearer option
#        descriptions, a colour-coded QUICK GUIDE for first-time users, and a universal keyboard
#        shortcuts reference. Rewrote every EXPLAINED briefing screen with plain-English
#        step-by-step instructions, real-world examples, what-to-expect notes, and tips covering
#        how to get a SharePoint URL, what a local path looks like, and when to use each mode.
#        Improved all input prompts to be more instructional (e.g. 'how to copy the URL').
#        Improved all error messages to suggest a corrective action rather than just stating the
#        problem. Added a HOTKEY CHEAT SHEET section to the main menu. All logic is identical to
#        v81 — only text and presentation were changed.
#   v83: CRITICAL BUG FIX: Rewrote Save-Favorites serialization to fix a PowerShell 5 quirk
#        where ConvertTo-Json -InputObject @($array) on an object[] containing multiple
#        PSCustomObjects produced a single merged JSON object instead of a JSON array — causing
#        all favourites to be concatenated into one corrupted entry on disk. Fix: each favourite
#        is now serialized individually with ConvertTo-Json -Compress -Depth 1 and the results
#        are joined into a valid JSON array string manually, completely bypassing the PS5 quirk.
#        REMOVED Quick-Save prompt from Read-PathInput entirely — the on-the-fly save offer was
#        disruptive during normal operation. Favourites should be configured from the Main Menu
#        [F] option only, which is the intended workflow.
#        UPDATED Quick Guide: changed confusing 'Moving files between two different companies'
#        text to 'Copy or move files locally between folders on your PC'.
#   v84: Full UNC Path Support for all local and download operations (\\server\share\folder).
#        ROOT CAUSE: The script previously used [System.IO.DriveInfo] for free-space checks,
#        which only understands local drive letters (C:\) and throws on UNC roots, causing the
#        pre-flight storage check to prompt 'proceed Y/N?' for every UNC operation.
#        FIX: Extended the Win32 kernel class with GetDiskFreeSpaceExW from kernel32.dll — the
#        same API Windows Explorer uses internally. This works transparently on BOTH local paths
#        and UNC paths (\\server\share\...) with zero special-casing required. Added a new
#        Get-PathFreeSpace helper that calls GetDiskFreeSpaceExW first and falls back to
#        DriveInfo for maximum compatibility. Rewrote Show-DestinationFreeSpace and
#        Invoke-PreflightStorageCheck to use the new helper — UNC paths now get full free-space
#        telemetry and smart overwrite calculation just like local drives.
#        Updated all relevant menu EXPLAINED screens and the Quick Guide with UNC path examples
#        (\\FileServer\HR\Leavers) so users know UNC is supported.
#   v85: VISUAL & UX OVERHAUL.
#        Visual: Replaced all plain = and - separators in the main menu and all EXPLAINED screens
#        with Unicode box-drawing characters (╔═╗║╚╝├┤─) for a cleaner, more professional look.
#        Session History: Added $script:PathHistory — a lightweight in-memory dictionary that
#        remembers the last-used value per prompt type within a session. A new [R] hotkey in
#        Read-PathInput lets users instantly re-use any path from the current session without
#        retyping or opening Favourites. History is lost when the script closes (intentional —
#        no sensitive path persistence between sessions).
#        Auto-Buffer Default: At the Temp Buffer prompt for downloads, pressing ENTER without
#        typing anything now automatically uses $env:TEMP (the system temp folder). This
#        eliminates the most repetitive prompt for users running repeated download jobs.
#        Smarter Prompts: Read-PathInput now shows both [F] and [R] hints simultaneously
#        when both Favourites and session history are available for that prompt type.
#        Improved batch index menus: added a compact legend explaining Items/Modified/Size
#        column headers so new users understand what they are reading.
#   v86: Added Write-SummaryBox helper — replaces plain ========= borders in all Grand Batch
#        Summary, Verification, and Download Complete output blocks with clean Unicode-bordered
#        result panels for a much more readable post-operation summary.
#        Upgraded Manage-Favorites and Show-FavoritePicker to full box-drawing visual style
#        matching the rest of the v85 design. Added [E] Edit Value action to Manage-Favorites
#        so users can update a saved URL or path without deleting and re-adding it.
#        Added [H] Quick Help from the main menu — a concise one-screen reference card covering
#        all hotkeys, selection syntax, UNC paths, and wildcard search patterns.
#        Added column-header legend above all three batch index menus (Cloud Batch, Local Batch,
#        Batch Download) explaining Items / Modified / Size columns for new users.
#   v87: Completed the Write-SummaryBox integration from v86. All three Grand Batch Summary
#        blocks (Cloud Batch, Local Batch, Batch Download) now render as clean Unicode-bordered
#        result panels instead of plain ========= log lines. Each panel shows the operation
#        outcome colour-coded: green for success, red for errors, yellow for warnings.
#        Upgraded Check-Interrupt pause banner from plain ===== lines to a full Unicode box
#        showing the exact folder and file where the job was paused — making it easier to
#        manually resume from the right place if needed.
#        Upgraded all three batch index menu footers from the plain ----- PRO TIP block to a
#        cleaner box-style footer with a concise one-line selection cheat-sheet.
#   v88: Completed the full output visual overhaul.
#   v89: Added full SharePoint UPLOAD capability for Windows PowerShell 5.1.
#        New: Ensure-CloudFolder — wraps Resolve-PnPFolder with a manual segment-by-
#        segment Add-PnPFolder fallback for older PS environments. Uses a HashSet cache
#        so repeated calls for the same folder avoid redundant API round-trips.
#        New: Invoke-UploadWithRetry — single-file upload via Add-PnPFile with the
#        same 5-attempt exponential backoff (5s,10s,15s,20s,25s) as SafeDownload.
#        New: Perform-DirectUpload — mirrors Perform-DirectDownload. Handles single-
#        file and full-directory recursive uploads with byte-size skip logic, cloud
#        path length guards (logs over-limit files to LongPaths_FailedToSync.txt),
#        per-file progress, Write-SummaryBox results, and batch-aware output.
#        New: Perform-LocalToCloudBatchUpload — mirrors CloudToLocalBatchDownload.
#        Shows local directory as a numbered index; supports range selection, [A]=all,
#        [S]=search, pause/cancel hotkeys, and a Grand Batch Summary box.
#        Menu: Added Section D UPLOAD OPERATIONS with [15] DIRECT UPLOAD and [16]
#        BATCH UPLOAD. Quick Guide updated with upload hint and cross-site bridge
#        workflow tip. Help card updated. Exit moved from [15] to [Q]-only.
#        Auth: Both upload options use Connect-PnPOnline -UseWebLogin — the same
#        tested, working method as all other cloud operations in this script.
#   v90: Added ARCHIVE OPERATIONS — Download + Verify + Delete — for leaver OneDrive
#        cleanup workflows where SharePoint storage needs to be reclaimed.
#        New: Invoke-ArchiveVerify — cross-checks download result stats (Success +
#        Skipped counts and byte totals) against a fresh Get-CloudInventory scan of
#        the cloud source. Compares file counts and total bytes, NOT filenames, so
#        Get-SafeWindowsName character substitutions never cause false failures.
#        Falls back to Get-PnPFile for single-file source URLs.
#        New: Perform-ArchiveDownload — three-step guided workflow for [17]:
#          Step 1: Download (Perform-DirectDownload, full output).
#          Step 2: Verify (Invoke-ArchiveVerify with Write-SummaryBox result).
#          Step 3: Offer deletion — requires explicit 'YES'; uses Remove-PnPFolder
#          with -Recycle (sends to SharePoint Recycle Bin, ~93 days recoverable).
#          Any download failure or verification mismatch halts before Step 3.
#        #        New: Perform-BatchArchiveDownload — batch version for [18]. Shows cloud
#        directory index, user picks items, each is processed one-by-one:
#        download → verify → per-item Y/N/A(bort) delete prompt. Grand summary
#        box shows Archived / Deleted / Skipped / Failed counts.
#        Menu: Added Section E ARCHIVE OPERATIONS with [17] ARCHIVE DOWNLOAD and
#        [18] BATCH ARCHIVE. Quick Guide updated. Help card updated.
#   v93: Removed all "leaver/leavers" terminology — script is general-purpose and
#        appropriate for any user or organisation. Updated all menu texts, explained
#        screens, examples, and function comments to use neutral language.
#        Removed all company-specific references and example paths.
#        Fixed confusing "NEW HERE?" banner — replaced with rotating helpful tips.
#        Read-PathInput: full path now always shown without truncation before [R].
#        [F] Favourites hint now always shown at every prompt regardless of whether
#        any favourites are saved — so users discover the feature immediately.
#        Added Write-PathBanner helper — shows Source/Destination/Buffer paths at
#        the top of every batch index menu refresh so users always know which
#        operation is in progress (Cloud Batch, Local Batch, Batch Download, Batch
#        Upload, Batch Archive).
#        Added "Press ENTER to return to the menu" pause after: [L] Open Logs,
#        [17] Archive Download completion, [19] Multi-Source Archive completion.
#        This gives users time to read the final summary before the menu redraws.
#        Visual: improved EXPLAINED screens for [15]/[16] upload options to remove
#        cross-site bridge company policy references. All text is now neutral.
#   v91: Added [19] MULTI-SOURCE ARCHIVE. Paste URLs from multiple leavers'
#        OneDrives (or any SharePoint sources) into a single queue. For each URL
#        the script extracts the cloud folder name as the local subfolder name,
#        then offers a friendly-name override so archives are clearly labelled
#        (e.g. "John Smith" instead of the raw SharePoint path segment). A plan
#        preview is shown before any work starts. Each source is then processed
#        unattended: Download → Verify → Auto-delete on pass. Re-authentication
#        is handled automatically when sources span different SharePoint sites.
#        Grand summary panel at the end lists every source and its outcome.
#   v92: Fixed version banner (was showing v90 instead of v91 due to missed substitution).
#        Added Invoke-PnPWithRetry — a universal throttle-resilient wrapper for all PnP
#        API write operations. Detects HTTP 429/503 responses, waits the Retry-After
#        header value when present, and falls back to exponential backoff (5s→10s→15s→
#        20s→25s, 5 attempts). Applied to: Move-PnPFile, Copy-PnPFile, Remove-PnPFolder,
#        Remove-PnPFile, and Add-PnPFolder. This ensures no cloud operation is silently
#        dropped under heavy Microsoft throttling conditions.
#        Added QUICK-START GUIDE to the top of the script — a plain-English orientation
#        for first-time users explaining what the script does, requirements, and the five
#        most common workflows with step-by-step instructions.
#        Improved all menu EXPLAINED screens with clearer beginner-friendly descriptions,
#        real example paths, and explicit "What you need before you start" checklists.
#        Added colour-coded section headers and a FIRST TIME? TIP banner to the main menu.
#        Improved documentation throughout all engine functions with plain-English comments
#        explaining the purpose, inputs, outputs, and key design decisions of each function.
#        Wired Write-SummaryBox into all remaining plain-text summary blocks: Cloud Verify,
#        Local Verify, Direct Download completion, and Local Transfer completion now all
#        render as clean bordered result panels colour-coded by outcome.
#        Added Write-FailedReport helper — a red-bordered panel that shows failed files with
#        their reasons clearly separated, replacing the plain ======= FAILED ITEMS blocks in
#        all 5 locations (Cloud Batch, Local Batch, Local Transfer, Direct Download, Batch DL).
#        Upgraded all three per-item batch progress headers from the plain ">>> CLOUD BATCH
#        ITEM X of Y: Name" text to compact directional dividers using Unicode arrows (─ ↑ ↓)
#        that are visually distinct and easier to scan in a long log output.
#        When Local Transfer and Direct Download are called from within a batch (not standalone),
#        they now emit a single concise log line instead of a full box — preventing the output
#        from being buried under repeated panels for every sub-item in a large batch job.
#   v95: Improved search results display for both Local Search and Cloud Search.
#        Results list now shows size and modified date inline for every item — no need to
#        click into each result just to see basic metadata. Format: two-line entries with
#        the filename highlighted in white and the full path shown below in muted grey,
#        columns aligned: [type] [size padded to 9 chars] [date yyyy-MM-dd HH:mm] [filename].
#        Folders show — for size in the list; files show — if size is not available.
#        Local detail panel: folders now show calculated size (files + sub-folders count)
#        on demand using Get-LocalDirectoryMetrics — previously the size line was suppressed
#        entirely for folders. Files with a zero-byte size now show — instead of nothing.
#        Cloud detail panel: added Size display for files, and a — note for folders
#        (cloud folder sizes are not available via the PnP folder item API without a
#        recursive scan). Added Modified date display for both files and folders.
#        Cloud Search engine: now captures TimeLastModified for both file and folder
#        FoundItems entries so the date can be displayed in the list and detail view.
#        CSV export for Cloud Search: now includes a Modified column.
#   v96: Changed search results metadata format from fixed-width column layout to the
#        friendlier inline tag style: filename [Size: 2.97 MB, Date Modified: 2025-09-13].
#        Applies to both Local Search and Cloud Search result lists.
#   v97: Reverted search results to single-line format: [N] [TYPE] full/path/to/item  [Size: x, Date Modified: yyyy-MM-dd].
#        Full path is shown directly so files and folders are immediately locatable without
#        having to read a separate second line. Applies to both Local Search and Cloud Search.
#   v98: Added [20] LEAVER ONEDRIVE ARCHIVE and [21] MULTI-SOURCE LEAVER ARCHIVE.
#        Root cause of zero-file downloads from leaver URLs solved: the /_layouts/15/
#        onedrive.aspx URL from the Exchange Admin portal is a browser UI page, not a data
#        path. The fix is to connect PnP directly to the personal site root, enumerate all
#        Document Libraries, then run a full Batch Archive on each.
#        Standard connection is attempted first (no extra permissions needed). If access is
#        blocked (account disabled/unlicensed), the user is offered escalation via Set-SPOUser
#        which grants temporary Site Collection Admin rights, then retries.
#        New: Ensure-SPOModule  — installs Microsoft.Online.SharePoint.PowerShell if absent.
#        New: ConvertTo-OneDriveSiteUrl — builds a personal site URL from a leaver UPN.
#        New: Grant-LeaverSiteAdmin  — Set-SPOUser grant with 5s propagation wait.
#        New: Revoke-LeaverSiteAdmin — removes temp grant; always called in finally block.
#        New: Perform-LeaverOneDriveArchive — archive engine: enumerates user libraries
#             (system libs like Form Templates/Site Pages/Style Library filtered out),
#             runs Perform-BatchArchiveDownload on each. Documents library content is placed
#             directly in the destination (no extra Documents\ subfolder). Fixed: Forms
#             virtual folder fast-fails on the known un-deleteable error instead of retrying
#             4 times; logged as INFO not ERROR.
#        New: Perform-MultiSourceLeaverArchive — unattended multi-leaver engine.
#             Queues N leavers (UPN or URL with friendly label), shows plan, then runs all:
#             standard connect → escalate if needed → archive each library → outcome table.
#             SPO admin credentials optionally pre-collected once for all auto-escalations.
#        New: Start-InterruptibleSleep — hotkey-aware sleep replacing bare Start-Sleep in
#             Invoke-PnPWithRetry. Checks [P]/[C] every second during retry wait periods
#             so the user can pause or cancel while waiting between retries.
#        Menu: Section F now has [20] single leaver and [21] multi-leaver.
#             Updated Quick Start Guide, Help card, not-recognised regex (1-21).
#   v99: Removed SPO Admin escalation path from [20] and [21].
#        The standard PnP UseWebLogin connection (which worked successfully in testing)
#        is the only method used. If a connection fails the user is informed clearly
#        and returned to the main menu — no Set-SPOUser, no admin URL prompts, no
#        Ensure-SPOModule calls from the menu handlers.
#        Perform-MultiSourceLeaverArchive simplified: SpoAdminUrl / AdminUpn / RevokeAfter
#        parameters removed; escalation block removed; standard connect only.
#        Option [20] intro text updated to remove escalation references.
#        Option [21] removes the escalation credentials section entirely.
#        Section F menu description updated. Help card updated.
#   v100: Added [22] LEAVER RECYCLE BIN LINKS — a dedicated option that takes one or
#        more leaver OneDrive URLs and outputs the corresponding SharePoint Recycle Bin
#        URLs ready to open in a browser for manual review and emptying.
#        New: Get-LeaverRecycleBinUrl — converts any leaver OneDrive URL or UPN-derived
#             site URL to the /_layouts/15/RecycleBin.aspx URL for that personal site.
#        Updated: Perform-LeaverOneDriveArchive — now returns LeaverSiteUrl in the result
#             hashtable so the [20] menu handler can display the recycle bin link in the
#             post-archive summary.
#        Updated: Perform-MultiSourceLeaverArchive — stores each leaver's SiteUrl and
#             outcome so the [21] menu handler can print a recycle bin link table after
#             the run completes, alongside the per-leaver outcome table.
#        Updated: [20] menu handler — shows recycle bin link after archive completes.
#        Updated: [21] menu handler — shows recycle bin link table after run completes.
#        Menu: Added [22] to Section F. Updated help card. Updated not-recognised regex.
#   v101: Improved [22] LEAVER RECYCLE BIN LINKS.
#        Removed the friendly label prompt — URLs are processed immediately with no
#        extra questions. Each recycle bin URL is confirmed inline as it is added.
#        After display, three post-actions are offered in a single prompt:
#          [C] Copy all recycle bin URLs to clipboard (one per line, URLs only).
#          [O] Open all links in the default browser (one tab per leaver, 1s apart).
#          [B] Both — copy AND open.
#          ENTER — skip / return to menu.
#        Updated Write-RecycleBinLinks (used in [20]/[21] summaries) to also offer
#        the same C/O/B prompt so admins can act immediately after an archive run.
#   v102: Fixed [21] MULTI-SOURCE LEAVER ARCHIVE showing an interactive batch selection
#        menu — the library index with [A]/numbers/[Q] — during what should be a fully
#        unattended run.
#        ROOT CAUSE: Perform-LeaverOneDriveArchive calls Perform-BatchArchiveDownload,
#        which always shows the interactive selection menu and waits for user input.
#        FIX: Added [switch]$AutoSelectAll to Perform-BatchArchiveDownload. When set,
#        the function skips the menu display and Read-Host entirely — it enumerates
#        the directory, selects ALL items immediately, and proceeds straight to the
#        Download → Verify → Delete loop. The menu is only shown when AutoSelectAll
#        is NOT set (i.e. for the interactive [18] and [20] entry-level flows).
#        Perform-LeaverOneDriveArchive now passes -AutoSelectAll to
#        Perform-BatchArchiveDownload, making [20] and [21] both fully unattended
#        after the user confirms YES.
#        The outer while($true) loop in Perform-BatchArchiveDownload is also bypassed
#        in AutoSelectAll mode — once the items are processed it returns immediately
#        without re-scanning and re-showing the menu.
#        Also: [21] end-of-run summary now calls Write-RecycleBinLinks with the
#        C/O/B prompt so admins can open/copy all recycle bin links immediately.
#   v103: Four targeted bug fixes from real test run logs:
#        FIX 1: Added $FastFailPatterns to Invoke-PnPWithRetry. "Forms directory" and
#               "checked out for editing" errors now short-circuit immediately without
#               burning retry attempts (saves 75s per occurrence).
#        FIX 2: Deletion catch block split into 3 distinct branches — Forms system folder
#               (INFO, silent), checked-out file (WARN + clear guidance + $GrandCheckedOut
#               tracking), and other errors (ERROR). Checked-out items listed in a
#               "DELETION BLOCKED" summary panel with step-by-step remediation.
#        FIX 3: $GrandCheckedOut tracker added. Summary shows "Blocked by Checkout: N"
#               row plus per-item detail and instructions to check in and re-run.
#        FIX 4: Join-String (PS6.2+ only) replaced with -join operator in
#               Write-RecycleBinLinks for full PowerShell 5.1 compatibility.
#   v104: Added [23] UNLICENSED LEAVER ONEDRIVE ARCHIVE.
#        NEW 1: Engine Suite G — for leavers whose E5 license was removed and whose
#               OneDrive is now frozen/inaccessible. Assigns a temporary SharePoint Plan 1
#               license via an Azure AD security group, waits for the OneDrive to become
#               accessible, runs the full archive (download+verify+delete), then removes
#               the user from the group immediately to free the license slot.
#        NEW 2: Ensure-MgGraphForLicensing — lazy-loads Microsoft.Graph.Authentication,
#               Microsoft.Graph.Groups, Microsoft.Graph.Users (only when [23] is first
#               used). Supports silent token reconnect + device code fallback.
#        NEW 3: Perform-LicensedLeaverArchive — per-leaver sequential engine.
#               Steps: resolve user (Graph) → add to group → poll OneDrive every 30s
#               (up to 30 min) → archive via Perform-LeaverOneDriveArchive →
#               remove from group. One leaver at a time = one license in use at a time.
#        NEW 4: Clipboard V-paste at leaver list prompt. Accepts the full 6-column
#               OneDrive usage report format (email, URL, Quota, Used, %, Folder name)
#               or a 3-column compact form (email, URL, Folder name). Columns 3-5 of
#               the 6-column format (quota/used/percent) are ignored automatically.
#        NEW 5: $global:MgGraphConnected state variable added for session persistence.
#   v105: Consistency pass — [23] aligned with [20]/[21] standards:
#        FIX 1: [23] local-dest and buffer inputs now use Read-PathInput (Favourites [F],
#               repeat-last [R], session history) instead of bare Read-Host.
#        FIX 2: [23] buffer defaults to $env:TEMP on blank ENTER (matching [20]/[21]).
#        FIX 3: [23] buffer path tested for write access before starting (matching [21]).
#        FIX 4: [23] path inputs cleaned (quotes stripped, trailing backslash removed)
#               via standard .Trim()/TrimEnd() calls, matching all other options.
#        FIX 5: Perform-LicensedLeaverArchive now collects SiteUrls of successfully
#               archived leavers and displays Write-RecycleBinLinks at the end
#               (matching [20] and [21] behaviour — recycle bin is shown after delete).
#        FIX 6: Help card [H] updated to list [23] in the Leaver Operations section,
#               and footer note updated to "show Recycle Bin links at end of run."
#        FIX 7: Two Write-Log calls in [23] used "INPUT" level; changed to "INFO"
#               to match the rest of the script (user inputs use "INFO" + "[USER INPUT]"
#               prefix in the message, not a custom level).
#   v106: Fix Microsoft.Graph module incompatibility with PowerShell 5.1.
#        ROOT CAUSE: Microsoft.Graph v2.x requires .NET 6+. PowerShell 5.1 runs on
#        .NET Framework 4.x. Importing Microsoft.Graph.Authentication v2.36.1 on PS 5.1
#        throws "Method 'GetTokenAsync' does not have an implementation."
#        FIX: Removed all Microsoft.Graph module dependencies from option [23].
#        All Graph operations now use native Invoke-RestMethod (built into PS 5.1).
#        NEW 1: Save-MgTokenCache / Load-MgTokenCache — persists the OAuth refresh
#               token to disk using Windows DPAPI encryption (per-user, per-machine)
#               for silent reconnect on subsequent sessions.
#        NEW 2: Update-MgGraphToken — silently refreshes the access token (1h TTL)
#               using the stored refresh token. Called proactively 5 min before expiry
#               so long-running archive jobs never hit an expired-token error mid-run.
#        NEW 3: Invoke-MgRestMethod — authenticated Graph REST helper that auto-refreshes
#               and returns parsed JSON. Replaces Get-MgUser, New-MgGroupMember,
#               Remove-MgGroupMemberByRef, Get-MgGroup throughout option [23].
#        NEW 4: Ensure-MgGraphForLicensing rewritten — device code flow via native REST
#               (POST to /oauth2/v2.0/devicecode, poll /oauth2/v2.0/token). No module
#               install, no Import-Module, no version conflict risk.
#   v107: Improved error handling in option [23] — stay in option on bad input.
#        FIX 1: Wrong group name no longer exits to main menu. Step 3 now wraps group
#               name prompt + resolution in a retry loop. On "not found" or API error,
#               the user is told what went wrong and re-prompted for the name. Q exits.
#        FIX 2: Graph connection failure shows a clear message and returns to menu
#               (previously the same `continue` as success-with-wrong-name, confusing).
#        FIX 3: Pressing ENTER on a blank Row 1 no longer silently exits to the menu.
#               It now warns "Nothing entered yet" and keeps the row prompt open.
#               Blank ENTER only closes row entry once at least one row has been typed.
#        FIX 4: All parse failures (bad format) no longer exit to menu. The outer
#               leaver-list loop re-enters row collection showing a format reminder.
#               The user can correct and re-paste without re-entering earlier steps.
#   v108: Large file progress + session transfer total.
#        NEW 1: Copy-FileWithProgress — stream-based local copy in 4 MB chunks.
#               For files ≥ 1 GB, replaces [System.IO.File]::Copy() with a chunk loop
#               that calls Write-Progress -Id 2 (child bar, parented to Id 1) every
#               400 ms. Shows: file name, bytes done / total, and MB/s transfer speed.
#               Files < 1 GB continue to use the fast kernel Copy() with no overhead.
#               Move operations are unchanged (same-drive moves are instant renames).
#        NEW 2: $global:SessionTransferredBytes — running byte total across all
#               operations in the current PowerShell session. Updated after every
#               successful local copy/move file AND every successful cloud download.
#               Shown in the Write-Progress CurrentOperation line and in every
#               summary box as "Session Total: X.XX GB transferred this session."
#        NEW 3: Cloud download large-file indicator — for files ≥ 1 GB, the
#               "Downloading:" line prints the file size in yellow ("large file,
#               please wait") so operators aren't left wondering if the job is stuck.
#               No change to the PnP download mechanism — zero additional API calls.
#        NEW 4: $script:LargeFileBytes = 1 GB threshold constant — adjustable in
#               the CONFIGURATION SECTION without touching individual functions.
#   v109: UI consistency sweep — full parity between download and upload engines, plus
#         ETA and average-speed displays across every batch engine. All changes are
#         pure-local computation: no extra API calls, no throttling risk, no slowdowns.
#        NEW 1: Format-Duration, Get-ETA, Format-AvgSpeed helpers — small pure functions
#               that convert raw numbers into human-readable display strings. Used
#               throughout the script so every progress bar and summary speaks the
#               same language ("5m 12s", "12.4 MB/s") regardless of which engine
#               produced it. Returns "" on invalid inputs so callers can safely
#               concatenate without conditionals.
#        UX 1:  ETA (estimated time remaining) added to every long-running engine:
#               Local Sync, Local Batch, Cloud Batch, Direct Upload (directory mode),
#               Batch Upload, Batch Download, and Copy-FileWithProgress. ETA is based
#               on file-count progress for batch parents and on byte progress for the
#               in-progress large file. ETA is suppressed during the first second to
#               avoid wildly inaccurate predictions.
#        UX 2:  "Average Speed" row added to every summary box that tracks bytes
#               (Local Sync, Local Batch, Direct Download, Batch Download, Direct
#               Upload, Batch Upload). Calculated as success-bytes / elapsed-seconds.
#               Suppressed if zero bytes were transferred or the operation ran for
#               less than one second.
#        FIX 1: Upload engine had NO session tracking in v108 — fixed. Perform-DirectUpload
#               now increments $global:SessionTransferredBytes after every successful
#               file upload (single-file and directory modes), and surfaces the
#               running total in both the Write-Progress CurrentOperation line and
#               the final summary box. This matches the download engine exactly.
#        FIX 2: Upload engine had NO "large file" indicator — fixed. Files ≥ 1 GB now
#               print the same yellow "(SIZE) — large file, please wait..." message
#               operators are used to seeing on downloads, in both single-file and
#               per-file paths of directory uploads.
#        FIX 3: Batch engines (Cloud Batch, Local Batch, Batch Download, Batch Upload)
#               displayed elapsed time but not ETA or session total in v108 — fixed.
#               Every batch Write-Progress now shows: Elapsed | ETA | Session | P=Pause C=Cancel.
#        FIX 4: Latent bug in v108 Direct Download summary — the inline conditional
#               `$(if($x) { @{...} })` injected a $null array element when false,
#               which Write-SummaryBox rendered as an empty padded row. Fixed two ways:
#               (a) Write-SummaryBox now skips $null rows defensively, and
#               (b) the Download / Upload / Batch Download / Batch Upload summaries
#               were converted from inline-array literals to typed
#               List<hashtable> so conditional rows can be Add()ed cleanly.
#        REL 1: Copy-FileWithProgress now uses three named stopwatches with distinct
#               purposes ($totalSw for ETA, $uiSw for 400 ms tick, $speedSw for the
#               2-second rolling speed window). Behavior is identical for the user;
#               the code is just clearer about what each stopwatch is for. ETA is
#               based on overall avg rate rather than the 2-second window to avoid
#               jitter in the displayed prediction.
#   v110: Storage classification, throughput metrics, scan heartbeat, and a clean
#         Write-Progress -Id hierarchy so batch + per-item + nested-copy bars render
#         as a visually nested stack instead of three parallel bars competing for
#         the same row. All changes are pure local computation: no API calls, no
#         throttling risk, no slowdowns. Verified PS 5.1 clean and AST-parse clean.
#        NEW 1: Get-PathStorageType — classifies a local path as LOCAL, MAP (mapped
#               drive), UNC, or OneDrive. Detection order matters: OneDrive sits
#               inside a LOCAL drive letter, so we test the OneDrive substring
#               before falling through to the LOCAL drive-letter test. Long-path
#               prefixes (\\?\C:\..., \\?\UNC\...) are normalised before testing.
#        NEW 2: Format-PathWithStorageTag — wraps a path with its classifier tag
#               for log lines, e.g. "SOURCE: [LOCAL] C:\Users\sam\file.txt".
#               Adopted by every operation-start log line that touches a local
#               path (LocalSync, LocalBatch, DirectUpload, BatchUpload,
#               DirectDownload local-dest, BatchDownload local-dest).
#        NEW 3: Format-FilesPerMin — formats files-per-minute throughput
#               ("1,250 files/min"). Surfaces in every transfer summary box when
#               the operation processed at least 50 files. Most useful for batches
#               of many small items where bytes/sec is misleading.
#        NEW 4: Get-ItemsWithHeartbeat — directory-tree enumerator that emits a
#               progress heartbeat every 2 seconds during the scan. Uses .NET
#               EnumerateFileSystemInfos under the hood to get a streaming yield,
#               with a Get-ChildItem fallback if EnumerateFileSystemInfos throws
#               (e.g. on access-denied paths). Adopted by Perform-LocalTransfer
#               for its initial source-tree scan. Self-suppresses for fast scans
#               (< 2 s) so the console stays clean.
#        UX 1:  "Storage" row added to every transfer summary box, showing the
#               classifier of the local path involved. Single-direction operations
#               show e.g. "LOCAL -> Cloud" (upload) or "Cloud -> OneDrive" (download).
#               Two-path local operations show e.g. "LOCAL -> MAP". When both ends
#               are cloud (Cloud Batch) or both can't be classified, the row is
#               suppressed.
#        UX 2:  "Throughput" row (files/min) added to every transfer summary box
#               below "Average Speed". Suppressed when fewer than 50 files were
#               processed (the rate isn't statistically meaningful for tiny jobs).
#        UX 3:  Heartbeat indicator during large enumerations — operators no
#               longer see a silent "Scanning..." for 30+ seconds while the script
#               walks a quarter-million-file tree. Now shows "[..] Scanning source :
#               175,000 items found (12s)..." updating every 2 seconds, with a
#               final "[OK] Scan complete: ..." summary.
#        UX 4:  Heartbeat during DirectUpload directory enumeration — same pattern,
#               inlined into the existing EnumerateFiles loop since it's already
#               a streaming enumerator. Self-suppresses if the scan finishes in
#               under 2 seconds.
#        UX 5:  Write-Progress -Id hierarchy fixed for visually nested progress
#               bars in Windows Terminal:
#                 Id 0 = batch engine (top-level bar — Cloud Batch, Local Batch,
#                        Batch Download, Batch Upload)
#                 Id 1 = per-item engine (LocalTransfer, DirectDownload,
#                        DirectUpload directory mode), with -ParentId 0 when
#                        called from batch mode so the renderer indents the bar
#                        under the batch bar.
#                 Id 2 = nested operation (Copy-FileWithProgress large-file copy),
#                        -ParentId 1 so it nests under the per-item bar.
#               In v108/v109 the batch engines used default Id 0 *and* the per-item
#               engines used Id 1 with batch-level percent — so the user saw two
#               bars showing the same progress. v110 makes the inner bar show
#               file-level progress (file X of Y within the current batch item)
#               so each bar conveys distinct information.
#   v111: Major UX corrections + failure diagnostics. Two long-standing bugs fixed,
#         every menu branch now keeps the operator inside the chosen operation until
#         they explicitly leave, and every transfer emits a post-op diagnostics
#         panel with failure rate, distinct error messages, and file-type breakdown.
#         Storage-aware tip helper kept on the shelf pending operator approval.
#        FIX 1: HEADER BANNER STUCK AT v102 — the opening doc-comment block in
#               every release from v103 through v110 still said "(v102)". Anyone
#               reading the script header concluded they were on an old version.
#               Now updated correctly. The script will continue to keep this
#               in sync with the main menu banner from v111 onwards.
#        FIX 2: OPERATIONS DID NOT STAY ON THE SAME MENU — In v110 and earlier,
#               selecting a menu option, running it (success OR failure), and then
#               watching the script jump straight back to the main menu was the
#               default behaviour. Operators reported losing their place,
#               missing the final summary as it scrolled away during the menu
#               redraw, and having to re-navigate the menu just to run a second
#               source in the same operation. Every one of the 23 menu branches
#               now wraps its prompt-and-engine logic in a stay-on-operation
#               loop:
#                 1) Op runs -> summary + diagnostics print
#                 2) Operator is prompted for another source (or [Q] to exit)
#                 3) On [Q], the new Show-PressAnyKeyToContinue helper pauses
#                    long enough for the operator to read the output before
#                    the main menu redraws.
#               The press-any-key pause is skipped when [Q] is hit at the very
#               first prompt (i.e. no op ran), so navigating in and out of a
#               menu without doing anything is still instant.
#        NEW 1: Show-PressAnyKeyToContinue helper — centralises the "Press any
#               key to return to the main menu..." prompt that was previously
#               duplicated (and inconsistently worded) across four leaver-op
#               branches and the multi-source archive branch. All five now use
#               the same helper. Falls back to Read-Host under non-interactive
#               hosts (Task Scheduler hidden window, some IDEs).
#        NEW 2: Get-FailureRateLevel + Format-FailureRate helpers — classify a
#               failure ratio into "" / "WARN" (>=10%) / "ERROR" (>=25%) and
#               render it as "15% (15 of 100 failed)" for display.
#        NEW 3: Get-DistinctErrors helper — groups identical error messages
#               from a FailedList and returns them sorted by frequency
#               descending. Default: ALL distinct messages (no top-N cap),
#               because in practice even a 5000-failure job tends to have
#               only a handful of distinct causes. Pass -Top N to cap.
#               The existing Write-FailedReport (the per-file dump that lists
#               every failure with its full path) is UNCHANGED and still
#               runs in every engine that previously called it — Get-DistinctErrors
#               is layered on top, not a replacement.
#        NEW 4: Get-TopExtensions helper — given any list of file items,
#               returns the top-N extensions with counts, e.g.
#               "234 .docx, 89 .pdf, 45 .xlsx, 12 .png, +3 more"
#               Items without a value in their Name/File property fall under
#               "(no ext)". Used for the per-extension breakdown row in
#               the diagnostics panel.
#        NEW 5: Show-OperationDiagnostics — the integration layer. Called from
#               every transfer summary box (LocalSync, LocalBatch, CloudBatch,
#               DirectDownload, BatchDownload, DirectUpload, BatchUpload) right
#               after the box prints, it emits a small bordered panel below
#               the summary box containing (in order, each line optional):
#                 1) Failure-rate alert  — yellow if >=10%, red if >=25%
#                 2) Distinct error messages sorted by frequency (ALL by default)
#                 3) Per-extension breakdown  — only if SuccessCount >= 50
#               Panel header is "OPERATION DIAGNOSTICS" when failures exist
#               (red/yellow), "OPERATION DETAILS" when clean (dark gray) so the
#               header colour matches the gravity of what's reported.
#        NEW 6: Show-StorageTip helper DEFINED BUT NOT CALLED — would emit a
#               one-time "[TIP] source is on a network drive (UNC). Expect
#               30-50% slower throughput..." message at the start of any op
#               touching a UNC or mapped drive. Kept dormant pending operator
#               decision on whether the tip is helpful or condescending. To
#               activate: call Show-StorageTip -Path $localPath at the start
#               of the engines.
# ========================================================================================
#>

# PowerShell 5.1 or later is required (ships as default on Windows 10/11).
# This directive makes PS throw a clear error immediately if run on PS 2/3/4
# rather than failing silently mid-operation.
#Requires -Version 5.1

# ========================================================================================
# QUICK-START GUIDE FOR FIRST-TIME USERS
# ========================================================================================
#
# WHAT IS THIS SCRIPT?
#   A file manager for SharePoint / OneDrive that works directly from the command line.
#   It can copy, move, download, upload, verify, search, and archive files — all from
#   a single menu-driven tool designed for IT admins handling leavers and migrations.
#
# REQUIREMENTS
#   • Windows PowerShell 5.1 (built into Windows — no install needed)
#   • PnP.PowerShell module v1.12.0 (installed automatically on first run)
#   • Internet connection + a Microsoft 365 account with access to the SharePoint sites
#
# HOW TO RUN
#   Right-click the .ps1 file → "Run with PowerShell"
#   OR in PowerShell: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
#                     .\RobustHybridFileManager_v92.ps1
#
# FIVE MOST COMMON WORKFLOWS
# ─────────────────────────────────────────────────────────────────────────────────────
#
#  1. COPY/MOVE FILES INSIDE SHAREPOINT (same site)
#     Use options [1] or [2]. Both source and destination must be on the same SharePoint
#     site. This is a SERVER-SIDE operation — your PC is not used as a relay, so it is
#     very fast even for huge folders.
#     How to get the URL: Open SharePoint in your browser → navigate to the folder →
#     copy the full address bar URL → paste it when prompted.
#
#  2. DOWNLOAD SHAREPOINT FILES TO YOUR PC
#     Use option [13] for a full folder, [14] to pick specific items.
#     You will need: the SharePoint URL, a local destination folder, and a temp folder.
#     Files are downloaded to a temp location first, then moved — no partial files.
#     If interrupted, re-run with the same paths to resume (already-downloaded files skip).
#
#  3. UPLOAD LOCAL FILES TO SHAREPOINT
#     Use option [15] for a full folder, [16] to pick specific items.
#     Sub-folders are created automatically. Same smart-skip logic as downloads.
#
#  4. ARCHIVE CLOUD FOLDERS (download + verify + delete)
#     Use option [17] for a single folder, [18] to pick items from a folder, or [19]
#     to queue multiple sources in one run.
#     The script downloads everything, verifies file counts and byte totals match, then
#     deletes from SharePoint ONLY if verification passes. Deleted items go to the
#     SharePoint Recycle Bin and remain recoverable for ~93 days.
#
#  5. ARCHIVE A LEAVER'S ONEDRIVE (disabled account / no licence)
#     Use option [20]. When a user's licence is removed, you cannot access their OneDrive
#     via a standard URL. This option grants your admin account temporary Site Collection
#     Admin rights on their personal site, downloads all document libraries, verifies,
#     and then revokes the admin grant when done.
#     You need: the leaver's email (UPN) OR their OneDrive URL, and your SPO Admin URL
#     (e.g. https://yourcompany-admin.sharepoint.com).
#
#  5. CROSS-COMPANY / CROSS-SITE TRANSFER
#     Direct cloud-to-cloud transfers between different tenants are blocked by Microsoft
#     security. The workaround: use [13/14] to download from Site A, then [15/16] to
#     upload to Site B. Your PC acts as a temporary relay.
#
# LOGS
#   Every operation is logged to:
#   C:\Users\Public\Documents\RobustCloudWork\Logs\HybridJobLog_<username>.log
#   Open the log folder at any time with [L] from the main menu.
#
# KEYBOARD SHORTCUTS (during any running job)
#   [P] Pause the current job    [C] Cancel immediately
#   [F] Open Favourites picker   [R] Repeat last-used path
#
# ========================================================================================

# --- POWERSHELL LESSON: What is a Script? ---
# Think of a script like a cooking recipe. The computer reads it from top to bottom.
# Before we can "cook" (move files), we have to gather our ingredients and set ground rules.

# ========================================================================================
# SECTION 1 — GLOBAL PARAMETERS (script settings you can customise)
# ========================================================================================
# WorkDir       : Root folder for logs and favourites. Default is a shared Public folder
#                 so all users on this PC share the same log history.
# MaxLogSizeMB  : Each log file rotates (gets renamed and a fresh log starts) when it
#                 reaches this size. Prevents logs from growing unboundedly.
# CloudPathLimit: SharePoint enforces a maximum URL length. Files whose full server-relative
#                 path exceeds this limit are skipped and reported in LongPaths*.txt.
#                 400 chars is a safe working limit (SharePoint's hard limit is higher but
#                 includes the domain, so this leaves headroom).
param (
    [string]$WorkDir    = "C:\Users\Public\Documents\RobustCloudWork",
    [int]$MaxLogSizeMB  = 5,
    [int]$CloudPathLimit = 400
)

$env:PNPPOWERSHELL_UPDATECHECK = 'Off'

# ========================================================================================
# SECTION 2 — RUNTIME SETUP (log files, paths, user isolation)
# ========================================================================================
# Log files are named per-user so multiple staff running the script simultaneously
# do not write to the same file, avoiding lock conflicts on shared machines.
$LogFileDirectory = Join-Path $WorkDir "Logs"

if (-not (Test-Path -LiteralPath $LogFileDirectory)) { New-Item -ItemType Directory -Path $LogFileDirectory -Force | Out-Null }

$ActiveUser      = $env:USERNAME
$LogFile         = Join-Path $LogFileDirectory "HybridJobLog_$ActiveUser.log"
$LongPathReport  = Join-Path $LogFileDirectory "LongPaths_FailedToSync_$ActiveUser.txt"

# --- NEW (v80): Favorites are stored as a simple JSON file in the WorkDir root. ---
# Each entry has: Name (friendly label), Type (Cloud/Local/Any), Value (the URL or path).
$FavoritesFile = Join-Path $WorkDir "Favorites_$ActiveUser.json"

# Microsoft Graph connection state — used by option [23] (Unlicensed Leaver Archive).
# All Graph operations use native REST (Invoke-RestMethod) — no module dependency.
$global:MgGraphConnected = $false
$global:MgAccessToken    = ""         # Bearer token for Graph REST calls
$global:MgRefreshToken   = ""         # Used to silently renew the access token
$global:MgTokenExpiry    = $null      # [DateTime] — auto-refresh 5 min before expiry

# Running total of bytes successfully transferred in this PowerShell session.
# Accumulated by every local copy/move and cloud download operation.
# Displayed in Write-Progress CurrentOperation and in each operation summary box.
$global:SessionTransferredBytes = [long]0

# Files at or above this size get a per-file progress bar (local) or a
# "large file — please wait" indicator (cloud, where streaming isn't available).
$script:LargeFileBytes = 1GB

# Well-known Microsoft Graph PowerShell public client ID (registered by Microsoft,
# device code flow enabled, no app registration needed, works on all M365 tenants).
$script:MgClientId       = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
$script:MgScopes         = "GroupMember.ReadWrite.All User.Read.All Group.Read.All offline_access"
# Token cache file — refresh token encrypted with Windows DPAPI (per-user, per-machine)
$script:MgTokenCacheFile = Join-Path $WorkDir "MgTokenCache_$($env:USERNAME).dat"

# Ensure the WorkDir itself exists (needed for the Favorites file).
if (-not (Test-Path -LiteralPath $WorkDir)) { New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null }

# ========================================================================================
# SECTION 3 — WINDOWS KERNEL BRIDGES (long-path and free-space support)
# ========================================================================================
# EXPLANATION: We compile two small C# bridges to kernel32.dll.
#   Win32:     MoveFileW / DeleteFileW — used for safe temp-file promotion during downloads.
#   Win32Disk: GetDiskFreeSpaceExW     — used for free-space checks on LOCAL and UNC paths.
#              [System.IO.DriveInfo] only understands drive letters (C:\) and throws on UNC.
#              GetDiskFreeSpaceExW is the same API Windows Explorer uses and works on both.
if (-not ("Win32" -as [type])) {
    try {
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern bool MoveFileW(string lpExistingFileName, string lpNewFileName);
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern bool DeleteFileW(string lpFileName);
        }
"@
    } catch { Write-Host "Warning: Could not compile Win32 Kernel bridge. Falling back to native .NET." -ForegroundColor Yellow }
}

if (-not ("Win32Disk" -as [type])) {
    try {
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32Disk {
            // GetDiskFreeSpaceExW works for: C:\folder, D:\, \\server\share, \\server\share\folder
            // lpFreeBytesAvailable = bytes available to the CURRENT USER (respects quotas)
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetDiskFreeSpaceExW(
                string  lpDirectoryName,
                out ulong lpFreeBytesAvailable,
                out ulong lpTotalNumberOfBytes,
                out ulong lpTotalNumberOfFreeBytes);
        }
"@
    } catch { Write-Host "Warning: Could not compile Win32Disk bridge. Free-space checks will use fallback." -ForegroundColor Yellow }
}

# ========================================================================================
# THE TOOLBOX (Functions)
# ========================================================================================

# --- FUNCTION: Format-FileSize ---
function Format-FileSize {
    param([long]$bytes)
    if ($bytes -ge 1GB) { return "$([math]::Round($bytes / 1GB, 2)) GB" }
    if ($bytes -ge 1MB) { return "$([math]::Round($bytes / 1MB, 2)) MB" }
    if ($bytes -ge 1KB) { return "$([math]::Round($bytes / 1KB, 2)) KB" }
    return "$bytes Bytes"
}

# --- FUNCTION: Format-Duration (NEW v109) ---
# PURPOSE: Convert a duration in seconds into a short, human-readable string
# such as "45s", "5m 12s", or "2h 15m". Used for ETA displays in progress bars
# and for "Average Speed" / "Total Time" rows in summary boxes.
#
# WHY NOT JUST USE [TimeSpan].ToString()?
#   TimeSpan ToString() gives "00:05:12" — visually noisy for at-a-glance reading.
#   Format-Duration trims to the largest meaningful unit and drops the rest.
#
# SAFETY: Returns "" for invalid inputs (negative, NaN, Infinity) so callers
# can safely concatenate the result into a UI line without conditionals.
function Format-Duration {
    param([double]$Seconds)
    if ($Seconds -lt 0)                  { return "" }
    if ([double]::IsNaN($Seconds))       { return "" }
    if ([double]::IsInfinity($Seconds))  { return "" }
    if ($Seconds -lt 1)    { return "<1s" }
    if ($Seconds -lt 60)   { return "{0}s"     -f [int]$Seconds }
    if ($Seconds -lt 3600) { return "{0}m {1}s" -f [int]($Seconds / 60), [int]($Seconds % 60) }
    return "{0}h {1}m" -f [int]($Seconds / 3600), [int](($Seconds % 3600) / 60)
}

# --- FUNCTION: Get-ETA (NEW v109) ---
# PURPOSE: Calculate "estimated time remaining" for a long-running operation
# based on how much has been processed so far and how long it has taken.
#
# Formula: ETA = (remaining_work) / (work_per_second so far)
#
# Returns a human-readable string ("5m 12s") or "" if ETA can't be computed
# (e.g. work just started, total unknown, or already complete).
#
# CAN BE USED FOR EITHER:
#   • File counts:  Get-ETA -Current 50 -Total 200 -Elapsed $sw.Elapsed
#   • Byte counts:  Get-ETA -Current 1GB -Total 10GB -Elapsed $sw.Elapsed
function Get-ETA {
    param(
        [long]$Current,
        [long]$Total,
        [TimeSpan]$Elapsed
    )
    if ($Current -le 0)                    { return "" }   # nothing done yet
    if ($Total  -le 0)                     { return "" }   # total unknown
    if ($Current -ge $Total)               { return "" }   # already complete
    if ($Elapsed.TotalSeconds -lt 1)       { return "" }   # too early to predict
    $ratePerSec = $Current / $Elapsed.TotalSeconds
    if ($ratePerSec -le 0)                 { return "" }
    $remainingSec = ($Total - $Current) / $ratePerSec
    return Format-Duration $remainingSec
}

# --- FUNCTION: Format-AvgSpeed (NEW v109) ---
# PURPOSE: Calculate and format average transfer speed (e.g. "12.5 MB/s") given
# total bytes transferred and total elapsed seconds. Used in summary boxes
# at the end of operations so the operator knows roughly what to expect on
# future runs of similar size.
#
# Returns "" for invalid inputs so callers can safely concatenate.
function Format-AvgSpeed {
    param(
        [long]$Bytes,
        [TimeSpan]$Elapsed
    )
    if ($Bytes -le 0)                  { return "" }
    if ($Elapsed.TotalSeconds -lt 1)   { return "" }
    $bytesPerSec = [long]($Bytes / $Elapsed.TotalSeconds)
    return "$(Format-FileSize $bytesPerSec)/s"
}

# --- FUNCTION: Format-FilesPerMin (NEW v110) ---
# PURPOSE: Calculate and format file processing rate (e.g. "1,250 files/min").
# Complement to Format-AvgSpeed for cases where bytes don't tell the whole story
# (e.g. a million tiny files where the bottleneck is file-open overhead, not
# bandwidth). Operators can use this to plan future runs and spot anomalies
# (a sudden drop in files/min often means cloud throttling kicked in).
#
# Returns "" for invalid inputs so callers can safely concatenate.
function Format-FilesPerMin {
    param(
        [int]$Count,
        [TimeSpan]$Elapsed
    )
    if ($Count -le 0)                  { return "" }
    if ($Elapsed.TotalSeconds -lt 1)   { return "" }
    $rate = [int]($Count / $Elapsed.TotalMinutes)
    if ($rate -le 0)                   { return "" }
    return ("{0:N0} files/min" -f $rate)
}

# --- FUNCTION: Get-PathStorageType (NEW v110) ---
# PURPOSE: Classify a path so the operator can see at a glance what kind of
# storage they're working with. Different storage types have very different
# performance profiles:
#
#   LOCAL    - Internal/external drive directly attached to this machine
#              (fastest; full IOPS available)
#   MAP      - Drive letter that resolves to a UNC path (mapped network drive)
#              (slower; subject to SMB and network latency)
#   UNC      - Bare UNC path "\\server\share\..."
#              (similar to MAP; sometimes auth/credential differences)
#   OneDrive - Local mirror of a OneDrive folder (looks LOCAL but writes
#              trigger background sync; large bursts can lag the cloud)
#
# Result is a short tag (3-8 chars) suitable for inline display in progress
# bars and summary boxes. Returns "" if the path is empty or unclassifiable.
#
# DETECTION ORDER MATTERS: OneDrive sits inside a LOCAL drive letter, so we
# must test the OneDrive substring BEFORE the drive-letter fall-through, or
# every OneDrive folder would be misreported as LOCAL.
function Get-PathStorageType {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }

    # OneDrive (must check before LOCAL — OneDrive sits inside C:\)
    # Matches \OneDrive\, \OneDrive -..., or \OneDrive_... folder names
    if ($Path -match '\\OneDrive(\s|-|_|\\|$)') { return "OneDrive" }

    # \\?\UNC\server\share\... — long UNC path
    if ($Path -match '^\\\\\?\\UNC\\') { return "UNC" }

    # \\server\share\... — standard UNC path
    if ($Path -match '^\\\\[^\\?]') { return "UNC" }

    # \\?\C:\... or C:\... — long local OR mapped drive
    $driveLetter = $null
    if ($Path -match '^\\\\\?\\([A-Za-z]):') { $driveLetter = $matches[1] }
    elseif ($Path -match '^([A-Za-z]):')    { $driveLetter = $matches[1] }

    if ($driveLetter) {
        try {
            $drive = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
            # DisplayRoot is set when the drive letter points to a UNC mapping
            if ($drive -and $drive.DisplayRoot -and $drive.DisplayRoot -match '^\\\\') {
                return "MAP"
            }
        } catch { }
        return "LOCAL"
    }

    return ""
}

# --- FUNCTION: Format-PathWithStorageTag (NEW v110) ---
# PURPOSE: Prefix a local path with its storage-type tag for log lines and
# summary boxes, e.g.:
#   C:\Users\foo\file.txt        ->  [LOCAL] C:\Users\foo\file.txt
#   \\server\share\file.txt      ->  [UNC] \\server\share\file.txt
#   M:\team\file.txt (mapped)    ->  [MAP] M:\team\file.txt
#   C:\Users\foo\OneDrive\file   ->  [OneDrive] C:\Users\foo\OneDrive\file
#
# If the path can't be classified (empty, cloud URL, etc.) the tag is omitted
# and the path is returned unchanged.
function Format-PathWithStorageTag {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $tag = Get-PathStorageType $Path
    if ($tag) { return "[$tag] $Path" }
    return $Path
}

# --- FUNCTION: Get-ItemsWithHeartbeat (NEW v110) ---
# PURPOSE: Enumerate a directory tree while showing a "scanning..." heartbeat
# every two seconds so the operator knows the script is still working during
# very large enumerations (hundreds of thousands of files).
#
# WHY THIS EXISTS:
#   Get-ChildItem -Recurse is a single synchronous call — there's no way to
#   show progress while it runs. Even on fast SSDs, walking a 500k-file tree
#   takes 30-60 seconds with no output. Operators have killed long jobs
#   thinking they were hung.
#
# HOW IT WORKS:
#   Uses [DirectoryInfo]::EnumerateFileSystemInfos which yields items one at
#   a time as the OS walks the tree. We accumulate them into a List and emit
#   a Write-Host heartbeat on a 2-second timer.
#
# OUTPUT:
#   Returns a List<FileSystemInfo> (or List<FileInfo> / List<DirectoryInfo>
#   depending on switches). Matches the shape callers expect from
#   Get-ChildItem so the rest of the pipeline doesn't need to change.
#
# FALLBACK:
#   If EnumerateFileSystemInfos throws (e.g. permission errors on a sub-tree),
#   we catch and fall through to Get-ChildItem so the operation still
#   completes — just without the heartbeat for that subtree.
function Get-ItemsWithHeartbeat {
    param(
        [string]$LongPath,
        [string]$Label = "Scanning",
        [switch]$FilesOnly,
        [switch]$DirsOnly
    )

    $list = [System.Collections.Generic.List[System.IO.FileSystemInfo]]::new()
    $heartbeatSw = [System.Diagnostics.Stopwatch]::StartNew()
    $totalSw     = [System.Diagnostics.Stopwatch]::StartNew()
    $count = 0
    $shownHeartbeat = $false

    try {
        $dirInfo = [System.IO.DirectoryInfo]::new($LongPath)
        $enumerator = if ($FilesOnly) {
            $dirInfo.EnumerateFiles("*", [System.IO.SearchOption]::AllDirectories)
        } elseif ($DirsOnly) {
            $dirInfo.EnumerateDirectories("*", [System.IO.SearchOption]::AllDirectories)
        } else {
            $dirInfo.EnumerateFileSystemInfos("*", [System.IO.SearchOption]::AllDirectories)
        }

        foreach ($item in $enumerator) {
            $list.Add($item) | Out-Null
            $count++

            # Heartbeat fires every 2 seconds — quiet scans (small folders that
            # finish in under 2s) emit nothing, so the console stays clean.
            if ($heartbeatSw.ElapsedMilliseconds -gt 2000) {
                $shownHeartbeat = $true
                # \r returns the cursor to column 0; PadRight wipes residue from
                # the previous (potentially longer) message. Stays on one line.
                $msg = "  [..] $Label : {0:N0} items found ({1}s)..." -f $count, [int]$totalSw.Elapsed.TotalSeconds
                Write-Host "`r$($msg.PadRight(72))" -NoNewline -ForegroundColor DarkGray
                $heartbeatSw.Restart()
            }
        }

        # Finalize the heartbeat line with a [OK] summary IF we showed any
        # heartbeats. Otherwise the scan was fast enough that no message
        # appeared and we leave the console alone.
        if ($shownHeartbeat) {
            $finalMsg = "  [OK] $Label complete: {0:N0} items in {1}s" -f $count, [int]$totalSw.Elapsed.TotalSeconds
            Write-Host "`r$($finalMsg.PadRight(72))" -ForegroundColor DarkGray
        }
    } catch {
        # If we got partway through and showed heartbeats, finalize the line
        if ($shownHeartbeat) { Write-Host "" }
        Write-Log "Enumeration via .NET enumerator failed for '$LongPath' — falling back to Get-ChildItem. Error: $($_.Exception.Message)" "WARN"

        # Fall back to Get-ChildItem with the right filter switches
        $params = @{ LiteralPath = $LongPath; Recurse = $true; Force = $true; ErrorAction = "SilentlyContinue" }
        if ($FilesOnly) { $params.File = $true }
        if ($DirsOnly)  { $params.Directory = $true }
        $fallback = Get-ChildItem @params
        if ($fallback) { foreach ($x in $fallback) { $list.Add($x) | Out-Null } }
    }

    return ,$list   # comma-prefix forces array return for empty results
}

# --- FUNCTION: Get-PathFreeSpace ---
# PURPOSE: Returns the available free bytes on ANY path — local (C:\) OR UNC (\\server\share).
#
# HOW IT WORKS:
#   Primary:  Win32Disk::GetDiskFreeSpaceExW — the same Win32 API Windows Explorer uses.
#             Works on local drive paths AND bare UNC paths with no special-casing.
#             Returns bytes available to the CURRENT USER (respects per-user disk quotas).
#   Fallback: [System.IO.DriveInfo] for local drives on systems where the P/Invoke fails.
#   Returns:  Available bytes as [long], or -1 if the space cannot be determined.
#
function Get-PathFreeSpace {
    param([string]$Path)
    try {
        # Use the path itself if the directory exists; otherwise fall back to the volume root.
        # GetDiskFreeSpaceExW accepts any valid directory path — it does not need to be the root.
        $checkPath = if ([System.IO.Directory]::Exists((Get-LongPath $Path))) { $Path } else {
            [System.IO.Path]::GetPathRoot((Get-LongPath $Path))
        }
        if ("Win32Disk" -as [type]) {
            [ulong]$free = 0; [ulong]$total = 0; [ulong]$totalFree = 0
            if ([Win32Disk]::GetDiskFreeSpaceExW($checkPath, [ref]$free, [ref]$total, [ref]$totalFree)) {
                return [long]$free
            }
        }
    } catch { }
    # Fallback: DriveInfo works for local lettered drives only
    try {
        $root  = [System.IO.Path]::GetPathRoot((Get-LongPath $Path))
        $drive = [System.IO.DriveInfo]::new($root)
        if ($drive.IsReady) { return $drive.AvailableFreeSpace }
    } catch { }
    return [long]-1   # -1 = could not determine
}

# --- FUNCTION: Get-FailureRateLevel (NEW v111) ---
# PURPOSE: Decide whether the failure rate of a completed operation is high
# enough to warrant a warning in the summary box. Returns a level string
# ("WARN", "ERROR", or "") suitable for passing directly to Write-SummaryBox
# as the row Level. Returns "" when the rate is acceptable so callers can
# skip the row entirely.
#
# THRESHOLDS:
#   0%       - no row
#   <10%     - no row (a few failures are normal in large jobs)
#   10%-25%  - "WARN"  yellow row in summary
#   >25%     - "ERROR" red row in summary  (something is clearly wrong)
function Get-FailureRateLevel {
    param([int]$Failed, [int]$Total)
    if ($Failed -le 0 -or $Total -le 0) { return "" }
    $rate = $Failed / $Total
    if ($rate -ge 0.25) { return "ERROR" }
    if ($rate -ge 0.10) { return "WARN" }
    return ""
}

# --- FUNCTION: Format-FailureRate (NEW v111) ---
# PURPOSE: Format a failure rate as "X% (N of M failed)" for display.
function Format-FailureRate {
    param([int]$Failed, [int]$Total)
    if ($Total -le 0) { return "" }
    $pct = [math]::Round(($Failed / $Total) * 100, 1)
    return "$pct% ($Failed of $Total failed)"
}

# --- FUNCTION: Get-TopExtensions (NEW v111) ---
# PURPOSE: Take a list of FileInfo or hashtable items and return the top N
# file extensions with their counts. Helps operators see at a glance what
# kind of content a job processed.
#
# INPUT:    [string[]] of filenames (or items with a Name/File property)
# OUTPUT:   "234 .docx, 89 .pdf, 45 .xlsx, 12 .png, 8 .zip"
#
# Items without an extension are grouped under "(no ext)". Returns "" if
# the input list is empty so callers can skip the row entirely.
function Get-TopExtensions {
    param(
        [System.Collections.IEnumerable]$Items,
        [int]$Top = 5
    )
    if (-not $Items) { return "" }

    # Use a plain hashtable for the counter — works in PS 5.1 without any
    # special collection imports. The key is the extension (lowercased) and
    # the value is the running count.
    $counter = @{}
    $total = 0
    foreach ($item in $Items) {
        $name = $null
        # Try common property names — works for FileInfo and our FailedList hashtables
        if     ($item -is [System.IO.FileSystemInfo]) { $name = $item.Name }
        elseif ($item.Name)                            { $name = $item.Name }
        elseif ($item.File)                            { $name = $item.File }
        elseif ($item -is [string])                    { $name = $item }
        else { continue }

        $ext = [System.IO.Path]::GetExtension($name)
        if ([string]::IsNullOrWhiteSpace($ext)) { $ext = "(no ext)" }
        else { $ext = $ext.ToLowerInvariant() }

        if ($counter.ContainsKey($ext)) { $counter[$ext]++ }
        else                            { $counter[$ext] = 1 }
        $total++
    }
    if ($total -eq 0) { return "" }

    # Sort by count desc, take top N
    $sorted = $counter.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First $Top
    $parts = @()
    foreach ($entry in $sorted) {
        $parts += ("{0:N0} {1}" -f $entry.Value, $entry.Key)
    }
    # If there are more extensions than we showed, indicate the remainder
    if ($counter.Count -gt $Top) {
        $remaining = $counter.Count - $Top
        $parts += "+$remaining more"
    }
    return ($parts -join ", ")
}

# --- FUNCTION: Get-DistinctErrors (NEW v111) ---
# PURPOSE: Take a FailedList (List<hashtable> with a Reason property) and
# return the distinct error messages with their occurrence counts, sorted
# by frequency descending. Helps triage what went wrong without scrolling
# through hundreds of repeated "the same error happened 234 times" lines
# in the per-file FAILED log.
#
# v111 NOTE: This is a NEW summary recap layered ON TOP of the existing
# Write-FailedReport function (which still dumps every failure with its
# full file path). The recap shows distinct messages so a human can spot
# the underlying pattern at a glance; the full report stays in the log
# for forensics. Both run — nothing was removed.
#
# OUTPUT (string[]):
#   "[x234] HTTP 429 Too Many Requests"
#   "[x12]  Access denied"
#   "[x3]   File locked by another process"
#
# Pass -Top N to cap the list (e.g. -Top 5 for "top 5 errors"). By
# default ALL distinct messages are returned — there's typically only a
# handful of distinct causes even in a job with thousands of failures.
function Get-DistinctErrors {
    param(
        [System.Collections.IEnumerable]$FailedList,
        [int]$Top = 0    # 0 = return all distinct errors (v111 default)
    )
    if (-not $FailedList) { return @() }
    $counter = @{}
    foreach ($entry in $FailedList) {
        $reason = if ($entry.Reason) { "$($entry.Reason)" } else { "(no reason recorded)" }
        # Normalize: collapse whitespace, trim, and truncate to 120 chars so
        # one row in the summary box stays readable.
        $reason = ($reason -replace '\s+', ' ').Trim()
        if ($reason.Length -gt 120) { $reason = $reason.Substring(0, 117) + "..." }

        if ($counter.ContainsKey($reason)) { $counter[$reason]++ }
        else                                { $counter[$reason] = 1 }
    }
    if ($counter.Count -eq 0) { return @() }

    $sorted = $counter.GetEnumerator() | Sort-Object Value -Descending
    if ($Top -gt 0) { $sorted = $sorted | Select-Object -First $Top }

    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $sorted) {
        # The "[x123]" prefix lines up vertically thanks to PadRight
        $prefix = ("[x{0}]" -f $entry.Value).PadRight(7)
        $lines.Add("$prefix $($entry.Key)") | Out-Null
    }
    return $lines.ToArray()
}

# --- FUNCTION: Show-OperationDiagnostics (NEW v111) ---
# PURPOSE: After every transfer operation, emit a small diagnostic panel just
# below the main summary box showing:
#   1) Failure rate alert  (only if Failed/Total > 10%)
#   2) Distinct error messages with counts  (only if there were failures)
#   3) Per-extension breakdown  (only if SuccessCount >= 50)
#
# The full per-file Write-FailedReport (which lists EVERY failure with its
# full path) is still emitted separately by each engine — these diagnostics
# are additive context, not a replacement.
#
# WHY THIS SITS OUTSIDE THE SUMMARY BOX:
#   The summary box uses a fixed 64-char width with vertical borders. Long
#   error messages and 5-line breakdowns don't fit cleanly inside without
#   wrapping or truncation. A separate panel keeps both readable.
#
# PARAMETERS:
#   -SuccessCount / -FailedCount  : counts for failure rate calculation
#   -FailedList                   : List<hashtable> from the engine (Reason property)
#   -ProcessedItems               : List of FileInfo or items with Name/File
#                                   property (for extension breakdown).
#                                   Pass $null to skip the breakdown.
function Show-OperationDiagnostics {
    param(
        [int]$SuccessCount,
        [int]$FailedCount,
        [System.Collections.IEnumerable]$FailedList,
        [System.Collections.IEnumerable]$ProcessedItems = $null
    )
    $totalProcessed = $SuccessCount + $FailedCount
    if ($totalProcessed -le 0) { return }

    $panelLines  = [System.Collections.Generic.List[string]]::new()
    $panelColor  = "DarkGray"

    # --- 1. Failure rate alert ---
    $rateLevel = Get-FailureRateLevel -Failed $FailedCount -Total $totalProcessed
    if ($rateLevel) {
        $rateText = Format-FailureRate -Failed $FailedCount -Total $totalProcessed
        $panelLines.Add("[!] FAILURE RATE: $rateText") | Out-Null
        if ($rateLevel -eq "ERROR")     { $panelColor = "Red" }
        elseif ($rateLevel -eq "WARN")  { $panelColor = "Yellow" }
    }

    # --- 2. Distinct errors recap (ALL by default — see Get-DistinctErrors) ---
    if ($FailedList -and @($FailedList).Count -gt 0) {
        $errorLines = Get-DistinctErrors -FailedList $FailedList
        if ($errorLines -and $errorLines.Count -gt 0) {
            if ($panelLines.Count -gt 0) { $panelLines.Add("") | Out-Null }
            $panelLines.Add("DISTINCT ERROR MESSAGES (sorted by frequency):") | Out-Null
            foreach ($el in $errorLines) { $panelLines.Add("  $el") | Out-Null }
            $panelLines.Add("(See log file and FAILED report above for the full per-file list.)") | Out-Null
        }
    }

    # --- 3. Per-extension breakdown (only for batches large enough to matter) ---
    if ($ProcessedItems -and $SuccessCount -ge 50) {
        $extLine = Get-TopExtensions -Items $ProcessedItems -Top 6
        if ($extLine) {
            if ($panelLines.Count -gt 0) { $panelLines.Add("") | Out-Null }
            $panelLines.Add("BY FILE TYPE: $extLine") | Out-Null
        }
    }

    # --- Emit the panel if we have anything to show ---
    if ($panelLines.Count -eq 0) { return }

    # v111: header reflects what's in the panel — "DIAGNOSTICS" implies a problem,
    # which is appropriate when failures exist. For clean runs that only show the
    # informational per-extension breakdown, use the gentler "OPERATION DETAILS".
    $header = if ($FailedCount -gt 0) { "OPERATION DIAGNOSTICS" } else { "OPERATION DETAILS" }

    Write-Host ""
    Write-Host "  ─── $header ──────────────────────────────────" -ForegroundColor $panelColor
    foreach ($line in $panelLines) { Write-Host "  $line" -ForegroundColor $panelColor }
    Write-Host "  ────────────────────────────────────────────────────────────" -ForegroundColor $panelColor
}

# --- FUNCTION: Show-StorageTip (NEW v111) ---
# PURPOSE: Emit a one-time tip when an operation involves a slower storage
# type (UNC or mapped network drive). Helps set operator expectations on
# throughput before the job starts, so a slow run isn't misread as a hang.
#
# Uses a script-level HashSet ($script:StorageTipShown) keyed by the storage
# type so the tip only fires once per session per type — not on every job.
# This avoids spam when the operator runs many jobs against the same drive.
function Show-StorageTip {
    param([string]$Path, [string]$Role = "source")
    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    $type = Get-PathStorageType $Path
    if ($type -notin @("UNC","MAP")) { return }

    if (-not $script:StorageTipShown) {
        $script:StorageTipShown = [System.Collections.Generic.HashSet[string]]::new()
    }
    if ($script:StorageTipShown.Contains($type)) { return }
    $script:StorageTipShown.Add($type) | Out-Null

    Write-Host ""
    Write-Host "  [TIP] $Role is on a network drive ($type)." -ForegroundColor Yellow
    Write-Host "        Expect 30-50% slower throughput than a local drive — this is normal."  -ForegroundColor DarkGray
    Write-Host "        (This tip shows once per session per drive type.)" -ForegroundColor DarkGray
    Write-Host ""
}

# --- FUNCTION: Show-PressAnyKeyToContinue (NEW v111) ---
# PURPOSE: Pause execution until the operator presses any key. Used after
# every menu operation completes (or aborts) so the operator can read the
# final summary box before the main menu redraws and scrolls it away.
#
# Why this matters: prior to v111, when an operation finished or failed it
# returned to the main menu immediately. If the failure summary scrolled
# past in the same render, the operator had no chance to see what happened.
# Now every menu branch ends with this pause, so the post-mortem stays
# on screen until the operator dismisses it.
#
# FALLBACK: [System.Console]::ReadKey() throws under non-interactive hosts
# (Task Scheduler hidden window, some IDE consoles). In those cases we fall
# back to Read-Host so the script doesn't crash. Under Task Scheduler the
# UserInteractive check upstream should prevent us reaching this anyway.
function Show-PressAnyKeyToContinue {
    Write-Host ""
    Write-Host "  Press any key to return to the main menu..." -ForegroundColor DarkGray
    try {
        [void][System.Console]::ReadKey($true)
    } catch {
        $null = Read-Host
    }
}

# --- FUNCTION: Write-Log ---
function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [switch]$ConsoleOnly, [switch]$FileOnly)

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry  = "[$Timestamp] [$Level] $Message"

    if (-not $FileOnly) {
        if    ($Level -eq "ERROR")   { Write-Host $LogEntry -ForegroundColor Red     }
        elseif($Level -eq "WARN")    { Write-Host $LogEntry -ForegroundColor Yellow  }
        elseif($Level -eq "SUCCESS") { Write-Host $LogEntry -ForegroundColor Green   }
        else                         { Write-Host $LogEntry -ForegroundColor Cyan    }
    }

    if ($ConsoleOnly) { return }

    $logRetry = 0
    while ($logRetry -lt 3) {
        try {
            if (Test-Path -LiteralPath $LogFile) {
                $fileInfo = [System.IO.FileInfo]::new($LogFile)
                if ($fileInfo.Length -gt ($MaxLogSizeMB * 1MB)) {
                    $RotationTime = Get-Date -Format "yyyyMMdd_HHmmss"
                    $RotatedLog   = Join-Path $LogFileDirectory "HybridJobLog_$($env:USERNAME)_$RotationTime.log"
                    Rename-Item -Path $LogFile -NewName (Split-Path $RotatedLog -Leaf) -ErrorAction Stop
                }
            }
            [System.IO.File]::AppendAllText($LogFile, "$LogEntry`r`n")
            break
        } catch {
            $logRetry++
            Start-Sleep -Milliseconds 100
        }
    }
}

# --- FUNCTION: Write-SummaryBox ---
# PURPOSE: Renders a clean Unicode-bordered summary panel to both the console and the log file.
# Used by Grand Batch Summary, Verification Results, and Download Complete blocks throughout
# the script — replacing the plain ========= lines with a professional result panel.
#
# USAGE EXAMPLE:
#   Write-SummaryBox -Title "GRAND BATCH SUMMARY" -Lines @(
#       @{ Label = "Successfully Processed"; Value = "47";        Level = "SUCCESS" },
#       @{ Label = "Skipped (already synced)"; Value = "3 files"; Level = "INFO"   },
#       @{ Label = "Failed";                 Value = "0";         Level = "SUCCESS" },
#       @{ Label = "Total Time";             Value = "0h 2m 14s"; Level = "INFO"   }
#   )
#
function Write-SummaryBox {
    param(
        [string]$Title,
        [hashtable[]]$Lines,         # Each entry: @{ Label="..."; Value="..."; Level="INFO|SUCCESS|ERROR|WARN" }
        [string]$BoxColor = "Cyan"   # Border colour
    )

    $width = 64  # inner content width

    $top  = "  ╔" + ("═" * $width) + "╗"
    $mid  = "  ╠" + ("═" * $width) + "╣"
    $bot  = "  ╚" + ("═" * $width) + "╝"

    # Centre the title inside the box width
    $pad        = [math]::Max(0, ($width - $Title.Length) / 2)
    $titleInner = (" " * [int]$pad) + $Title
    $titleInner = $titleInner.PadRight($width)
    $titleLine  = "  ║ $titleInner║"

    Write-Host $top -ForegroundColor $BoxColor
    Write-Host $titleLine -ForegroundColor $BoxColor
    Write-Host $mid -ForegroundColor $BoxColor

    $ts       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logParts = [System.Collections.Generic.List[string]]::new()
    $logParts.Add("[$ts] [INFO] ╔═ $Title")

    foreach ($row in $Lines) {
        # Defensive: callers sometimes pass an inline conditional that evaluates to $null
        # (e.g. "$(if($x) { @{...} })" — when $x is false the array element is $null).
        # Skipping those produces a clean box rather than an empty padded row.
        if ($null -eq $row) { continue }
        $label = if ($row.Label) { "$($row.Label)" } else { "" }
        $value = if ($row.Value) { "$($row.Value)" } else { "" }
        $level = if ($row.Level) { $row.Level } else { "INFO" }

        # Build "  Label ................. Value" padded to box width
        $gap    = $width - $label.Length - $value.Length - 4
        $dots   = if ($gap -gt 1) { " " + ("·" * ($gap - 1)) + " " } else { " " }
        $inner  = "  $label$dots$value"
        $inner  = $inner.PadRight($width)
        $boxLine = "  ║ $inner║"

        $colour = switch ($level) {
            "SUCCESS" { "Green"  }
            "ERROR"   { "Red"    }
            "WARN"    { "Yellow" }
            default   { "Cyan"   }
        }
        Write-Host $boxLine -ForegroundColor $colour
        $logParts.Add("[$ts] [$level]  $label $value")
    }

    Write-Host $bot -ForegroundColor $BoxColor
    $logParts.Add("[$ts] [INFO] ╚" + ("═" * ($width + 2)))

    # Write entire summary to log in one shot (Bookend Audit pattern from v60)
    $payload  = ($logParts -join "`r`n") + "`r`n"
    $retry    = 0
    while ($retry -lt 3) {
        try { [System.IO.File]::AppendAllText($LogFile, $payload); break }
        catch { $retry++; Start-Sleep -Milliseconds 100 }
    }
}


# --- FUNCTION: Write-FailedReport ---
# PURPOSE: Renders a red-bordered failed items panel to console and log file.
# Called whenever a batch or transfer ends with one or more failures.
function Write-FailedReport {
    param([object]$FailedList)
    if (-not $FailedList -or $FailedList.Count -eq 0) { return }

    $width   = 64
    $top     = "  ╔" + ("═" * $width) + "╗"
    $titleIn = ("  FAILED ITEMS  ($($FailedList.Count) item$(if($FailedList.Count -ne 1){'s'}))").PadRight($width)
    $bot     = "  ╚" + ("═" * $width) + "╝"

    Write-Host $top  -ForegroundColor Red
    Write-Host "  ║ $titleIn║" -ForegroundColor Red

    $ts      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logBuf  = [System.Collections.Generic.List[string]]::new()
    $logBuf.Add("[$ts] [ERROR] ╔═ FAILED ITEMS ($($FailedList.Count))")

    foreach ($f in $FailedList) {
        $fileName = if ($f.File) { "$($f.File)" } else { "Unknown" }
        $reason   = if ($f.Reason) { "$($f.Reason)" } else { "Unknown error" }
        # Truncate for display
        $dispName = if ($fileName.Length -gt 40) { $fileName.Substring(0,37) + "..." } else { $fileName }
        $dispRsn  = if ($reason.Length  -gt ($width - $dispName.Length - 5)) {
            $reason.Substring(0, [math]::Max(0, $width - $dispName.Length - 8)) + "..."
        } else { $reason }

        Write-Host "  ║  ▸ $dispName" -ForegroundColor Red
        Write-Host "  ║    Reason: $dispRsn" -ForegroundColor DarkRed
        $logBuf.Add("[$ts] [ERROR]  FILE: $fileName  |  REASON: $reason")
    }

    Write-Host $bot -ForegroundColor Red
    $logBuf.Add("[$ts] [ERROR] ╚═ END FAILED ITEMS")
    $payload = ($logBuf -join "`r`n") + "`r`n"
    $retry = 0
    while ($retry -lt 3) {
        try { [System.IO.File]::AppendAllText($LogFile, $payload); break }
        catch { $retry++; Start-Sleep -Milliseconds 100 }
    }
}


# ========================================================================================
# THROTTLE-RESILIENT PNP WRAPPER
# ========================================================================================

# --- FUNCTION: Invoke-PnPWithRetry ---
# PURPOSE: A universal retry wrapper for any PnP PowerShell API call that might fail
# due to Microsoft throttling (HTTP 429 Too Many Requests or 503 Service Unavailable).
#
# HOW IT WORKS:
#   Microsoft's SharePoint Online API enforces rate limits. When you exceed them it returns
#   HTTP 429 with a "Retry-After" header telling you how many seconds to wait. If we ignore
#   this and retry immediately, the request is rejected again and the wait gets longer.
#   This function:
#     1. Executes the provided ScriptBlock (e.g. { Move-PnPFile ... })
#     2. If it throws, checks whether the error message contains 429/503/throttle keywords
#     3. Waits the Retry-After value from the error if found, otherwise uses exponential
#        backoff: 5s → 10s → 15s → 20s → 25s across up to 5 attempts
#     4. Returns $true on success, $false after all retries are exhausted
#
# USAGE EXAMPLE:
#   $ok = Invoke-PnPWithRetry -ScriptBlock { Move-PnPFile -SourceUrl $src -TargetUrl $dst -Force -Overwrite }
#   if (-not $ok) { Write-Log "Move failed after retries." "ERROR" }
#
# PS5.1 NOTE: ScriptBlock parameter type is [scriptblock] — fully supported in PS5.1.
# The -ErrorAction Stop inside the ScriptBlock is important: without it, non-terminating
# errors won't be caught by the try/catch.
function Invoke-PnPWithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName = "PnP Operation",
        [int]$MaxRetries = 5,
        # Regex patterns that mean the error is permanent — skip all retries immediately.
        # Examples: SharePoint Forms folder, checked-out files. Caller handles the throw.
        [string[]]$FastFailPatterns = @()
    )

    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            & $ScriptBlock
            return $true   # Success — exit immediately
        } catch {
            $attempt++
            $errMsg = $_.Exception.Message

            # Fast-fail: permanent errors that will never succeed on retry
            foreach ($pattern in $FastFailPatterns) {
                if ($errMsg -match $pattern) {
                    throw $_   # Re-throw immediately — caller handles it
                }
            }

            if ($attempt -ge $MaxRetries) {
                # All retries exhausted — let the caller handle the failure
                throw $_
            }

            # Check for throttling signals in the error message
            $isThrottle = ($errMsg -match "429|503|throttl|Too Many Requests|Service Unavailable|Retry-After")

            # Try to extract a Retry-After value from the error (e.g. "Retry after 30 seconds")
            $waitSecs = $attempt * 5   # default: exponential backoff 5s/10s/15s/20s/25s
            if ($errMsg -match "(?i)retry.{0,15}?(\d+)") {
                $parsed = [int]$matches[1]
                # Sanity-check: only use the parsed value if it's between 1 and 120 seconds
                if ($parsed -ge 1 -and $parsed -le 120) { $waitSecs = $parsed }
            }

            if ($isThrottle) {
                Write-Host "   [Throttled] $OperationName — waiting ${waitSecs}s before retry $attempt/$MaxRetries..." -ForegroundColor Yellow
                Write-Log "$OperationName throttled (attempt $attempt/$MaxRetries). Waiting ${waitSecs}s." "WARN"
            } else {
                Write-Host "   [Retry] $OperationName failed (attempt $attempt/$MaxRetries). Retrying in ${waitSecs}s..." -ForegroundColor Yellow
                Write-Log "$OperationName failed (attempt $attempt/$MaxRetries): $errMsg. Retrying in ${waitSecs}s." "WARN"
            }
            Start-InterruptibleSleep -Seconds $waitSecs -Label "Retry in ${waitSecs}s"
            if ($global:JobCancelled) { throw [System.OperationCanceledException]"Cancelled by user during retry wait." }
        }
    }
    return $false
}

# --- FUNCTION: Start-InterruptibleSleep ---
# PURPOSE: A hotkey-aware replacement for bare Start-Sleep used inside retry wait periods.
# Instead of blocking for the full wait duration, this sleeps in 1-second slices and
# calls Check-Interrupt each second. This means [P] and [C] work even while the script
# is waiting between retry attempts (e.g. after a throttled API call or a delete failure).
#
# BEHAVIOUR:
#   - Sleeps in 1-second chunks for $Seconds total.
#   - Calls Check-Interrupt after each chunk. If the user presses [C] or pauses then
#     cancels, $global:JobCancelled is set to $true and this function returns immediately.
#   - Any remaining wait time is abandoned once JobCancelled is true.
#   - Shows a per-second countdown so the user can see how long the wait is.
#
# PARAMETERS:
#   Seconds     : Total seconds to wait.
#   Label       : Short description shown in the countdown bar (e.g. "Retry in 10s").
function Start-InterruptibleSleep {
    param([int]$Seconds, [string]$Label = "Waiting")
    for ($s = $Seconds; $s -gt 0; $s--) {
        if ($global:JobCancelled) { return }
        Write-Host "   $Label — $s second$(if($s -ne 1){'s'}) remaining... [P]=Pause  [C]=Cancel" `
            -ForegroundColor DarkGray -NoNewline
        Write-Host "`r" -NoNewline
        Start-Sleep -Seconds 1
        Check-Interrupt -CurrentFolder "" -CurrentFile "" | Out-Null
        if ($global:JobCancelled) { Write-Host ""; return }
    }
    Write-Host (" " * 72) + "`r" -NoNewline   # clear the countdown line
}

# --- FUNCTION: Invoke-PnPWithRetry ---
# ========================================================================================

# --- FUNCTION: Write-PathBanner ---
# PURPOSE: Displays a compact source/destination/buffer path banner at the top of every
# batch index menu refresh so users always know which operation is currently in progress.
# Called at the start of every while($true) loop in all five batch functions.
# Parameters accept empty strings — labels with empty paths are silently omitted.
function Write-PathBanner {
    param(
        [string]$Label1 = "Source",      [string]$Path1 = "",
        [string]$Label2 = "Destination", [string]$Path2 = "",
        [string]$Label3 = "Buffer",      [string]$Path3 = "",
        [string]$BoxColor = "DarkCyan"
    )
    $rows = @()
    if ($Path1) { $rows += ,@($Label1, $Path1) }
    if ($Path2) { $rows += ,@($Label2, $Path2) }
    if ($Path3) { $rows += ,@($Label3, $Path3) }
    if ($rows.Count -eq 0) { return }
    Write-Host ""
    Write-Host "  ┌─ ACTIVE OPERATION ─────────────────────────────────────────" -ForegroundColor $BoxColor
    foreach ($row in $rows) {
        $lbl  = $row[0].PadRight(11)
        $path = $row[1]
        Write-Host "  │  $lbl $path" -ForegroundColor $BoxColor
    }
    Write-Host "  └────────────────────────────────────────────────────────────" -ForegroundColor $BoxColor
}

# ========================================================================================
# FAVOURITES & BOOKMARKS SYSTEM
# ========================================================================================

# --- FUNCTION: Get-Favorites ---
# Loads the saved favorites list from disk as an array of PSCustomObjects.
#
# POWERSHELL 5 LESSON: Why PSCustomObject instead of Hashtable?
# ConvertFrom-Json natively returns PSCustomObjects. Converting them to hashtables (as v80 did)
# introduces a subtle bug: when hashtable[] is piped through Where-Object, PowerShell 5 can
# resolve $_.Type to the built-in .GetType() method instead of the hashtable key named "Type".
# Keeping everything as PSCustomObject eliminates this ambiguity entirely — property access with
# dot notation is always unambiguous on PSCustomObjects.
function Get-Favorites {
    if (-not (Test-Path -LiteralPath $FavoritesFile)) { return @() }
    try {
        $raw = [System.IO.File]::ReadAllText($FavoritesFile)
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

        # Use -InputObject (not pipeline) and wrap in @() to guarantee we always get an array,
        # even when the JSON file contains only a single object instead of a JSON array.
        $parsed = @(ConvertFrom-Json -InputObject $raw)

        # Normalize each entry to a clean PSCustomObject with guaranteed string properties.
        # "$($item.Prop)" converts null safely to empty string without throwing.
        $normalized = foreach ($item in $parsed) {
            [PSCustomObject]@{
                Name  = "$($item.Name)"
                Type  = "$($item.Type)"
                Value = "$($item.Value)"
                Added = "$($item.Added)"
            }
        }
        return @($normalized)
    } catch {
        Write-Host "Warning: Could not read favorites file. Starting fresh." -ForegroundColor Yellow
        return @()
    }
}

# --- FUNCTION: Save-Favorites ---
# Writes the current favourites list back to disk as a JSON array.
#
# POWERSHELL 5 SERIALIZATION FIX (v83):
# ConvertTo-Json -InputObject @($array) in PS5 has a bug where passing an object[] containing
# multiple PSCustomObjects can produce a single merged JSON object rather than a JSON array.
# This caused all favourites to be concatenated into one corrupted entry (the v81/v82 bug).
#
# THE FIX: Serialize each favourite individually using ConvertTo-Json -Compress -Depth 1
# (which reliably serializes a SINGLE object as "{"key":"value"}"), then manually wrap
# the comma-joined results in "[" and "]" to build a well-formed JSON array string.
# This completely bypasses the PS5 multi-object array serialization quirk.
function Save-Favorites {
    param([object[]]$Favorites)
    try {
        if ($null -eq $Favorites -or $Favorites.Count -eq 0) {
            [System.IO.File]::WriteAllText($FavoritesFile, "[]")
            return
        }

        # Serialize each entry as a standalone compact JSON object, then join into an array.
        $jsonParts = [System.Collections.Generic.List[string]]::new()
        foreach ($f in $Favorites) {
            # Rebuild as a clean PSCustomObject with guaranteed string properties before serializing.
            $clean = [PSCustomObject]@{
                Name  = "$($f.Name)"
                Type  = "$($f.Type)"
                Value = "$($f.Value)"
                Added = "$($f.Added)"
            }
            # -Compress produces a single-line {"key":"val",...} — safe, predictable, no array quirks.
            $jsonParts.Add(($clean | ConvertTo-Json -Compress -Depth 1))
        }

        $json = "[`r`n  " + ($jsonParts -join ",`r`n  ") + "`r`n]"
        [System.IO.File]::WriteAllText($FavoritesFile, $json)
    } catch {
        Write-Host "Warning: Could not save favourites. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# --- FUNCTION: Show-FavoritePicker ---
# Displays an interactive numbered list of saved favorites and returns the chosen value.
# Pass -Type "Cloud" or -Type "Local" to show relevant favorites first, or omit for all.
# Returns the selected Value string, or $null if the user cancels.
#
# POWERSHELL 5 NOTE: Filtering is done with a manual for-loop (NOT Where-Object pipeline).
# This is intentional — even though Get-Favorites now returns PSCustomObjects, using
# a for-loop with direct index access is the most reliable approach across all PS versions.
function Show-FavoritePicker {
    param([string]$Type = "Any")

    $allFavs = @(Get-Favorites)

    # Build filtered list using a for-loop to avoid any pipeline type resolution quirks.
    $filtered = [System.Collections.Generic.List[PSCustomObject]]::new()
    for ($i = 0; $i -lt $allFavs.Count; $i++) {
        $f = $allFavs[$i]
        if ($Type -eq "Any" -or $f.Type -eq $Type -or $f.Type -eq "Any") {
            $filtered.Add($f)
        }
    }

    # Fallback: if the type filter produced no matches, show everything rather than nothing.
    if ($filtered.Count -eq 0 -and $allFavs.Count -gt 0) {
        for ($i = 0; $i -lt $allFavs.Count; $i++) { $filtered.Add($allFavs[$i]) }
    }

    if ($filtered.Count -eq 0) {
        Write-Host ""
        Write-Host " No favorites saved yet. Use [F] from the Main Menu to add some." -ForegroundColor Yellow
        Write-Host ""
        return $null
    }

    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║               FAVOURITES PICKER                     ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    if ($Type -ne "Any") { Write-Host "  Showing $Type favourites first (+ Any type)." -ForegroundColor DarkGray }
    Write-Host ""

    for ($i = 0; $i -lt $filtered.Count; $i++) {
        $f = $filtered[$i]
        # Access .Type on PSCustomObject — always unambiguous, never conflicts with .GetType()
        $typeLabel = switch ($f.Type) {
            "Cloud" { " CLOUD " }
            "Local" { " LOCAL " }
            default { "  ANY  " }
        }
        $typeColor = switch ($f.Type) {
            "Cloud" { "Yellow" }
            "Local" { "Green"  }
            default { "White"  }
        }
        Write-Host "  [$($i+1)]" -ForegroundColor Cyan -NoNewline
        Write-Host " [$typeLabel]" -ForegroundColor $typeColor -NoNewline
        Write-Host " $($f.Name)" -ForegroundColor White
        $displayVal = if ($f.Value.Length -gt 55) { $f.Value.Substring(0,52) + "..." } else { $f.Value }
        Write-Host "        $displayVal" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "  ──────────────────────────────────────────────────────" -ForegroundColor Cyan
    $pick = Read-Host "  Enter a number, or ENTER to cancel"

    if ([string]::IsNullOrWhiteSpace($pick)) { return $null }

    if ($pick -match '^\d+$' -and [int]$pick -gt 0 -and [int]$pick -le $filtered.Count) {
        $chosen = $filtered[[int]$pick - 1]
        Write-Host "  ✔ Using: $($chosen.Name)" -ForegroundColor Green
        Write-Log "[FAVOURITES] User selected: '$($chosen.Name)' => '$($chosen.Value)'" "INFO" -FileOnly
        return $chosen.Value
    }

    Write-Host "  Invalid selection." -ForegroundColor Red
    return $null
}

# --- FUNCTION: Manage-Favorites ---
# Full management UI: list, add, rename, delete favorites.
# All entries are stored and manipulated as PSCustomObjects for consistency with Get-Favorites.
function Manage-Favorites {
    while ($true) {
        $favs = @(Get-Favorites)

        Write-Host ""
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "  ║               FAVOURITES MANAGER                    ║" -ForegroundColor Cyan
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host "  Saved favourites are available at any URL/path prompt by typing [F]." -ForegroundColor DarkGray
        Write-Host ""

        if ($favs.Count -eq 0) {
            Write-Host "  ┌─ No favourites saved yet ───────────────────────────" -ForegroundColor DarkGray
            Write-Host "  │  Use [A] to add your first favourite." -ForegroundColor DarkGray
            Write-Host "  └─────────────────────────────────────────────────────" -ForegroundColor DarkGray
        } else {
            Write-Host "  ┌─ Saved Favourites ─────────────────────────────────" -ForegroundColor Cyan
            for ($i = 0; $i -lt $favs.Count; $i++) {
                $f = $favs[$i]
                $typeLabel = switch ($f.Type) {
                    "Cloud" { " CLOUD " }
                    "Local" { " LOCAL " }
                    default { "  ANY  " }
                }
                $typeColor = switch ($f.Type) {
                    "Cloud" { "Yellow"  }
                    "Local" { "Green"   }
                    default { "White"   }
                }
                Write-Host "  │  [$($i+1)]" -ForegroundColor Cyan -NoNewline
                Write-Host " [$typeLabel]" -ForegroundColor $typeColor -NoNewline
                Write-Host " $($f.Name)" -ForegroundColor White
                # Truncate long values for display
                $dispVal = if ($f.Value.Length -gt 55) { $f.Value.Substring(0,52) + "..." } else { $f.Value }
                Write-Host "  │        $dispVal" -ForegroundColor DarkGray
                Write-Host "  │        Added: $($f.Added)" -ForegroundColor DarkGray
                if ($i -lt $favs.Count - 1) { Write-Host "  │" -ForegroundColor Cyan }
            }
            Write-Host "  └─────────────────────────────────────────────────────" -ForegroundColor Cyan
        }

        Write-Host ""
        Write-Host "  [A] Add      [R] Rename     [E] Edit Value     [D] Delete     [Q] Back" -ForegroundColor White
        Write-Host ""
        $action = Read-Host "  Action"
        Write-Log "[USER INPUT] Favourites Manager Action: '$action'" "INFO" -FileOnly

        # --- ADD ---
        if ($action -match '^[Aa]$') {
            Write-Host ""
            Write-Host "  ┌─ Add New Favourite ────────────────────────────────" -ForegroundColor Green
            Write-Host "  │  What type of path is this?" -ForegroundColor White
            Write-Host "  │  [1] Cloud URL   (e.g. https://yourorg.sharepoint.com/...)" -ForegroundColor DarkGray
            Write-Host "  │  [2] Local Path  (e.g. C:\Archives\2024  or  \\Server\Share)" -ForegroundColor DarkGray
            Write-Host "  │  [3] Either      (appears in all pickers)" -ForegroundColor DarkGray
            Write-Host "  └─────────────────────────────────────────────────────" -ForegroundColor Green
            $typeChoice = Read-Host "  Type [1/2/3]"
            Write-Log "[USER INPUT] Favourites Add Type: '$typeChoice'" "INFO" -FileOnly

            $favType = switch ($typeChoice) {
                "1"     { "Cloud" }
                "2"     { "Local" }
                "3"     { "Any"   }
                default { Write-Host "  Defaulting to Any." -ForegroundColor Yellow; "Any" }
            }

            # Enter the URL/path FIRST, then suggest a name derived from it
            $promptLabel = if ($favType -eq "Cloud") { "Cloud URL" } elseif ($favType -eq "Local") { "Local or UNC path" } else { "URL or Path" }
            $favValue = Read-Host "  $promptLabel"
            Write-Log "[USER INPUT] Favourites Add Value: '$favValue'" "INFO" -FileOnly
            if ([string]::IsNullOrWhiteSpace($favValue)) { Write-Host "  Value cannot be empty." -ForegroundColor Red; continue }

            # Derive a suggested name from the URL or path
            $suggestedName = try {
                if ($favValue -match "(?i)^https?://") {
                    # Cloud URL — take the last non-empty decoded path segment
                    $decoded = [uri]::UnescapeDataString($favValue -replace '\?.*$','')
                    $segs = ($decoded -split '/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    if ($segs.Count -gt 0) { $segs[-1] } else { "Favourite" }
                } else {
                    # Local / UNC path — take the leaf folder/file name
                    Split-Path $favValue.TrimEnd('\/') -Leaf
                }
            } catch { "Favourite" }
            $suggestedName = Get-SafeWindowsName $suggestedName

            Write-Host "  Suggested name: '$suggestedName'" -ForegroundColor DarkGray
            $favName = Read-Host "  Friendly name  (press ENTER to use '$suggestedName')"
            Write-Log "[USER INPUT] Favourites Add Name: '$favName'" "INFO" -FileOnly
            if ([string]::IsNullOrWhiteSpace($favName)) { $favName = $suggestedName }
            if ([string]::IsNullOrWhiteSpace($favName)) { Write-Host "  Name cannot be empty." -ForegroundColor Red; continue }

            $updatedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($f in $favs) { $updatedList.Add($f) }
            $updatedList.Add([PSCustomObject]@{
                Name  = $favName.Trim()
                Type  = $favType
                Value = $favValue.Trim()
                Added = (Get-Date -Format "yyyy-MM-dd")
            })
            Save-Favorites -Favorites $updatedList.ToArray()
            Write-Host "  ✔ Saved '$($favName.Trim())'." -ForegroundColor Green
            Write-Log "Favourites: Added '$($favName.Trim())' [$favType] => '$($favValue.Trim())'" "SUCCESS"
        }

        # --- RENAME ---
        elseif ($action -match '^[Rr]$') {
            if ($favs.Count -eq 0) { Write-Host "  No favourites to rename." -ForegroundColor Yellow; continue }
            $idxStr = $null
            while ($true) {
                $idxStr = Read-Host "  Number to rename (1-$($favs.Count)) or [Q] cancel"
                Write-Log "[USER INPUT] Favourites Rename Index: '$idxStr'" "INFO" -FileOnly
                if ($idxStr -match '^[Qq]$') { $idxStr = $null; break }
                if ($idxStr -match '^\d+$' -and [int]$idxStr -ge 1 -and [int]$idxStr -le $favs.Count) { break }
                Write-Host "  Please enter a number between 1 and $($favs.Count)." -ForegroundColor Red
            }
            if (-not $idxStr) { continue }
            $idx     = [int]$idxStr - 1
            $oldName = $favs[$idx].Name
            $newName = Read-Host "  New name for '$oldName'"
            Write-Log "[USER INPUT] Favourites New Name: '$newName'" "INFO" -FileOnly
            if ([string]::IsNullOrWhiteSpace($newName)) { Write-Host "  Name cannot be empty." -ForegroundColor Red; continue }

            $updatedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            for ($i = 0; $i -lt $favs.Count; $i++) {
                if ($i -eq $idx) {
                    $updatedList.Add([PSCustomObject]@{ Name = $newName.Trim(); Type = $favs[$idx].Type; Value = $favs[$idx].Value; Added = $favs[$idx].Added })
                } else { $updatedList.Add($favs[$i]) }
            }
            Save-Favorites -Favorites $updatedList.ToArray()
            Write-Host "  ✔ Renamed '$oldName' to '$($newName.Trim())'." -ForegroundColor Green
            Write-Log "Favourites: Renamed '$oldName' to '$($newName.Trim())'" "SUCCESS"
        }

        # --- EDIT VALUE (new in v86) ---
        elseif ($action -match '^[Ee]$') {
            if ($favs.Count -eq 0) { Write-Host "  No favourites to edit." -ForegroundColor Yellow; continue }
            $idxStr = $null
            while ($true) {
                $idxStr = Read-Host "  Number to edit (1-$($favs.Count)) or [Q] cancel"
                Write-Log "[USER INPUT] Favourites Edit Index: '$idxStr'" "INFO" -FileOnly
                if ($idxStr -match '^[Qq]$') { $idxStr = $null; break }
                if ($idxStr -match '^\d+$' -and [int]$idxStr -ge 1 -and [int]$idxStr -le $favs.Count) { break }
                Write-Host "  Please enter a number between 1 and $($favs.Count)." -ForegroundColor Red
            }
            if (-not $idxStr) { continue }
            $idx     = [int]$idxStr - 1
            $oldVal  = $favs[$idx].Value
            Write-Host "  Current value: $oldVal" -ForegroundColor DarkGray
            $newVal  = Read-Host "  New URL or path (or ENTER to keep current)"
            Write-Log "[USER INPUT] Favourites Edit New Value: '$newVal'" "INFO" -FileOnly
            if ([string]::IsNullOrWhiteSpace($newVal)) { Write-Host "  No change made." -ForegroundColor DarkGray; continue }

            $updatedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            for ($i = 0; $i -lt $favs.Count; $i++) {
                if ($i -eq $idx) {
                    $updatedList.Add([PSCustomObject]@{ Name = $favs[$idx].Name; Type = $favs[$idx].Type; Value = $newVal.Trim(); Added = $favs[$idx].Added })
                } else { $updatedList.Add($favs[$i]) }
            }
            Save-Favorites -Favorites $updatedList.ToArray()
            Write-Host "  ✔ Updated '$($favs[$idx].Name)'." -ForegroundColor Green
            Write-Log "Favourites: Edited value for '$($favs[$idx].Name)' => '$($newVal.Trim())'" "SUCCESS"
        }

        # --- DELETE ---
        elseif ($action -match '^[Dd]$') {
            if ($favs.Count -eq 0) { Write-Host "  No favourites to delete." -ForegroundColor Yellow; continue }
            $idxStr = $null
            while ($true) {
                $idxStr = Read-Host "  Number to delete (1-$($favs.Count)) or [Q] cancel"
                Write-Log "[USER INPUT] Favourites Delete Index: '$idxStr'" "INFO" -FileOnly
                if ($idxStr -match '^[Qq]$') { $idxStr = $null; break }
                if ($idxStr -match '^\d+$' -and [int]$idxStr -ge 1 -and [int]$idxStr -le $favs.Count) { break }
                Write-Host "  Please enter a number between 1 and $($favs.Count)." -ForegroundColor Red
            }
            if (-not $idxStr) { continue }
            $idx     = [int]$idxStr - 1
            $delName = $favs[$idx].Name
            $confirm = Read-Host "  Type 'YES' to delete '$delName'"
            Write-Log "[USER INPUT] Favourites Delete Confirm: '$confirm'" "INFO" -FileOnly
            if ($confirm -cne 'YES') { Write-Host "  Delete cancelled." -ForegroundColor Yellow; continue }

            $updatedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            for ($i = 0; $i -lt $favs.Count; $i++) { if ($i -ne $idx) { $updatedList.Add($favs[$i]) } }
            Save-Favorites -Favorites $updatedList.ToArray()
            Write-Host "  ✔ Deleted '$delName'." -ForegroundColor Green
            Write-Log "Favourites: Deleted '$delName'" "SUCCESS"
        }

        # --- QUIT ---
        elseif ($action -match '^[Qq]$') {
            Write-Log "Exiting Favourites Manager." "INFO"
            break
        } else {
            Write-Host "  Not recognised. Use A, R, E, D, or Q." -ForegroundColor Red
        }
    }
}

# --- FUNCTION: Read-PathInput ---
# PURPOSE: Smart wrapper around Read-Host for every URL/path prompt in the script.
#
# FEATURES:
#   [F]  Opens the Favourites Picker (filtered by Cloud/Local/Any).
#   [R]  Re-uses the last value entered for this same prompt type in the current session.
#        History is stored in $script:PathHistory — in-memory only, never written to disk.
#        This means no sensitive paths are persisted between sessions.
#   [Q]  Cancels the current operation and returns the caller to the previous menu.
#   ENTER on an optional field returns "" when -AllowEmpty is set.
#   All inputs are logged automatically to the audit log.
#
function Read-PathInput {
    param(
        [string]$Prompt,
        [string]$Type       = "Any",    # "Cloud", "Local", or "Any" — filters the picker
        [string]$LogLabel   = "",       # Optional label override for log entries
        [switch]$AllowEmpty             # Allow ENTER with no input as a valid "skip"
    )

    # Initialise the session history store if this is the first call in the session.
    if (-not $script:PathHistory) { $script:PathHistory = @{} }

    $favs     = @(Get-Favorites)
    $hasFavs  = $favs.Count -gt 0
    # History key: combine Type + LogLabel so "Source URL (Cloud)" and "Dest Path (Local)" are separate.
    $histKey  = "$Type|$LogLabel"
    $hasHist  = $script:PathHistory.ContainsKey($histKey)
    $lastVal  = if ($hasHist) { $script:PathHistory[$histKey] } else { "" }

    # Build the hint line dynamically based on what is available.
    $hints = [System.Collections.Generic.List[string]]::new()
    $hints.Add("[F] Favourites")   # always shown so users discover the feature
    if ($hasHist)  { $hints.Add("[R] Repeat last") }
    $hints.Add("[Q] Cancel")
    $hintLine = $hints -join "  |  "

    $label = if ($LogLabel) { $LogLabel } else { $Prompt }

    while ($true) {
        Write-Host ""
        # Show the last-used value as a hint below the prompt so users know what [R] will give them.
        if ($hasHist) {
            Write-Host "  Last used: $lastVal" -ForegroundColor DarkGray
        }
        $val = Read-Host "$Prompt`n  $hintLine"
        Write-Log "[USER INPUT] $label`: '$val'" "INFO" -FileOnly

        # [Q] — cancel
        if ($val -match "^[Qq]$") { return "Q" }

        # [F] — open Favourites Picker
        if ($val -match "^[Ff]$") {
            if (-not $hasFavs) {
                Write-Host "  No Favourites saved yet. Use [F] from the Main Menu to add some." -ForegroundColor Yellow
                continue
            }
            $picked = Show-FavoritePicker -Type $Type
            if ($picked) {
                Write-Log "[USER INPUT] $label (via Favourites): '$picked'" "INFO" -FileOnly
                $script:PathHistory[$histKey] = $picked   # also save to session history
                return $picked
            }
            continue
        }

        # [R] — repeat last used value for this prompt type
        if ($val -match "^[Rr]$") {
            if ($hasHist) {
                Write-Host "  Using: $lastVal" -ForegroundColor Green
                Write-Log "[USER INPUT] $label (via History): '$lastVal'" "INFO" -FileOnly
                return $lastVal
            }
            Write-Host "  No history yet for this prompt in this session. Please type a value." -ForegroundColor Yellow
            continue
        }

        # Empty input
        if ([string]::IsNullOrWhiteSpace($val)) {
            if ($AllowEmpty) { return "" }
            Write-Host "  Input cannot be empty. Enter a value, or use [F]/[R]/[Q]." -ForegroundColor Yellow
            continue
        }

        # Valid input — store in session history before returning
        $script:PathHistory[$histKey] = $val
        # Re-read history flag for future loops (in case user loops back via [F] cancel etc.)
        $hasHist = $true; $lastVal = $val
        return $val
    }
}

# ========================================================================================
# END OF v83 FAVOURITES SYSTEM — All functions below are unchanged from v79/v80
# ========================================================================================

# --- FUNCTION: Ensure-PnPModule ---
function Ensure-PnPModule {
    $pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version.ToString() -eq "1.12.0" }

    if (-not $pnpModule) {
        Write-Log "Installing PnP.PowerShell v1.12.0 for Cloud features..." "WARN"
        try {
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
            }
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Uninstall-Module -Name PnP.PowerShell -AllVersions -Force -ErrorAction SilentlyContinue
            Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Log "CRITICAL ERROR: Failed to install PnP.PowerShell." "ERROR"
            throw "Script halted due to installation failure."
        }
    }
    try {
        Import-Module PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction Stop
    } catch {
        Write-Log "CRITICAL ERROR: Failed to import PnP.PowerShell module." "ERROR"
        throw "Script halted due to import failure."
    }
}

Ensure-PnPModule

# --- FUNCTION: Clear-KeyBuffer ---
function Clear-KeyBuffer {
    try {
        while ([console]::KeyAvailable) { $null = [console]::ReadKey($true) }
    } catch {}
}

# --- FUNCTION: Get-LongPath ---
function Get-LongPath {
    param([string]$Path)
    $Path = $Path -replace '/', '\'
    if ($Path.Length -lt 240 -and -not [System.IO.Path]::IsPathRooted($Path)) {
        $Path = [System.IO.Path]::GetFullPath($Path)
    }
    if ($Path.Length -ge 240 -and -not $Path.StartsWith("\\?\")) {
        if ($Path.StartsWith("\\")) { return "\\?\UNC\" + $Path.Substring(2) }
        return "\\?\$Path"
    }
    return $Path
}

# --- FUNCTION: Get-SafeWindowsName ---
function Get-SafeWindowsName {
    param([string]$Name)
    $Name = $Name -replace '[<>:"/\\|?*\x00-\x1F]', '-'
    $Name = $Name.TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($Name)) { return "Unnamed_Item_$([guid]::NewGuid().ToString().Substring(0,8))" }
    return $Name
}

# --- FUNCTION: Parse-CloudUrl ---
function Parse-CloudUrl {
    param([string]$Url)
    try {
        if ($Url -notmatch "(?i)^https?://") {
            Write-Log "Invalid Cloud URL. Ensure you paste a valid web link starting with http:// or https://" "ERROR"
            return $null
        }
        $Uri      = [System.Uri]$Url
        $HostName = "https://$($Uri.Host)"
        $DecodedUrl = [uri]::UnescapeDataString($Url)

        if ($DecodedUrl -match "id=([^&]+)") { $ServerRelativeUrl = $matches[1] }
        else {
            $CleanUrl = $Url -replace "/:[a-zA-Z]:/[a-zA-Z]/", "/"
            if ($CleanUrl -match "^([^?]+)") { $CleanUrl = $matches[1] }
            $CleanUri = [System.Uri]$CleanUrl
            $ServerRelativeUrl = $CleanUri.AbsolutePath
        }
        $ServerRelativeUrl = [uri]::UnescapeDataString($ServerRelativeUrl)

        if ($ServerRelativeUrl -match "(?i)^(/sites/[^/]+|/personal/[^/]+)") {
            $SiteUrl = $HostName + $matches[1]
        } else {
            $SiteUrl = $HostName
        }
        return @{ SiteUrl = $SiteUrl; FolderPath = $ServerRelativeUrl }
    } catch {
        Write-Log "Failed to parse URL. Reason: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# --- FUNCTION: Check-Interrupt ---
function Check-Interrupt {
    param(
        [string]$CurrentFile   = "N/A",
        [string]$CurrentFolder = "N/A"
    )
    try {
        if ([console]::KeyAvailable) {
            $key = [console]::ReadKey($true)
            if ($key.Key -eq 'P' -or $key.Key -eq 'p') {
                Write-Host ""
                Write-Host "=================================================" -ForegroundColor Magenta
                Write-Host "   JOB PAUSED" -ForegroundColor Magenta
                Write-Host "   Last Active Folder: $CurrentFolder" -ForegroundColor Magenta
                Write-Host "   Last Active File: $CurrentFile" -ForegroundColor Magenta
                Write-Host "=================================================" -ForegroundColor Magenta
                Write-Log "JOB PAUSED. Last active folder: [$CurrentFolder] | Last active file: [$CurrentFile]" "WARN"
                Write-Host "Press 'C' to Cancel the rest of the job, or ANY OTHER KEY to resume..." -ForegroundColor Yellow
                $resumeKey = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                if ($resumeKey.Character -match "^[Cc]$") {
                    Write-Log "User triggered CANCEL command. Stopped at folder: [$CurrentFolder] | file: [$CurrentFile]" "WARN"
                    $global:JobCancelled = $true
                    return "CANCEL"
                }
                Write-Log "User RESUMED job. Continuing at folder: [$CurrentFolder] | file: [$CurrentFile]" "INFO"
                Write-Host "--- RESUMING JOB ---" -ForegroundColor Magenta
                return "RESUME"
            }
            elseif ($key.Key -eq 'C' -or $key.Key -eq 'c') {
                Write-Host ""
                Write-Log "User triggered CANCEL command. Stopped at folder: [$CurrentFolder] | file: [$CurrentFile]" "WARN"
                $global:JobCancelled = $true
                return "CANCEL"
            }
        }
    } catch {}
    return "CONTINUE"
}

# --- FUNCTION: Invoke-ForceDelete ---
function Invoke-ForceDelete {
    param([string]$Path, [switch]$IsFolder)
    try {
        if ($IsFolder) { [System.IO.Directory]::Delete($Path, $true) }
        else           { [System.IO.File]::Delete($Path) }
        return $true
    } catch {
        try {
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            if ($IsFolder) {
                $dirInfo = [System.IO.DirectoryInfo]::new($Path)
                $dirInfo.Attributes = 'Normal'
                [System.IO.Directory]::Delete($Path, $true)
            } else {
                $fileInfo = [System.IO.FileInfo]::new($Path)
                $fileInfo.Attributes = 'Normal'
                [System.IO.File]::Delete($Path)
            }
            return $true
        } catch { return $false }
    }
}

# --- FUNCTION: Remove-EmptyDirectories ---
function Remove-EmptyDirectories {
    param([string]$TargetFolder)
    $Dirs  = @(Get-ChildItem -LiteralPath $TargetFolder -Recurse -Directory -Force | Sort-Object FullName -Descending)
    $Dirs += Get-Item -LiteralPath $TargetFolder -Force
    foreach ($dir in $Dirs) {
        if (-not [System.IO.Directory]::Exists($dir.FullName)) { continue }
        $Contents = Get-ChildItem -LiteralPath $dir.FullName -Force
        if ($Contents.Count -eq 0) {
            $LongDir = Get-LongPath $dir.FullName
            if (Invoke-ForceDelete -Path $LongDir -IsFolder) {
                Write-Log "Deleted Empty Folder Skeleton: $($dir.Name)" "SUCCESS" -ConsoleOnly
            }
        }
    }
}

# --- FUNCTION: Get-LocalDirectoryMetrics ---
function Get-LocalDirectoryMetrics {
    param([string]$Path)
    $LongPath = Get-LongPath $Path
    if (-not [System.IO.Directory]::Exists($LongPath)) { return $null }
    $size = 0; $files = 0; $folders = 0
    try {
        $dirInfo = [System.IO.DirectoryInfo]::new($LongPath)
        foreach ($file in $dirInfo.EnumerateFiles("*", [System.IO.SearchOption]::AllDirectories)) {
            $size += $file.Length
            $files++
        }
        foreach ($dir in $dirInfo.EnumerateDirectories("*", [System.IO.SearchOption]::AllDirectories)) {
            $folders++
        }
    } catch {
        $allItems    = Get-ChildItem -LiteralPath $LongPath -Recurse -Force -ErrorAction SilentlyContinue
        $filesArray  = @($allItems | Where-Object { -not $_.PSIsContainer })
        $foldersArray= @($allItems | Where-Object {  $_.PSIsContainer })
        $files   = $filesArray.Count
        $folders = $foldersArray.Count
        $size    = ($filesArray | Measure-Object -Property Length -Sum).Sum
    }
    $sizeLabel = Format-FileSize $size
    return @{ Size = $sizeLabel; FileCount = $files; FolderCount = $folders; RawBytes = $size }
}

# --- FUNCTION: Show-DestinationFreeSpace ---
# Uses Get-PathFreeSpace which works for both local (C:\) and UNC (\\server\share) paths.
function Show-DestinationFreeSpace {
    param([string]$DestPath)
    try {
        $freeBytes = Get-PathFreeSpace -Path $DestPath
        $displayPath = if ($DestPath.StartsWith("\\")) {
            # For UNC, show the share root (\\server\share) rather than the full sub-path
            $parts = $DestPath.TrimStart('\').Split('\')
            if ($parts.Count -ge 2) { "\\$($parts[0])\$($parts[1])" } else { $DestPath }
        } else {
            [System.IO.Path]::GetPathRoot((Get-LongPath $DestPath))
        }

        if ($freeBytes -ge 0) {
            Write-Log "Destination Free Space ($displayPath): $(Format-FileSize $freeBytes)" "INFO"
        } else {
            Write-Log "Destination Free Space: Could not be determined for '$displayPath'." "INFO"
        }
    } catch { }
}

# --- FUNCTION: Invoke-PreflightStorageCheck ---
# EXPLANATION: Gatekeeper for all local and download operations.
# Uses Get-PathFreeSpace which calls GetDiskFreeSpaceExW — works for LOCAL and UNC paths.
# The old DriveInfo approach only worked for local drive letters and caused a disruptive
# "proceed Y/N?" prompt whenever the destination was a UNC path. That is now fixed.
function Invoke-PreflightStorageCheck {
    param(
        [string]$DestRoot,
        [long]$TotalSourceBytes,
        [hashtable[]]$BatchItems
    )

    # Determine a display label for logs (\\server\share or C:\)
    $isUNC = $DestRoot.StartsWith("\\")
    $displayRoot = if ($isUNC) {
        $parts = $DestRoot.TrimStart('\').Split('\')
        if ($parts.Count -ge 2) { "\\$($parts[0])\$($parts[1])" } else { $DestRoot }
    } else {
        [System.IO.Path]::GetPathRoot((Get-LongPath $DestRoot))
    }

    [long]$FreeSpace  = Get-PathFreeSpace -Path $DestRoot
    $CanDetermine     = ($FreeSpace -ge 0)

    if (-not $CanDetermine) {
        # Space genuinely cannot be determined (very rare — unusual OS/network config).
        # For UNC paths this should not happen now that we use GetDiskFreeSpaceExW,
        # but we handle it gracefully just in case.
        Write-Host ""
        Write-Log "WARNING: Unable to determine free space on destination ($displayRoot)." "WARN"
        Write-Log "This can happen on some network shares with restricted permissions." "WARN"
        $proceed = Read-Host "Do you want to proceed with the operation anyway? (Y/N)"
        Write-Log "[USER INPUT] Storage Warning Override (Unknown Space): '$proceed'" "INFO" -FileOnly
        if ($proceed -notmatch "^[Yy]") { return $false }
        return $true
    }

    $FreeStr = Format-FileSize $FreeSpace
    $ReqStr  = Format-FileSize $TotalSourceBytes

    Write-Host ""
    Write-Log "--- PRE-FLIGHT STORAGE DIAGNOSTICS ---" "INFO"
    Write-Log "Destination ($displayRoot) Free Space: $FreeStr" "INFO"
    Write-Log "Total Source Payload:                  $ReqStr" "INFO"

    if ($TotalSourceBytes -le $FreeSpace) {
        Write-Log "Storage Check Passed. Sufficient space available." "SUCCESS"
        return $true
    }

    Write-Log "Standard payload exceeds free space. Initiating Smart Overwrite Calculation..." "WARN"
    [long]$NetBytes = 0

    foreach ($item in $BatchItems) {
        $src    = $item.Src
        $dst    = $item.Dst
        $srcLong = Get-LongPath $src
        if ([System.IO.File]::Exists($srcLong)) {
            $dLong = Get-LongPath $dst
            if ([System.IO.File]::Exists($dLong)) {
                $sInfo = [System.IO.FileInfo]::new($srcLong)
                $dInfo = [System.IO.FileInfo]::new($dLong)
                $diff  = $sInfo.Length - $dInfo.Length
                if ($diff -gt 0) { $NetBytes += $diff }
            } else {
                $sInfo = [System.IO.FileInfo]::new($srcLong)
                $NetBytes += $sInfo.Length
            }
        } elseif ([System.IO.Directory]::Exists($srcLong)) {
            $SrcFiles = Get-ChildItem -LiteralPath $srcLong -Recurse -File -Force -ErrorAction SilentlyContinue
            if ($SrcFiles) {
                foreach ($f in $SrcFiles) {
                    $rel    = $f.FullName.Substring($srcLong.Length).TrimStart('\')
                    $dPath  = Join-Path $dst $rel
                    $dLong  = Get-LongPath $dPath
                    if ([System.IO.File]::Exists($dLong)) {
                        $dInfo = [System.IO.FileInfo]::new($dLong)
                        $diff  = $f.Length - $dInfo.Length
                        if ($diff -gt 0) { $NetBytes += $diff }
                    } else { $NetBytes += $f.Length }
                }
            }
        }
    }

    $NetStr = Format-FileSize $NetBytes
    Write-Log "Smart Calculation Complete. Net Required Payload: $NetStr" "INFO"

    if ($NetBytes -le $FreeSpace) {
        Write-Log "Storage Check Passed via Overwrite logic. Proceeding..." "SUCCESS"
        return $true
    }

    Write-Host ""
    Write-Host "!!! CRITICAL STORAGE WARNING !!!" -ForegroundColor Red
    Write-Host "Destination does not have enough free space, even after calculating file overwrites." -ForegroundColor Red
    Write-Host "Net Required: $NetStr | Available: $FreeStr on $displayRoot" -ForegroundColor Red

    $proceed = Read-Host "Do you want to ignore this warning and force the operation to start? (Y/N)"
    Write-Log "[USER INPUT] Critical Storage Warning Override: '$proceed'" "INFO" -FileOnly
    if ($proceed -match "^[Yy]") {
        Write-Log "User elected to force the operation despite storage warnings." "WARN"
        return $true
    }
    Write-Log "Operation cancelled by user due to insufficient storage." "WARN"
    return $false
}

# --- FUNCTION: Perform-LocalSearch ---
function Perform-LocalSearch {
    param([string]$SearchPath, [string]$SearchTerm)
    Write-Log "--- STARTING LOCAL FILE SEARCH ---" "INFO"
    Write-Log "SEARCH PATH: $SearchPath"
    Write-Log "SEARCH TERM: $SearchTerm"

    $LongPath = Get-LongPath $SearchPath
    if (-not [System.IO.Directory]::Exists($LongPath)) {
        Write-Log "Search path does not exist!" "ERROR"
        return
    }

    $global:JobCancelled = $false
    Clear-KeyBuffer
    $sw          = [System.Diagnostics.Stopwatch]::StartNew()
    $FoundItems  = [System.Collections.Generic.List[hashtable]]::new()
    $ScannedFolders = 0
    $ScannedFiles   = 0

    Write-Host "Scanning directory... (Press 'C' to Cancel at any time)" -ForegroundColor DarkGray

    $Queue = [System.Collections.Generic.Queue[string]]::new()
    $Queue.Enqueue($LongPath)

    $SearchPattern = "*$SearchTerm*"
    if ($SearchTerm -match "\*") { $SearchPattern = $SearchTerm }

    $jobStartTime = Get-Date
    $startTimeStr = $jobStartTime.ToString("HH:mm:ss")
    $uiStopwatch  = [System.Diagnostics.Stopwatch]::StartNew()

    while ($Queue.Count -gt 0) {
        if ($global:JobCancelled) { break }
        $currentDir = $Queue.Dequeue()

        Check-Interrupt -CurrentFolder $currentDir -CurrentFile "Scanning..." | Out-Null
        if ($global:JobCancelled) { break }

        $ScannedFolders++

        if ($uiStopwatch.ElapsedMilliseconds -gt 500) {
            $elapsed    = (Get-Date) - $jobStartTime
            $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
            Write-Progress -Activity "Local Search Engine [Started: $startTimeStr]" -Status "Found: $($FoundItems.Count) items" -CurrentOperation "Elapsed Time: $elapsedStr | Scanning: $currentDir | HOTKEYS: Press 'P' to Pause | 'C' to Cancel"
            $uiStopwatch.Restart()
        }

        try {
            $dirInfo = [System.IO.DirectoryInfo]::new($currentDir)

            foreach ($file in $dirInfo.EnumerateFiles($SearchPattern, [System.IO.SearchOption]::TopDirectoryOnly)) {
                $ScannedFiles++
                $FoundItems.Add(@{
                    Name       = $file.Name
                    Path       = $file.FullName
                    Type       = "FILE"
                    Size       = $file.Length
                    Modified   = $file.LastWriteTime
                    ParentPath = $dirInfo.FullName
                })
                Write-Log "MATCH FOUND: [FILE] $($file.FullName) [Size: $(Format-FileSize $file.Length)]" "SUCCESS" -ConsoleOnly
            }

            foreach ($subDir in $dirInfo.EnumerateDirectories("*", [System.IO.SearchOption]::TopDirectoryOnly)) {
                if ($subDir.Name -like $SearchPattern) {
                    $FoundItems.Add(@{
                        Name       = $subDir.Name
                        Path       = $subDir.FullName
                        Type       = "FOLDER"
                        Size       = 0
                        Modified   = $subDir.LastWriteTime
                        ParentPath = $dirInfo.FullName
                    })
                    Write-Log "MATCH FOUND: [FOLDER] $($subDir.FullName)" "SUCCESS" -ConsoleOnly
                }
                $Queue.Enqueue($subDir.FullName)
            }
        } catch { }
    }

    Write-Progress -Activity "Local Search Engine [Started: $startTimeStr]" -Completed
    $sw.Stop()

    Write-Host ""
    if ($global:JobCancelled) { Write-Log "--- SEARCH ABORTED BY USER ---" "WARN" }
    else                      { Write-Log "--- SEARCH COMPLETE ---" "INFO" }

    Write-Log "Total Folders Scanned: $ScannedFolders" "INFO"
    Write-Log "Total Files Scanned:   $ScannedFiles" "INFO"
    Write-Log "Total Items Matched:   $($FoundItems.Count)" "SUCCESS"
    Write-Log "Total Time Elapsed:    $([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s." "INFO"

    if ($FoundItems.Count -gt 0) {
        $ts         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogPayload = "[$ts] [INFO] --- SEARCH RESULTS FOR: '$SearchTerm' ---`r`n"
        foreach ($f in $FoundItems) {
            $sz = if ($f.Type -eq "FILE") { " [Size: $(Format-FileSize $f.Size)]" } else { "" }
            $LogPayload += "[$ts] [MATCH] [VERIFIED: $($f.Type)] $($f.Path)$sz [Modified: $($f.Modified.ToString('yyyy-MM-dd HH:mm:ss'))]`r`n"
        }
        [System.IO.File]::AppendAllText($LogFile, $LogPayload)
        Write-Host "Full search results have been safely permanently stored in the log file." -ForegroundColor DarkGray

        # Calculate sizes for any matched folders before displaying results.
        # This runs once here so the list always shows folder sizes, not just on detail view.
        $folderCount = ($FoundItems | Where-Object { $_.Type -eq "FOLDER" }).Count
        if ($folderCount -gt 0) {
            Write-Host "  Calculating size of $folderCount matched folder(s)..." -ForegroundColor DarkGray
            for ($fi = 0; $fi -lt $FoundItems.Count; $fi++) {
                if ($FoundItems[$fi].Type -eq "FOLDER") {
                    $fm = Get-LocalDirectoryMetrics -Path $FoundItems[$fi].Path
                    if ($fm) {
                        # Store the calculated size (as bytes) back into the hashtable entry
                        # so the results list can display it just like a file size.
                        $FoundItems[$fi].Size     = $fm.RawBytes
                        $FoundItems[$fi].FolderInfo = "$($fm.FileCount) files, $($fm.FolderCount) sub-folders"
                    }
                }
            }
        }

        while ($true) {
            Write-Host "`n--- LOCAL SEARCH RESULTS ($($FoundItems.Count) items) ---" -ForegroundColor Cyan
            $numWidth = ([string]$FoundItems.Count).Length + 2
            for ($i = 0; $i -lt $FoundItems.Count; $i++) {
                $f        = $FoundItems[$i]
                $typeTag  = if ($f.Type -eq "FOLDER") { "[FOLDER]" } else { "[FILE]  " }
                $numLabel = "[$($i+1)]".PadLeft($numWidth)
                # Build the metadata tag at the end: [Size: x, Date Modified: yyyy-MM-dd]
                $metaParts = [System.Collections.Generic.List[string]]::new()
                if ($f.Size -gt 0) { $metaParts.Add("Size: $(Format-FileSize ([long]$f.Size))") }
                if ($f.Modified) { $metaParts.Add("Date Modified: $($f.Modified.ToString('yyyy-MM-dd'))") }
                $metaTag  = if ($metaParts.Count -gt 0) { "  [" + ($metaParts -join ", ") + "]" } else { "" }
                # Single line: index  type  full-path  [metadata]
                Write-Host " $numLabel $typeTag $($f.Path)" -NoNewline
                Write-Host $metaTag -ForegroundColor DarkGray
            }
            Write-Host "  ──────────────────────────────────────────────────────────" -ForegroundColor Cyan
            $opt = Read-Host "  Number = open details  |  [E] Export CSV  |  [Q] or ENTER = exit"

            if ([string]::IsNullOrWhiteSpace($opt) -or $opt -match '^[Qq]$') { break }

            if ($opt -match '^[Ee]$') {
                $csvPath = Join-Path $LogFileDirectory "LocalSearch_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
                $csvRows = $FoundItems | ForEach-Object {
                    $h = $_
                    [PSCustomObject]@{
                        Type       = "$($h.Type)"
                        Name       = "$($h.Name)"
                        Path       = "$($h.Path)"
                        ParentPath = "$($h.ParentPath)"
                        Size       = if ($h.Size) { Format-FileSize ([long]$h.Size) } else { "—" }
                        Modified   = if ($h.Modified) { $h.Modified.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                    }
                }
                $csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                Write-Host ""
                Write-Host "  ✔ Exported $($FoundItems.Count) result(s) to:" -ForegroundColor Green
                Write-Host "    $csvPath" -ForegroundColor Green
                $openCsv = Read-Host "  [O] Open the CSV file now, or ENTER to continue"
                if ($openCsv -match '^[Oo]$') { try { Invoke-Item $csvPath } catch { Write-Host "  Could not open file." -ForegroundColor Red } }
            }
            elseif ($opt -match '^\d+$' -and [int]$opt -gt 0 -and [int]$opt -le $FoundItems.Count) {
                $selItem = $FoundItems[[int]$opt - 1]
                $isFolder = ($selItem.Type -eq "FOLDER")
                $itemKind = if ($isFolder) { "FOLDER" } else { "FILE" }
                Write-Host ""
                Write-Host "  ┌─ ${itemKind}: $($selItem.Name) ────────────────────────────────" -ForegroundColor Magenta
                Write-Host "  │  Path          : $($selItem.Path)" -ForegroundColor Green
                Write-Host "  │  Parent Folder : $($selItem.ParentPath)" -ForegroundColor Yellow
                if ($isFolder) {
                    # Calculate folder size on demand — may take a moment for large folders
                    Write-Host "  │  Calculating folder size..." -NoNewline -ForegroundColor DarkGray
                    $fMetrics = Get-LocalDirectoryMetrics -Path $selItem.Path
                    if ($fMetrics) {
                        Write-Host "`r  │  Contents      : $($fMetrics.Size)  ($($fMetrics.FileCount) files, $($fMetrics.FolderCount) sub-folders)   " -ForegroundColor DarkGray
                    } else {
                        Write-Host "`r  │  Contents      : Could not be calculated (folder may be locked)    " -ForegroundColor DarkGray
                    }
                } elseif ($selItem.Size -gt 0) {
                    Write-Host "  │  Size          : $(Format-FileSize ([long]$selItem.Size))" -ForegroundColor DarkGray
                } else {
                    Write-Host "  │  Size          : —" -ForegroundColor DarkGray
                }
                if ($selItem.Modified) { Write-Host "  │  Modified      : $($selItem.Modified.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray }
                Write-Host "  └────────────────────────────────────────────────────────────" -ForegroundColor Magenta
                # Inner loop — stay in detail view
                while ($true) {
                    $openLabel = if ($isFolder) { "[I] Open folder in Explorer  " } else { "[I] Open file in Explorer    " }
                    $openOpt = Read-Host "  [O] Open parent folder  |  $openLabel|  [Q] Back to results"
                    if ($openOpt -match '^[Oo]$') {
                        try { Invoke-Item $selItem.ParentPath; Write-Host "  Opened parent folder." -ForegroundColor DarkGray }
                        catch { Write-Host "  Could not open location." -ForegroundColor Red }
                    } elseif ($openOpt -match '^[Ii]$') {
                        try { Invoke-Item $selItem.Path; Write-Host "  Opened $($itemKind.ToLower())." -ForegroundColor DarkGray }
                        catch { Write-Host "  Could not open item." -ForegroundColor Red }
                    } elseif ($openOpt -match '^[Qq]$' -or [string]::IsNullOrWhiteSpace($openOpt)) { break }
                }
            } else {
                Write-Host "  Invalid selection — type a number, E, Q, or ENTER." -ForegroundColor Red
            }
        }
    }
}

# --- FUNCTION: Perform-CloudSearch ---
function Perform-CloudSearch {
    param([hashtable]$CloudData, [string]$SearchTerm, [switch]$SkipAuth)
    Write-Log "--- STARTING CLOUD FILE SEARCH ---" "INFO"
    Write-Log "SEARCH PATH: $($CloudData.FolderPath)"
    Write-Log "SEARCH TERM: $SearchTerm"

    $global:JobCancelled = $false
    Clear-KeyBuffer
    $sw          = [System.Diagnostics.Stopwatch]::StartNew()
    $FoundItems  = [System.Collections.Generic.List[hashtable]]::new()
    $ScannedFolders = 0
    $ScannedFiles   = 0

    try {
        if (-not $SkipAuth) {
            Write-Log "Authenticating to Site: $($CloudData.SiteUrl)..." "WARN"
            $originalWarning = $WarningPreference
            $WarningPreference = "SilentlyContinue"
            Connect-PnPOnline -Url $CloudData.SiteUrl -UseWebLogin -ErrorAction Stop
            $WarningPreference = $originalWarning
        }

        $HostAuthority = [uri]::new($CloudData.SiteUrl).GetLeftPart('Authority')
        $SitePath      = [uri]::new($CloudData.SiteUrl).AbsolutePath
        $BaseRelUrl    = $CloudData.FolderPath.Substring($SitePath.Length).TrimStart('/')

        $Queue = [System.Collections.Generic.Queue[string]]::new()
        $Queue.Enqueue($BaseRelUrl)

        $SearchPattern = "*$SearchTerm*"
        if ($SearchTerm -match "\*") { $SearchPattern = $SearchTerm }

        Write-Host "Scanning cloud directory... (Press 'C' to Cancel at any time)" -ForegroundColor DarkGray

        $jobStartTime = Get-Date
        $startTimeStr = $jobStartTime.ToString("HH:mm:ss")
        $uiStopwatch  = [System.Diagnostics.Stopwatch]::StartNew()

        while ($Queue.Count -gt 0) {
            if ($global:JobCancelled) { break }
            $currentCloudDir = $Queue.Dequeue()

            Check-Interrupt -CurrentFolder $currentCloudDir -CurrentFile "Scanning..." | Out-Null
            if ($global:JobCancelled) { break }

            $ScannedFolders++

            if ($uiStopwatch.ElapsedMilliseconds -gt 500) {
                $elapsed    = (Get-Date) - $jobStartTime
                $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                Write-Progress -Activity "Cloud Search Engine [Started: $startTimeStr]" -Status "Found: $($FoundItems.Count) items" -CurrentOperation "Elapsed Time: $elapsedStr | Scanning: /$currentCloudDir | HOTKEYS: Press 'P' to Pause | 'C' to Cancel"
                $uiStopwatch.Restart()
            }

            Start-Sleep -Milliseconds 100

            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $currentCloudDir -ErrorAction SilentlyContinue
            if ($items) {
                foreach ($item in $items) {
                    if ($item.GetType().Name -eq "Folder") {
                        if ($item.Name -like $SearchPattern) {
                            $encRel       = $item.ServerRelativeUrl -replace ' ', '%20'
                            $encParentRel = ($SitePath + "/$currentCloudDir") -replace ' ', '%20'
                            $FoundItems.Add(@{
                                Name      = $item.Name
                                Path      = $item.ServerRelativeUrl
                                Type      = "FOLDER"
                                Size      = 0
                                Modified  = $item.TimeLastModified
                                Url       = $HostAuthority + $encRel
                                ParentUrl = $HostAuthority + $encParentRel
                            })
                            Write-Log "MATCH FOUND: [FOLDER] $($item.ServerRelativeUrl)" "SUCCESS" -ConsoleOnly
                        }
                        $Queue.Enqueue("$currentCloudDir/$($item.Name)")
                    } else {
                        $ScannedFiles++
                        if ($item.Name -like $SearchPattern) {
                            $encRel       = $item.ServerRelativeUrl -replace ' ', '%20'
                            $encParentRel = ($SitePath + "/$currentCloudDir") -replace ' ', '%20'
                            $FoundItems.Add(@{
                                Name      = $item.Name
                                Path      = $item.ServerRelativeUrl
                                Type      = "FILE"
                                Size      = $item.Length
                                Modified  = $item.TimeLastModified
                                Url       = ($HostAuthority + $encRel) + "?web=1"
                                ParentUrl = $HostAuthority + $encParentRel
                            })
                            Write-Log "MATCH FOUND: [FILE] $($item.ServerRelativeUrl) [Size: $(Format-FileSize $item.Length)]" "SUCCESS" -ConsoleOnly
                        }
                    }
                }
            }
        }
    } catch {
        Write-Log "Cloud Search Error: $($_.Exception.Message)" "ERROR"
    } finally {
        if (-not $SkipAuth) { try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {} }
    }

    Write-Progress -Activity "Cloud Search Engine [Started: $startTimeStr]" -Completed
    $sw.Stop()

    Write-Host ""
    if ($global:JobCancelled) { Write-Log "--- SEARCH ABORTED BY USER ---" "WARN" }
    else                      { Write-Log "--- SEARCH COMPLETE ---" "INFO" }

    Write-Log "Total Folders Scanned: $ScannedFolders" "INFO"
    Write-Log "Total Files Scanned:   $ScannedFiles" "INFO"
    Write-Log "Total Items Matched:   $($FoundItems.Count)" "SUCCESS"
    Write-Log "Total Time Elapsed:    $([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s." "INFO"

    if ($FoundItems.Count -gt 0) {
        $ts         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogPayload = "[$ts] [INFO] --- CLOUD SEARCH RESULTS FOR: '$SearchTerm' ---`r`n"
        foreach ($f in $FoundItems) {
            $sz = if ($f.Type -eq "FILE") { " [Size: $(Format-FileSize $f.Size)]" } else { "" }
            $LogPayload += "[$ts] [MATCH] [VERIFIED: $($f.Type)] $($f.Path)$sz`r`n"
        }
        [System.IO.File]::AppendAllText($LogFile, $LogPayload)
        Write-Host "Full search results have been safely permanently stored in the log file." -ForegroundColor DarkGray

        while ($true) {
            Write-Host "`n--- CLOUD SEARCH RESULTS ($($FoundItems.Count) items) ---" -ForegroundColor Cyan
            $numWidth = ([string]$FoundItems.Count).Length + 2
            for ($i = 0; $i -lt $FoundItems.Count; $i++) {
                $f        = $FoundItems[$i]
                $typeTag  = if ($f.Type -eq "FOLDER") { "[FOLDER]" } else { "[FILE]  " }
                $numLabel = "[$($i+1)]".PadLeft($numWidth)
                # Build the metadata tag at the end: [Size: x, Date Modified: yyyy-MM-dd]
                $metaParts = [System.Collections.Generic.List[string]]::new()
                if ($f.Type -eq "FILE" -and $f.Size) { $metaParts.Add("Size: $(Format-FileSize ([long]$f.Size))") }
                if ($f.Modified -and "$($f.Modified)" -ne "") {
                    try { $metaParts.Add("Date Modified: $(([datetime]$f.Modified).ToString('yyyy-MM-dd'))") } catch { }
                }
                $metaTag  = if ($metaParts.Count -gt 0) { "  [" + ($metaParts -join ", ") + "]" } else { "" }
                # Single line: index  type  full-path  [metadata]
                Write-Host " $numLabel $typeTag $($f.Path)" -NoNewline
                Write-Host $metaTag -ForegroundColor DarkGray
            }
            Write-Host "  ──────────────────────────────────────────────────────────" -ForegroundColor Cyan
            $opt = Read-Host "  Number = open details  |  [E] Export CSV  |  [Q] or ENTER = exit"

            if ([string]::IsNullOrWhiteSpace($opt) -or $opt -match '^[Qq]$') { break }

            if ($opt -match '^[Ee]$') {
                # Fix: convert hashtable list to PSCustomObjects for proper CSV serialisation
                $csvPath = Join-Path $LogFileDirectory "CloudSearch_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
                $csvRows = $FoundItems | ForEach-Object {
                    $h = $_
                    $modStr = if ($h.Modified -and "$($h.Modified)" -ne "") {
                        try { ([datetime]$h.Modified).ToString("yyyy-MM-dd HH:mm") } catch { "" }
                    } else { "" }
                    [PSCustomObject]@{
                        Type      = "$($h.Type)"
                        Name      = "$($h.Name)"
                        Path      = "$($h.Path)"
                        ParentUrl = "$($h.ParentUrl)"
                        Url       = "$($h.Url)"
                        Size      = if ($h.Size) { Format-FileSize ([long]$h.Size) } else { "—" }
                        Modified  = $modStr
                    }
                }
                $csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                Write-Host ""
                Write-Host "  ✔ Exported $($FoundItems.Count) result(s) to:" -ForegroundColor Green
                Write-Host "    $csvPath" -ForegroundColor Green
                $openCsv = Read-Host "  [O] Open the CSV file now, or ENTER to continue"
                if ($openCsv -match '^[Oo]$') { try { Invoke-Item $csvPath } catch { Write-Host "  Could not open file." -ForegroundColor Red } }
            }
            elseif ($opt -match '^\d+$' -and [int]$opt -gt 0 -and [int]$opt -le $FoundItems.Count) {
                $selItem = $FoundItems[[int]$opt - 1]
                $isFolder = ($selItem.Type -eq "FOLDER")
                $itemKind = if ($isFolder) { "FOLDER" } else { "FILE" }
                Write-Host ""
                Write-Host "  ┌─ ${itemKind}: $($selItem.Name) ────────────────────────────────" -ForegroundColor Magenta
                Write-Host "  │  $itemKind URL  : $($selItem.Url)" -ForegroundColor Green
                Write-Host "  │  Parent Folder : $($selItem.ParentUrl)" -ForegroundColor Yellow
                if (-not $isFolder -and $selItem.Size -gt 0) {
                    Write-Host "  │  Size          : $(Format-FileSize ([long]$selItem.Size))" -ForegroundColor DarkGray
                } elseif ($isFolder) {
                    Write-Host "  │  Size          : —  (cloud folder sizes are not fetched by this engine)" -ForegroundColor DarkGray
                } else {
                    Write-Host "  │  Size          : —" -ForegroundColor DarkGray
                }
                if ($selItem.Modified -and "$($selItem.Modified)" -ne "") {
                    try {
                        $modDt = [datetime]$selItem.Modified
                        Write-Host "  │  Modified      : $($modDt.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
                    } catch { }
                }
                Write-Host "  └────────────────────────────────────────────────────────────" -ForegroundColor Magenta
                # Inner loop — stay in detail view so user can re-open without scrolling results
                while ($true) {
                    $openLabel = if ($isFolder) { "[I] Open folder  " } else { "[I] Open file    " }
                    $openOpt = Read-Host "  [O] Open parent folder  |  $openLabel|  [Q] Back to results"
                    if ($openOpt -match '^[Oo]$') {
                        try { Start-Process $selItem.ParentUrl; Write-Host "  Opened parent folder in browser." -ForegroundColor DarkGray }
                        catch { Write-Host "  Could not open location." -ForegroundColor Red }
                    } elseif ($openOpt -match '^[Ii]$') {
                        try { Start-Process $selItem.Url; Write-Host "  Opened $($itemKind.ToLower()) in browser." -ForegroundColor DarkGray }
                        catch { Write-Host "  Could not open item." -ForegroundColor Red }
                    } elseif ($openOpt -match '^[Qq]$' -or [string]::IsNullOrWhiteSpace($openOpt)) { break }
                }
            } else {
                Write-Host "  Invalid selection — type a number, E, Q, or ENTER." -ForegroundColor Red
            }
        }
    }
}

# ========================================================================================
# ENGINE SUITE B: CLOUD API OPERATIONS
# ========================================================================================

# --- FUNCTION: Perform-CloudTransfer ---
# PURPOSE: Performs a SERVER-SIDE copy or move of a folder within SharePoint.
# "Server-side" means Microsoft's own servers do the work — your PC sends one API call
# and the data never travels through your internet connection or hard drive.
# This makes it near-instant for even very large folders.
#
# IMPORTANT LIMITATION: Both source and destination must be on the SAME SharePoint site
# (same /sites/SiteName or /personal/username prefix). Cross-site server-side transfers
# are not supported by Microsoft's API — use Download + Upload for those.
#
# AUTH: Connects using WebLogin (browser pop-up). One connection per call.
# THROTTLE: Move-PnPFile / Copy-PnPFile are wrapped with Invoke-PnPWithRetry.
function Perform-CloudTransfer {
    param([hashtable]$SrcData, [hashtable]$DstData, [switch]$IsMove)
    $ModeName = if ($IsMove) { "MOVE" } else { "COPY" }
    Write-Log "--- STARTING CLOUD $ModeName JOB ---"
    Write-Log "SOURCE FOLDER: $($SrcData.FolderPath)"
    Write-Log "DESTINATION FOLDER: $($DstData.FolderPath)"

    $global:JobCancelled = $false
    Clear-KeyBuffer

    try {
        Write-Log "Authenticating to Site: $($SrcData.SiteUrl)..." "WARN"
        $originalWarning = $WarningPreference
        $WarningPreference = "SilentlyContinue"
        Connect-PnPOnline -Url $SrcData.SiteUrl -UseWebLogin -ErrorAction Stop
        $WarningPreference = $originalWarning

        Write-Host ""
        Write-Host "============================================================" -ForegroundColor Magenta
        Write-Host " Microsoft Servers are processing your $ModeName request." -ForegroundColor Magenta
        Write-Host " Do not close this window. Please stand by..." -ForegroundColor Magenta
        Write-Host "============================================================" -ForegroundColor Magenta
        Write-Host ""

        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($IsMove) {
            Invoke-PnPWithRetry -OperationName "MOVE $($SrcData.FolderPath)" -ScriptBlock {
                Move-PnPFile -SourceUrl $SrcData.FolderPath -TargetUrl $DstData.FolderPath -Force -Overwrite -ErrorAction Stop
            } | Out-Null
        } else {
            Invoke-PnPWithRetry -OperationName "COPY $($SrcData.FolderPath)" -ScriptBlock {
                Copy-PnPFile -SourceUrl $SrcData.FolderPath -TargetUrl $DstData.FolderPath -Force -Overwrite -ErrorAction Stop
            } | Out-Null
        }

        $sw.Stop()
        Write-Log "SUCCESS! Transferred to '$($DstData.FolderPath)'" "SUCCESS"
        Write-Log "Total Time Elapsed: $([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s." "SUCCESS"
    } catch {
        Write-Log "$ModeName Failed. Reason: $($_.Exception.Message)" "ERROR"
    } finally {
        try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
    }
}

# --- FUNCTION: Perform-CloudBatch ---
function Perform-CloudBatch {
    param([hashtable]$SrcData, [hashtable]$DstData)
    Write-Log "--- ENTERING CLOUD BATCH MODE ---"
    Write-Log "SOURCE FOLDER: $($SrcData.FolderPath)"
    if ($DstData) { Write-Log "DESTINATION FOLDER: $($DstData.FolderPath)" }

    try {
        Write-Log "Authenticating to Site: $($SrcData.SiteUrl)..." "WARN"
        $originalWarning = $WarningPreference
        $WarningPreference = "SilentlyContinue"
        Connect-PnPOnline -Url $SrcData.SiteUrl -UseWebLogin -ErrorAction Stop
        $WarningPreference = $originalWarning

        $SitePath       = [uri]::new($SrcData.SiteUrl).AbsolutePath
        $SiteRelativeSrc = $SrcData.FolderPath.Substring($SitePath.Length).TrimStart('/')

        while ($true) {
            Write-PathBanner -Label1 "Source" -Path1 $SrcData.FolderPath -Label2 "Dest" -Path2 $(if($DstData){$DstData.FolderPath}else{""}) -BoxColor "Cyan"
            Write-Host "Fetching directory inventory from the cloud..." -ForegroundColor DarkGray
            $RawItems = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeSrc -ErrorAction SilentlyContinue

            if (-not $RawItems -or $RawItems.Count -eq 0) {
                Write-Log "Source directory is completely empty. Exiting Cloud Batch Mode..." "WARN"
                break
            }

            $Items = @($RawItems | Sort-Object @{Expression={$_.GetType().Name -eq "Folder"};Descending=$true}, @{Expression={$_.Name};Ascending=$true})

            Write-Host ""
            Write-Host "  ┌─ CLOUD DIRECTORY INDEX ──────────────────────────────" -ForegroundColor Cyan
            Write-Host "  │  # │ Type   │ Name                    │ Items │ Modified   " -ForegroundColor DarkGray
            Write-Host "  │  Items = top-level count shown by Microsoft API  │ Modified = last changed date" -ForegroundColor DarkGray

            $MenuAudit = [System.Collections.Generic.List[string]]::new()
            $MenuAudit.Add("--- CLOUD DIRECTORY INDEX ---")

            for ($i = 0; $i -lt $Items.Count; $i++) {
                $item = $Items[$i]
                $type = if ($item.GetType().Name -eq "Folder") { "[FOLDER]" } else { "[FILE]  " }
                if ($item.GetType().Name -eq "Folder") {
                    $modDate = if ($item.TimeLastModified) { $item.TimeLastModified.ToString("MM/dd/yyyy") } else { "Unknown" }
                    $iCount  = if ($null -ne $item.ItemCount) { $item.ItemCount } else { "0" }
                    $logStr  = "[$($i + 1)] $type $($item.Name) [Items: $iCount | Modified: $modDate]"
                } else {
                    $sizeStr = Format-FileSize $item.Length
                    $logStr  = "[$($i + 1)] $type $($item.Name) [Size: $sizeStr]"
                }
                Write-Host " $logStr"
                $MenuAudit.Add($logStr)
            }

            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $LogPayload = ""
            foreach ($line in $MenuAudit) { $LogPayload += "[$ts] [MENU] $line`r`n" }
            [System.IO.File]::AppendAllText($LogFile, $LogPayload)

            Write-Host ""
            Write-Host " [A] Process ALL items in this directory" -ForegroundColor Green
            Write-Host "  │  [S] Search this directory" -ForegroundColor Cyan
            Write-Host "  └─ [Q] Back to Main Menu ─────────────────────────────────" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Selection: 3  │  2-5  │  1,3,7  │  A (all)  │  ENTER (refresh)" -ForegroundColor DarkGray
            Write-Host "  Ranges auto-sort (10-3 = same as 3-10). Duplicates removed." -ForegroundColor DarkGray

            $selectedIndices = @()
            $UserWantsToQuit = $false
            $RefreshList     = $false

            while ($true) {
                $range = Read-Host "Enter choice (Numbers, 'A', 'S', 'Q', or ENTER)"
                Write-Log "[USER INPUT] Batch Range Selection: '$range'" "INFO" -FileOnly

                if ([string]::IsNullOrWhiteSpace($range)) {
                    Write-Log "Refreshing directory inventory..." "INFO"
                    $RefreshList = $true
                    break
                }
                if ($range -match "^[Qq]$") { Write-Log "Exiting Cloud Batch Mode..." "INFO"; $UserWantsToQuit = $true; break }

                if ($range -match "^[Ss]$") {
                    while ($true) {
                        $searchTerm = Read-Host "Enter FILE NAME or EXTENSION to search for (or 'Q' to exit search)"
                        Write-Log "[USER INPUT] Search Term: '$searchTerm'" "INFO" -FileOnly
                        if ($searchTerm -match "^[Qq]$") { break }
                        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                            $TargetCloudData = @{ SiteUrl = $SrcData.SiteUrl; FolderPath = $SrcData.FolderPath }
                            Perform-CloudSearch -CloudData $TargetCloudData -SearchTerm $searchTerm -SkipAuth
                            Write-Host "`nPress ENTER to return to the batch menu, or type a new search term..." -ForegroundColor DarkGray
                        }
                    }
                    $RefreshList = $true
                    break
                }

                $invalidInput = $false
                $tempIndices  = @()

                if ($range -match "^[Aa]$") {
                    if ($Items.Count -eq 0) { Write-Log "Nothing to process." "WARN"; $invalidInput = $true }
                    else {
                        $tempIndices = 0..($Items.Count - 1)
                        Write-Log "Selected ALL items (1 through $($Items.Count))." "INFO"
                    }
                } else {
                    $parts = $range -split ','
                    foreach ($part in $parts) {
                        $part = $part.Trim()
                        if ([string]::IsNullOrWhiteSpace($part)) { continue }

                        if ($part -match "^(\d+)\s*-\s*(\d+)$") {
                            $s = [int]$matches[1] - 1; $e = [int]$matches[2] - 1
                            if ($s -gt $e) { $temp = $s; $s = $e; $e = $temp }
                            if ($s -lt 0 -or $e -ge $Items.Count) {
                                Write-Host "ERROR: Range '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red
                                $invalidInput = $true; break
                            }
                            $tempIndices += $s..$e
                        } elseif ($part -match "^(\d+)$") {
                            $idx = [int]$matches[1] - 1
                            if ($idx -lt 0 -or $idx -ge $Items.Count) {
                                Write-Host "ERROR: Item '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red
                                $invalidInput = $true; break
                            }
                            $tempIndices += $idx
                        } else {
                            Write-Host "ERROR: Invalid format '$part'. Use numbers or ranges (e.g., 1, 3-5)." -ForegroundColor Red
                            $invalidInput = $true; break
                        }
                    }
                }

                if ($invalidInput -or $tempIndices.Count -eq 0) { Write-Host "Please try entering your selection again.`n" -ForegroundColor Yellow; continue }

                $selectedIndices = $tempIndices | Select-Object -Unique | Sort-Object
                Write-Log "Selected $($selectedIndices.Count) items for processing." "INFO"
                break
            }

            if ($UserWantsToQuit) { break }
            if ($RefreshList) { continue }

            $action = Read-Host "Type 'C' (Copy), 'M' (Move), 'D' (Delete), or 'V' (Verify) these items"
            Write-Log "[USER INPUT] Batch Action Selection: '$action'" "INFO" -FileOnly

            if ($action -notmatch "^[CcMmDdVv]$") {
                Write-Log "Invalid action. Aborting this chunk." "ERROR"
                continue
            }

            $isMove   = ($action -match "^[Mm]$")
            $isDelete = ($action -match "^[Dd]$")
            $isVerify = ($action -match "^[Vv]$")
            $actionWord = if ($isDelete) { "DELETE" } elseif ($isVerify) { "VERIFY" } elseif ($isMove) { "MOVE" } else { "COPY" }

            if ($isDelete) {
                Write-Host ""
                Write-Host "!!! WARNING !!! You selected to permanently DELETE items." -ForegroundColor Red
                $confirmDel = Read-Host "Type 'YES' to confirm this destruction"
                Write-Log "[USER INPUT] Cloud Batch Delete Confirmation: '$confirmDel'" "INFO" -FileOnly
                if ($confirmDel -cne 'YES') { Write-Log "Deletion aborted by user. Returning to index..." "WARN"; continue }
            }

            if (-not $isDelete -and -not $DstData) {
                Write-Log "You did not provide a DESTINATION URL! Cannot Copy, Move, or Verify." "ERROR"
                continue
            }

            Write-Log "Initializing Cloud Batch $actionWord..." "INFO"

            $global:JobCancelled = $false
            Clear-KeyBuffer

            $batchSw              = [System.Diagnostics.Stopwatch]::StartNew()
            $uiStopwatch          = [System.Diagnostics.Stopwatch]::StartNew()
            $jobStartTime         = Get-Date
            $startTimeStr         = $jobStartTime.ToString("HH:mm:ss")
            $totalItemsToProcess  = $selectedIndices.Count
            $currentItemProcessed = 0

            [long]$GrandSrcBytes = 0; [long]$GrandDstBytes = 0
            $GrandPerfect = 0; $GrandMismatch = 0; $GrandMissing = 0; $GrandOrphaned = 0
            $GrandSuccess = 0; $GrandFailed   = 0

            $BatchFailedList = [System.Collections.Generic.List[hashtable]]::new()

            foreach ($i in $selectedIndices) {
                $item = $Items[$i]
                $CurrentFolderContext = $SrcData.FolderPath
                $CurrentFileContext   = $item.Name

                if ($global:JobCancelled) { break }
                Check-Interrupt -CurrentFile $CurrentFileContext -CurrentFolder $CurrentFolderContext | Out-Null
                if ($global:JobCancelled) { break }

                $currentItemProcessed++

                if ($uiStopwatch.ElapsedMilliseconds -gt 5000 -or $currentItemProcessed -eq $totalItemsToProcess) {
                    $percentComplete = [math]::Min(100, [math]::Round(($currentItemProcessed / $totalItemsToProcess) * 100))
                    $elapsed    = (Get-Date) - $jobStartTime
                    $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                    # v109: consistency with other engines — show ETA + session total
                    $sessionStr = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
                    $eta = Get-ETA -Current $currentItemProcessed -Total $totalItemsToProcess -Elapsed $elapsed
                    $etaStr = if ($eta) { " | ETA: $eta" } else { "" }
                    Write-Progress -Id 0 -Activity "Cloud Batch Engine [Started: $startTimeStr]" -Status "Processing Item $currentItemProcessed of $($totalItemsToProcess): $($item.Name)" -PercentComplete $percentComplete -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
                    $uiStopwatch.Restart()
                }

                $ItemSrcUrl = "$($SrcData.FolderPath)/$($item.Name)" -replace '//', '/'

                Write-Host ""
                Write-Host "  ── [↑] CLOUD [$currentItemProcessed/$totalItemsToProcess] $($item.Name)" -ForegroundColor Magenta
                Write-Log "CLOUD ITEM $currentItemProcessed/$totalItemsToProcess : $($item.Name)" "INFO" -FileOnly

                try {
                    if ($isVerify) {
                        $ItemDstUrl = "$($DstData.FolderPath)/$($item.Name)" -replace '//', '/'
                        if ($item.GetType().Name -eq "Folder") {
                            $ItemSrcData = @{ SiteUrl = $SrcData.SiteUrl; FolderPath = $ItemSrcUrl }
                            $ItemDstData = @{ SiteUrl = $DstData.SiteUrl; FolderPath = $ItemDstUrl }
                            $DstSitePath = [uri]::new($DstData.SiteUrl).AbsolutePath
                            $DstSiteRel  = $ItemDstUrl.Substring($DstSitePath.Length).TrimStart('/')
                            $targetFolderExists = $true
                            try { Get-PnPFolder -Url $DstSiteRel -ErrorAction Stop | Out-Null } catch { $targetFolderExists = $false }

                            if (-not $targetFolderExists) {
                                Write-Log "MISSING: Folder '$($item.Name)' does not exist in the destination!" "ERROR"
                                $GrandMissing++
                            } else {
                                $vRes = Perform-CloudVerify -SrcData $ItemSrcData -DstData $ItemDstData -SkipAuth | Select-Object -Last 1
                                if ($vRes) {
                                    $GrandPerfect  += $vRes.Perfect
                                    $GrandMismatch += $vRes.Mismatch
                                    $GrandMissing  += $vRes.MissingInDst
                                    $GrandOrphaned += $vRes.MissingInSrc
                                    $GrandSrcBytes += $vRes.SrcBytes
                                    $GrandDstBytes += $vRes.DstBytes
                                }
                            }
                        } else {
                            $DstSitePath = [uri]::new($DstData.SiteUrl).AbsolutePath
                            $DstSiteRel  = $ItemDstUrl.Substring($DstSitePath.Length).TrimStart('/')
                            $dFile = Get-PnPFile -Url $DstSiteRel -ErrorAction SilentlyContinue
                            if (-not $dFile) {
                                Write-Log "MISSING: File '$($item.Name)' does not exist in the destination!" "ERROR"
                                $GrandMissing++
                                $GrandSrcBytes += $item.Length
                            } else {
                                $sizeStr = Format-FileSize $item.Length
                                $GrandSrcBytes += $item.Length
                                $GrandDstBytes += $dFile.Length
                                if ($item.Length -eq $dFile.Length) {
                                    Write-Log "Perfectly Synced (Exact Byte Size): $($item.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                                    $GrandPerfect++
                                } else {
                                    Write-Log "MISMATCH (Byte Size differs): $($item.Name) [Size: $sizeStr]" "ERROR"
                                    $GrandMismatch++
                                }
                            }
                        }
                    }
                    elseif ($isDelete) {
                        $delSw = [System.Diagnostics.Stopwatch]::StartNew()
                        if ($item.GetType().Name -eq "Folder") {
                            Invoke-PnPWithRetry -OperationName "DELETE FOLDER $($item.Name)" -ScriptBlock {
                                Remove-PnPFolder -Name $item.Name -Folder $SrcData.FolderPath -Force -Recycle -ErrorAction Stop | Out-Null
                            } | Out-Null
                        } else {
                            Invoke-PnPWithRetry -OperationName "DELETE FILE $($item.Name)" -ScriptBlock {
                                Remove-PnPFile -ServerRelativeUrl $ItemSrcUrl -Force -Recycle -ErrorAction Stop | Out-Null
                            } | Out-Null
                        }
                        $delSw.Stop()
                        Write-Host ""
                        Write-Log "--- CLOUD DELETE COMPLETE: $($item.Name) ---" "INFO"
                        Write-Log "Successfully Deleted: 1 Item (Sent to Recycle Bin)" "SUCCESS"
                        Write-Log "Total Time Elapsed: $([math]::Floor($delSw.Elapsed.TotalHours))h $($delSw.Elapsed.Minutes)m $($delSw.Elapsed.Seconds)s." "INFO"
                        $GrandSuccess++
                    } else {
                        $ItemDstUrl = "$($DstData.FolderPath)/$($item.Name)" -replace '//', '/'
                        $cmSw = [System.Diagnostics.Stopwatch]::StartNew()
                        if ($isMove) {
                            Invoke-PnPWithRetry -OperationName "MOVE $($item.Name)" -ScriptBlock {
                                Move-PnPFile -SourceUrl $ItemSrcUrl -TargetUrl $ItemDstUrl -Force -Overwrite -ErrorAction Stop
                            } | Out-Null
                        } else {
                            Invoke-PnPWithRetry -OperationName "COPY $($item.Name)" -ScriptBlock {
                                Copy-PnPFile -SourceUrl $ItemSrcUrl -TargetUrl $ItemDstUrl -Force -Overwrite -ErrorAction Stop
                            } | Out-Null
                        }
                        $cmSw.Stop()
                        Write-Host ""
                        Write-Log "--- CLOUD $actionWord COMPLETE: $($item.Name) ---" "INFO"
                        Write-Log "Successfully Processed Server-Side" "SUCCESS"
                        Write-Log "Total Time Elapsed: $([math]::Floor($cmSw.Elapsed.TotalHours))h $($cmSw.Elapsed.Minutes)m $($cmSw.Elapsed.Seconds)s." "INFO"
                        $GrandSuccess++
                    }
                } catch {
                    Write-Log "Failed: $($item.Name) [Reason: $($_.Exception.Message)]" "ERROR"
                    $GrandFailed++
                    $BatchFailedList.Add(@{ File = $item.Name; Path = $ItemSrcUrl; Reason = $_.Exception.Message })
                }
            }

            Write-Progress -Id 0 -Activity "Cloud Batch Engine [Started: $startTimeStr]" -Completed
            $batchSw.Stop()

            if ($BatchFailedList.Count -gt 0) {
                Write-Host ""
                Write-FailedReport -FailedList $BatchFailedList
            }

            if ($global:JobCancelled) { Write-Log "--- CLOUD BATCH JOB ABORTED BY USER ---" "WARN" }
            else                      { Write-Log "--- BATCH COMPLETION ---" "INFO" }

            Write-Host ""
            $summaryLines = [System.Collections.Generic.List[hashtable]]::new()
            if ($isVerify) {
                $summaryLines.Add(@{ Label = "Perfectly Synced";          Value = "$GrandPerfect";                               Level = "SUCCESS" })
                $summaryLines.Add(@{ Label = "Mismatched";                Value = "$GrandMismatch";                              Level = $(if($GrandMismatch -gt 0){"ERROR"}else{"SUCCESS"}) })
                $summaryLines.Add(@{ Label = "Missing in Destination";    Value = "$GrandMissing";                               Level = $(if($GrandMissing  -gt 0){"ERROR"}else{"SUCCESS"}) })
                $summaryLines.Add(@{ Label = "Orphaned in Destination";   Value = "$GrandOrphaned";                              Level = $(if($GrandOrphaned -gt 0){"WARN" }else{"SUCCESS"}) })
                $summaryLines.Add(@{ Label = "Source Size Verified";      Value = "$(Format-FileSize $GrandSrcBytes)";           Level = "INFO" })
                $summaryLines.Add(@{ Label = "Destination Size Verified"; Value = "$(Format-FileSize $GrandDstBytes)";           Level = "INFO" })
                $matchLevel = if($GrandSrcBytes -eq $GrandDstBytes -and $GrandSrcBytes -gt 0){"SUCCESS"} elseif($GrandSrcBytes -eq 0 -and $GrandDstBytes -eq 0){"WARN"} else{"ERROR"}
                $matchVal   = if($GrandSrcBytes -eq $GrandDstBytes -and $GrandSrcBytes -gt 0){"YES"} elseif($GrandSrcBytes -eq 0 -and $GrandDstBytes -eq 0){"N/A (Empty)"} else{"NO — MISMATCH"}
                $summaryLines.Add(@{ Label = "Size Match";                Value = $matchVal;                                     Level = $matchLevel })
            } elseif ($isDelete) {
                $summaryLines.Add(@{ Label = "Items Deleted";             Value = "$GrandSuccess";                               Level = "SUCCESS" })
                $summaryLines.Add(@{ Label = "Deletion Failures";         Value = "$GrandFailed";                                Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) })
            } else {
                $summaryLines.Add(@{ Label = "Successfully Processed";    Value = "$GrandSuccess items";                         Level = "SUCCESS" })
                $summaryLines.Add(@{ Label = "Failed to Transfer";        Value = "$GrandFailed";                                Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) })
            }
            $summaryLines.Add(@{ Label = "Total Time";                    Value = "$([math]::Floor($batchSw.Elapsed.TotalHours))h $($batchSw.Elapsed.Minutes)m $($batchSw.Elapsed.Seconds)s"; Level = "INFO" })
            # v109: session total for cross-operation context (cloud batch is server-side
            # so no local bytes flow — but the running session total is still useful info)
            if ($global:SessionTransferredBytes -gt 0) {
                $summaryLines.Add(@{ Label = "Session Total";             Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
            }
            Write-SummaryBox -Title "CLOUD BATCH COMPLETE" -Lines $summaryLines.ToArray() -BoxColor "Yellow"
            # v111: failure-rate alert + ALL distinct errors recap (no local file list — server-side op)
            Show-OperationDiagnostics -SuccessCount $GrandSuccess -FailedCount $GrandFailed -FailedList $BatchFailedList -ProcessedItems $null
            Write-Host ""
            Write-Host "  Scanning remaining items for next batch..." -ForegroundColor DarkGray
        }
    } catch {
        Write-Log "Cloud Batch Encountered an Error: $($_.Exception.Message)" "ERROR"
    } finally {
        try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
    }
}

# --- FUNCTION: Get-CloudInventory ---
function Get-CloudInventory {
    param([string]$SiteUrl, [string]$FolderPath, [string]$Tag, [switch]$SkipAuth)
    try {
        if (-not $SkipAuth) {
            $originalWarning = $WarningPreference
            $WarningPreference = "SilentlyContinue"
            Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop
            $WarningPreference = $originalWarning
        }
        $SitePath   = [uri]::new($SiteUrl).AbsolutePath
        $BaseRelUrl = $FolderPath.Substring($SitePath.Length).TrimStart('/')
        $Queue      = [System.Collections.Generic.Queue[hashtable]]::new()
        $Queue.Enqueue(@{ CloudUrl = $BaseRelUrl; RelPath = "" })
        $Files = @(); $Folders = @()

        while ($Queue.Count -gt 0) {
            $current = $Queue.Dequeue()
            if ($global:JobCancelled) { break }
            Check-Interrupt -CurrentFolder $current.CloudUrl -CurrentFile "Scanning Directory" | Out-Null
            if ($global:JobCancelled) { break }
            if ($current.RelPath -ne "") { Write-Host " Scanning $Tag : /$($current.RelPath)" -ForegroundColor DarkGray }
            else                         { Write-Host " Scanning $Tag : /Root" -ForegroundColor DarkGray }
            Start-Sleep -Milliseconds 100
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $current.CloudUrl -ErrorAction SilentlyContinue
            if ($items) {
                foreach ($item in $items) {
                    $itemRelPath = if ($current.RelPath) { "$($current.RelPath)/$($item.Name)" } else { $item.Name }
                    if ($item.GetType().Name -eq "Folder") {
                        $Folders += $itemRelPath
                        $Queue.Enqueue(@{ CloudUrl = "$($current.CloudUrl)/$($item.Name)"; RelPath = $itemRelPath })
                    } else {
                        $Files += @{ RelPath = $itemRelPath; Name = $item.Name; Length = $item.Length }
                    }
                }
            }
        }
        return @{ Files = $Files; Folders = $Folders }
    } catch { throw $_ }
}

# --- FUNCTION: Perform-CloudVerify ---
function Perform-CloudVerify {
    param([hashtable]$SrcData, [hashtable]$DstData, [switch]$SkipAuth)
    Write-Log "--- STARTING CLOUD VERIFICATION ENGINE ---" "INFO"
    Write-Log "SOURCE FOLDER: $($SrcData.FolderPath)"
    Write-Log "DESTINATION FOLDER: $($DstData.FolderPath)"

    if (-not $SkipAuth) { $global:JobCancelled = $false; Clear-KeyBuffer }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Host "Connecting and deep-scanning Source Cloud Directory..." -ForegroundColor DarkGray
    $SrcInv = Get-CloudInventory -SiteUrl $SrcData.SiteUrl -FolderPath $SrcData.FolderPath -Tag "Source" -SkipAuth:$SkipAuth
    if ($global:JobCancelled) { return }

    $DstSkip = if ($SkipAuth -or ($SrcData.SiteUrl -eq $DstData.SiteUrl)) { $true } else { $false }

    Write-Host "Connecting and deep-scanning Destination Cloud Directory..." -ForegroundColor DarkGray
    $DstInv = Get-CloudInventory -SiteUrl $DstData.SiteUrl -FolderPath $DstData.FolderPath -Tag "Destination" -SkipAuth:$DstSkip
    if ($global:JobCancelled) { return }

    $SrcHash = @{}; foreach ($f in $SrcInv.Files) { $SrcHash[$f.RelPath] = $f }
    $DstHash = @{}; foreach ($f in $DstInv.Files) { $DstHash[$f.RelPath] = $f }

    $MissingInDst = 0; $MissingInSrc = 0; $Mismatch = 0; $Perfect = 0
    [long]$SrcTotalBytes = 0; [long]$DstTotalBytes = 0

    Write-Log "Cross-referencing $($SrcInv.Files.Count) source files against $($DstInv.Files.Count) destination files..." "INFO"

    foreach ($key in $SrcHash.Keys) {
        if (-not $DstHash.ContainsKey($key)) {
            $MissingInDst++
            Write-Log "MISSING (Failed to transfer): $key" "WARN"
            $SrcTotalBytes += $SrcHash[$key].Length
        } else {
            $sFile = $SrcHash[$key]; $dFile = $DstHash[$key]
            $SrcTotalBytes += $sFile.Length; $DstTotalBytes += $dFile.Length
            $sizeStr = Format-FileSize $sFile.Length
            if ($sFile.Length -ne $dFile.Length) {
                $Mismatch++
                Write-Log "MISMATCH (Name matches, but Byte Size differs): $key [Size: $sizeStr]" "ERROR"
            } else {
                $Perfect++
                Write-Log "Perfectly Synced: $key [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
            }
        }
    }

    foreach ($key in $DstHash.Keys) {
        if (-not $SrcHash.ContainsKey($key)) {
            $MissingInSrc++
            $sizeStr = Format-FileSize $DstHash[$key].Length
            Write-Log "ORPHANED (Exists in Dest, but not Source): $key [Size: $sizeStr]" "WARN"
        }
    }

    $sw.Stop()
    Write-Host ""
    Write-SummaryBox -Title "CLOUD VERIFICATION RESULTS" -Lines @(
        @{ Label = "Source";              Value = "$($SrcInv.Folders.Count) folders  $($SrcInv.Files.Count) files  $(Format-FileSize $SrcTotalBytes)"; Level = "INFO" }
        @{ Label = "Destination";         Value = "$($DstInv.Folders.Count) folders  $($DstInv.Files.Count) files  $(Format-FileSize $DstTotalBytes)"; Level = "INFO" }
        @{ Label = "Perfectly Synced";    Value = "$Perfect";   Level = "SUCCESS" }
        @{ Label = "Mismatched (corrupt)";Value = "$Mismatch";  Level = $(if($Mismatch   -gt 0){"ERROR"}else{"SUCCESS"}) }
        @{ Label = "Missing in Dest";     Value = "$MissingInDst"; Level = $(if($MissingInDst -gt 0){"ERROR"}else{"SUCCESS"}) }
        @{ Label = "Orphaned in Dest";    Value = "$MissingInSrc"; Level = $(if($MissingInSrc -gt 0){"WARN" }else{"SUCCESS"}) }
        @{ Label = "Verification Time";   Value = "$([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s"; Level = "INFO" }
    ) -BoxColor "Yellow"

    return @{ Perfect = $Perfect; Mismatch = $Mismatch; MissingInDst = $MissingInDst; MissingInSrc = $MissingInSrc; SrcBytes = $SrcTotalBytes; DstBytes = $DstTotalBytes }
}

# --- FUNCTION: Perform-CloudWipe ---
function Perform-CloudWipe {
    param([hashtable]$TargetData)
    Write-Log "--- STARTING CLOUD DIRECTORY WIPE ---" "WARN"
    Write-Log "TARGET: $($TargetData.FolderPath)"
    Write-Host "!!! DANGER ZONE !!!" -ForegroundColor Red

    $confirm = Read-Host "Type 'YES' to permanently send this to the cloud recycle bin"
    Write-Log "[USER INPUT] Cloud Wipe Confirmation: '$confirm'" "INFO" -FileOnly

    if ($confirm -cne 'YES') { return }

    $global:JobCancelled = $false
    Clear-KeyBuffer

    try {
        Write-Log "Authenticating to Site: $($TargetData.SiteUrl)..." "WARN"
        $originalWarning = $WarningPreference
        $WarningPreference = "SilentlyContinue"
        Connect-PnPOnline -Url $TargetData.SiteUrl -UseWebLogin -ErrorAction Stop
        $WarningPreference = $originalWarning

        $ParentPath = (Split-Path $TargetData.FolderPath -Parent) -replace '\\', '/'
        $FolderName = Split-Path $TargetData.FolderPath -Leaf

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Log "Executing Server-Side Deletion..." "INFO"
        Remove-PnPFolder -Name $FolderName -Folder $ParentPath -Force -Recycle -ErrorAction Stop
        $sw.Stop()

        Write-Log "SUCCESSFULLY WIPED DIRECTORY: $($TargetData.FolderPath)" "SUCCESS"
        Write-Log "Total Time Elapsed: $([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s." "SUCCESS"
    } catch {
        Write-Log "Failed to wipe directory. Reason: $($_.Exception.Message)" "ERROR"
    } finally {
        try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
    }
}

# ========================================================================================
# ENGINE SUITE C: CROSS-TENANT & LOCAL FILE MANAGER TOOLS
# ========================================================================================

# --- FUNCTION: Perform-LocalTransfer ---

# --- FUNCTION: Copy-FileWithProgress ---
# POWERSHELL LESSON: Why [System.IO.File]::Copy() gives no progress feedback
#   The built-in Copy() call is a single blocking OS call — the OS copies bytes in the
#   kernel and PowerShell gets control back only when it is done. There is no hook
#   for PowerShell to interrupt and report progress mid-copy.
#
#   The workaround: open both files as streams and copy in small chunks (4 MB each).
#   Between chunks we can call Write-Progress to update the bar. This is slightly
#   slower than a kernel copy (one extra user/kernel boundary per chunk) but for
#   files ≥ 1 GB the user-experience benefit far outweighs the small overhead.
#
# Uses Write-Progress -Id 2 so it renders as a NESTED child bar under the parent
# file-count bar (-Id 1 used in Perform-LocalTransfer) — both visible simultaneously.
function Copy-FileWithProgress {
    param(
        [string]$Source,
        [string]$Dest,
        [string]$FileName,
        [long]$FileSize
    )

    $bufSize = 4 * 1024 * 1024   # 4 MB per chunk — good balance of speed and progress granularity
    $srcStream = $dstStream = $null
    # v109: three stopwatches with distinct purposes:
    #   $totalSw  — runs from start to finish; used to compute ETA from overall avg
    #   $uiSw     — fires every 400 ms to refresh the progress bar
    #   $speedSw  — 2-second rolling window for the "current" speed display
    $totalSw = [System.Diagnostics.Stopwatch]::StartNew()
    $uiSw    = [System.Diagnostics.Stopwatch]::StartNew()
    $speedSw = [System.Diagnostics.Stopwatch]::StartNew()
    [long]$written    = 0
    [long]$speedBytes = 0
    $speedStr = ""

    try {
        $srcStream = [System.IO.File]::OpenRead($Source)
        $dstStream = [System.IO.File]::Create($Dest)
        $buf       = [byte[]]::new($bufSize)

        while ($true) {
            $read = $srcStream.Read($buf, 0, $bufSize)
            if ($read -eq 0) { break }

            $dstStream.Write($buf, 0, $read)
            $written    += $read
            $speedBytes += $read

            if ($uiSw.ElapsedMilliseconds -gt 400) {
                $pct   = [math]::Min(100, [int](($written / $FileSize) * 100))
                $done  = Format-FileSize $written
                $total = Format-FileSize $FileSize

                # Recalculate transfer speed every 2 s to avoid jitter
                if ($speedSw.ElapsedMilliseconds -gt 2000 -and $speedSw.Elapsed.TotalSeconds -gt 0) {
                    $speedStr   = "$(Format-FileSize ([long]($speedBytes / $speedSw.Elapsed.TotalSeconds)))/s"
                    $speedBytes = 0; $speedSw.Restart()
                }

                # v109: ETA based on overall avg rate (not the 2s rolling window).
                # Using the overall rate smooths out short-term I/O jitter so the
                # ETA doesn't jump around wildly.
                $eta = Get-ETA -Current $written -Total $FileSize -Elapsed $totalSw.Elapsed
                $etaStr = if ($eta) { " | ETA: $eta" } else { "" }

                Write-Progress -Id 2 -ParentId 1 `
                    -Activity " ▶ Large file copying" `
                    -Status   "  $FileName" `
                    -PercentComplete $pct `
                    -CurrentOperation "  $done of $total  $speedStr$etaStr"
                $uiSw.Restart()
            }
        }
    } finally {
        if ($srcStream) { $srcStream.Close(); $srcStream.Dispose() }
        if ($dstStream) { $dstStream.Close(); $dstStream.Dispose() }
        Write-Progress -Id 2 -Activity " ▶ Large file copying" -Completed
    }
}

function Perform-LocalTransfer {
    param(
        [string]$Src,
        [string]$Dst,
        [switch]$IsMove,
        [datetime]$InheritedStartTime = [datetime]::MinValue,
        [string]$BatchStatus  = "",
        [int]$BatchPercent    = -1,
        [switch]$SkipStorageCheck
    )
    $ModeName  = if ($IsMove) { "MOVE" } else { "COPY" }
    $TargetName = Split-Path $Src -Leaf
    $FailedList = [System.Collections.Generic.List[hashtable]]::new()

    Write-Log "--- LOCAL SYNC ENGINE ($ModeName) ---"
    Write-Log "SOURCE: $(Format-PathWithStorageTag $Src)"
    Write-Log "DESTINATION: $(Format-PathWithStorageTag $Dst)"

    $global:JobCancelled = $false
    Clear-KeyBuffer

    if (-not $SkipStorageCheck) {
        Write-Host "Calculating payload for pre-flight storage check..." -ForegroundColor DarkGray
        $metrics    = Get-LocalDirectoryMetrics -Path $Src
        $TotalBytes = if ($metrics) { $metrics.RawBytes } else { 0 }
        $batchPairs = @( @{ Src = $Src; Dst = $Dst } )
        $canProceed = Invoke-PreflightStorageCheck -DestRoot $Dst -TotalSourceBytes $TotalBytes -BatchItems $batchPairs
        if (-not $canProceed) { Write-Log "Operation aborted due to insufficient storage space." "WARN"; return $null }
    }

    # v110: heartbeat-aware enumeration so very large folders don't appear hung
    # during the initial scan. Falls back to Get-ChildItem internally if the .NET
    # enumerator hits permission errors. Sort-Object preserves the original
    # deterministic-order behavior callers depended on.
    # Pass $Src directly (not long-pathed) — items' FullName must match $Src
    # style so the later $item.FullName.Substring($Src.Length) slicing works.
    $Items       = Get-ItemsWithHeartbeat -LongPath $Src -Label "Scanning source" | Sort-Object FullName
    $TotalItems  = $Items.Count
    $CurrentItemIndex = 0

    $SuccessCount = 0; $SkippedCount = 0; $FailedCount = 0; $TooLongCount = 0
    [long]$SuccessBytes = 0; [long]$SkippedBytes = 0

    $sw          = [System.Diagnostics.Stopwatch]::StartNew()
    $uiStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $jobStartTime = if ($InheritedStartTime -ne [datetime]::MinValue) { $InheritedStartTime } else { Get-Date }
    $startTimeStr = $jobStartTime.ToString("HH:mm:ss")

    foreach ($item in $Items) {
        $CurrentFolderContext = if ($item.PSIsContainer) { $item.FullName } else { $item.DirectoryName }
        $CurrentFileContext   = $item.Name

        if ($global:JobCancelled) { break }
        Check-Interrupt -CurrentFile $CurrentFileContext -CurrentFolder $CurrentFolderContext | Out-Null
        if ($global:JobCancelled) { break }

        $CurrentItemIndex++

        if ($uiStopwatch.ElapsedMilliseconds -gt 5000 -or $CurrentItemIndex -eq $TotalItems) {
            $filePercent = [math]::Min(100, [math]::Round(($CurrentItemIndex / $TotalItems) * 100))
            $elapsed     = (Get-Date) - $jobStartTime
            $elapsedStr  = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
            $sessionStr  = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
            # v109: ETA based on file count progress so far
            $eta = Get-ETA -Current $CurrentItemIndex -Total $TotalItems -Elapsed $elapsed
            $etaStr = if ($eta) { " | ETA: $eta" } else { "" }
            if ($BatchStatus) {
                # v110: nested child bar — shows FILE-level progress (not duplicated batch
                # progress as in v108/v109). The parent Id 0 bar already shows item-of-N
                # at the batch level; this child fills in "and we're currently on file X of Y
                # within that item". -ParentId 0 makes the renderer indent it under the batch.
                Write-Progress -Id 1 -ParentId 0 -Activity "  └─ Local Sync Engine ($ModeName)" -Status "File $CurrentItemIndex of $TotalItems" -PercentComplete $filePercent -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
            } else {
                Write-Progress -Id 1 -Activity "Local Sync Engine ($ModeName) [Started: $startTimeStr]" -Status "Processing File $CurrentItemIndex of $TotalItems" -PercentComplete $filePercent -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
            }
            $uiStopwatch.Restart()
        }

        $StandardSourcePath = $item.FullName
        $RelativePath       = $StandardSourcePath.Substring($Src.Length).TrimStart('\')
        $StandardDestPath   = Join-Path $Dst $RelativePath

        $sizeStr = ""
        if (-not $item.PSIsContainer) { $sizeStr = Format-FileSize $item.Length }

        if ($StandardDestPath.Length -ge $CloudPathLimit) {
            Write-Log "SKIPPED: Exceeds Limit ($($StandardDestPath.Length)/$CloudPathLimit char): $RelativePath [Size: $sizeStr]" "WARN" -ConsoleOnly
            $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            "[$Timestamp] OVER LIMIT ($($StandardDestPath.Length) chars) - Src: $StandardSourcePath | Dst: $StandardDestPath" | Out-File -FilePath $LongPathReport -Append -Encoding UTF8
            $TooLongCount++
            if (-not $item.PSIsContainer) { $SkippedBytes += $item.Length }
            continue
        }

        $SourceItemPath = Get-LongPath $StandardSourcePath
        $DestItemPath   = Get-LongPath $StandardDestPath

        if ($item.PSIsContainer) {
            if (-not [System.IO.Directory]::Exists($DestItemPath)) {
                try { [System.IO.Directory]::CreateDirectory($DestItemPath) | Out-Null } catch {}
            }
        } else {
            $DestDirPath     = [System.IO.Path]::GetDirectoryName($StandardDestPath)
            $DestDirLongPath = Get-LongPath $DestDirPath
            if (-not [System.IO.Directory]::Exists($DestDirLongPath)) {
                try { [System.IO.Directory]::CreateDirectory($DestDirLongPath) | Out-Null } catch {}
            }

            $ShouldTransfer = $true

            if ([System.IO.File]::Exists($DestItemPath)) {
                $SourceFileInfo = [System.IO.FileInfo]::new($SourceItemPath)
                $DestFileInfo   = [System.IO.FileInfo]::new($DestItemPath)
                if (($SourceFileInfo.Length -eq $DestFileInfo.Length) -and ($SourceFileInfo.LastWriteTime.ToString() -eq $DestFileInfo.LastWriteTime.ToString())) {
                    $ShouldTransfer = $false
                    $SkippedCount++
                    $SkippedBytes += $SourceFileInfo.Length
                    Write-Log "SKIPPED (Already Synced): $RelativePath [Size: $sizeStr]" "WARN" -ConsoleOnly
                    if ($IsMove) { Invoke-ForceDelete -Path $SourceItemPath | Out-Null }
                } else {
                    if ($IsMove) { Invoke-ForceDelete -Path $DestItemPath | Out-Null }
                }
            }

            if ($ShouldTransfer) {
                try {
                    if ($IsMove) {
                        [System.IO.File]::Move($SourceItemPath, $DestItemPath)
                    } elseif ($item.Length -ge $script:LargeFileBytes) {
                        # Large file: stream copy with nested progress bar (Id=2, parented to Id=1)
                        # Forces the copy to happen in 4 MB chunks so we can report bytes written.
                        Copy-FileWithProgress -Source $SourceItemPath -Dest $DestItemPath `
                            -FileName $item.Name -FileSize $item.Length
                    } else {
                        [System.IO.File]::Copy($SourceItemPath, $DestItemPath, $true)
                    }

                    $OriginalDates = [System.IO.FileInfo]::new($SourceItemPath)
                    if (-not $OriginalDates.Exists -and $IsMove) { $OriginalDates = [System.IO.FileInfo]::new($DestItemPath) }
                    [System.IO.File]::SetCreationTime($DestItemPath, $OriginalDates.CreationTime)
                    [System.IO.File]::SetLastWriteTime($DestItemPath, $OriginalDates.LastWriteTime)

                    $SuccessCount++
                    $SuccessBytes += $OriginalDates.Length
                    $global:SessionTransferredBytes += $OriginalDates.Length   # running session total
                    Write-Log "$ModeName [$CurrentItemIndex/$TotalItems]: $RelativePath [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                } catch {
                    Write-Log "Failed: $RelativePath [Reason: $($_.Exception.Message)]" "ERROR"
                    $FailedCount++
                    $FailedList.Add(@{ File = $item.Name; Path = $RelativePath; Reason = $_.Exception.Message })
                }
            }
        }
    }

    if ($InheritedStartTime -eq [datetime]::MinValue) { Write-Progress -Id 1 -Activity "Local Sync Engine ($ModeName) [Started: $startTimeStr]" -Completed }
    if ($IsMove -and -not $global:JobCancelled) { Remove-EmptyDirectories -TargetFolder $Src }
    $sw.Stop()

    Write-Host ""
    if (-not $BatchStatus -and $FailedList.Count -gt 0) {
        Write-FailedReport -FailedList $FailedList
    }

    $statusTitle = if ($global:JobCancelled) { "LOCAL SYNC ABORTED" } else { "LOCAL SYNC COMPLETE" }
    $statusColor = if ($global:JobCancelled) { "Yellow" } else { "Green" }
    $syncLines = [System.Collections.Generic.List[hashtable]]::new()
    # v110: storage type at the top so operators see what kind of drive(s) this job touches
    $srcTag = Get-PathStorageType $Src; $dstTag = Get-PathStorageType $Dst
    if ($srcTag -or $dstTag) {
        $arrow = if ($srcTag -and $dstTag) { "$srcTag -> $dstTag" } elseif ($srcTag) { $srcTag } else { $dstTag }
        $syncLines.Add(@{ Label = "Storage";                Value = $arrow;                                       Level = "INFO" })
    }
    $syncLines.Add(@{ Label = "Target";                Value = $TargetName;                                   Level = "INFO" })
    $syncLines.Add(@{ Label = "Successfully Processed";Value = "$SuccessCount  |  $(Format-FileSize $SuccessBytes)"; Level = "SUCCESS" })
    $syncLines.Add(@{ Label = "Skipped (already synced)"; Value = "$SkippedCount  |  $(Format-FileSize $SkippedBytes)"; Level = "INFO" })
    $syncLines.Add(@{ Label = "Failed";                Value = "$FailedCount";                                Level = $(if($FailedCount -gt 0){"ERROR"}else{"SUCCESS"}) })
    if ($TooLongCount -gt 0) {
        $syncLines.Add(@{ Label = "Skipped (path too long)"; Value = "$TooLongCount (see LongPaths log)";   Level = "WARN" })
    }
    $syncLines.Add(@{ Label = "Total Time";            Value = "$([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s"; Level = "INFO" })
    # v109: Average throughput — only meaningful if we actually moved bytes
    $avgSpeed = Format-AvgSpeed -Bytes $SuccessBytes -Elapsed $sw.Elapsed
    if ($avgSpeed) {
        $syncLines.Add(@{ Label = "Average Speed";      Value = $avgSpeed;                                     Level = "INFO" })
    }
    # v110: files/min — useful when bytes don't tell the whole story (lots of tiny files)
    $filesPerMin = Format-FilesPerMin -Count $SuccessCount -Elapsed $sw.Elapsed
    if ($filesPerMin -and $SuccessCount -ge 50) {
        $syncLines.Add(@{ Label = "Throughput";         Value = $filesPerMin;                                  Level = "INFO" })
    }
    if ($global:SessionTransferredBytes -gt 0) {
        $syncLines.Add(@{ Label = "Session Total";     Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
    }
    if (-not $BatchStatus) {
        Write-SummaryBox -Title $statusTitle -Lines $syncLines.ToArray() -BoxColor $statusColor
        # v111: failure-rate alert + ALL distinct errors recap + per-extension breakdown
        # ($Items contains every FileSystemInfo we enumerated — passing the files-only
        # subset gives the breakdown a cleaner picture by skipping folder entries.)
        $itemFiles = @($Items | Where-Object { -not $_.PSIsContainer })
        Show-OperationDiagnostics -SuccessCount $SuccessCount -FailedCount $FailedCount -FailedList $FailedList -ProcessedItems $itemFiles
    }
    else {
        # When called from a batch, just log the key numbers without printing a box (batch has its own summary)
        Write-Log "$statusTitle : $TargetName | Processed: $SuccessCount | Skipped: $SkippedCount | Failed: $FailedCount | $(Format-FileSize $SuccessBytes)" $(if($global:JobCancelled){"WARN"}else{"INFO"})
    }

    return @{ Success = $SuccessCount; Skipped = $SkippedCount; Failed = $FailedCount; SuccessBytes = $SuccessBytes; SkippedBytes = $SkippedBytes; FailedItems = $FailedList }
}

# --- FUNCTION: Perform-LocalVerify ---
function Perform-LocalVerify {
    param([string]$Src, [string]$Dst)
    Write-Log "--- STARTING LOCAL VERIFICATION ENGINE ---" "INFO"
    Write-Log "SOURCE: $(Format-PathWithStorageTag $Src)"
    Write-Log "DESTINATION: $(Format-PathWithStorageTag $Dst)"

    $global:JobCancelled = $false
    Clear-KeyBuffer

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Host ""
    $RootItems = @(Get-ChildItem -LiteralPath $Src -Force | Sort-Object @{Expression={$_.PSIsContainer};Descending=$true}, @{Expression={$_.Name};Ascending=$true})
    $MenuAudit = [System.Collections.Generic.List[string]]::new()
    $MenuAudit.Add("--- SOURCE ROOT INVENTORY ---")
    Write-Host "--- SOURCE ROOT INVENTORY ---" -ForegroundColor Cyan

    if ($RootItems.Count -eq 0) {
        Write-Host " [Directory is completely empty]" -ForegroundColor DarkGray
        $MenuAudit.Add(" [Directory is completely empty]")
    } else {
        foreach ($item in $RootItems) {
            if ($item.PSIsContainer) {
                $subItems = @(Get-ChildItem -LiteralPath $item.FullName -Recurse -Force)
                $fCount   = @($subItems | Where-Object { -not $_.PSIsContainer }).Count
                $dCount   = @($subItems | Where-Object {  $_.PSIsContainer }).Count
                $logStr   = " [FOLDER] $($item.Name) (Contains: $dCount Sub-folders, $fCount Files)"
            } else {
                $logStr = " [FILE]   $($item.Name) [Size: $(Format-FileSize $item.Length)]"
            }
            Write-Host $logStr
            $MenuAudit.Add($logStr)
        }
    }
    Write-Host "-----------------------------" -ForegroundColor Cyan
    $MenuAudit.Add("-----------------------------")
    Write-Host ""

    $ts         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogPayload = ""
    foreach ($line in $MenuAudit) { $LogPayload += "[$ts] [INFO] $line`r`n" }
    [System.IO.File]::AppendAllText($LogFile, $LogPayload)

    Write-Host "Scanning Deep Source Directory... This may take a moment for massive folders." -ForegroundColor DarkGray
    $SrcAll   = @(Get-ChildItem -LiteralPath $Src -Recurse -Force)
    $SrcFiles = @($SrcAll | Where-Object { -not $_.PSIsContainer })
    $SrcDirs  = @($SrcAll | Where-Object {  $_.PSIsContainer })

    Write-Host "Scanning Deep Destination Directory..." -ForegroundColor DarkGray
    $DstAll   = @(Get-ChildItem -LiteralPath $Dst -Recurse -Force)
    $DstFiles = @($DstAll | Where-Object { -not $_.PSIsContainer })
    $DstDirs  = @($DstAll | Where-Object {  $_.PSIsContainer })

    $SrcHash = @{}
    foreach ($f in $SrcFiles) {
        $relPath = $f.FullName.Substring($Src.Length).TrimStart('\')
        $SrcHash[$relPath] = $f
    }
    $DstHash = @{}
    foreach ($f in $DstFiles) {
        $relPath = $f.FullName.Substring($Dst.Length).TrimStart('\')
        $DstHash[$relPath] = $f
    }

    $MissingInDst = 0; $MissingInSrc = 0; $Mismatch = 0; $Perfect = 0
    [long]$SrcTotalBytes = 0; [long]$DstTotalBytes = 0

    Write-Log "Cross-referencing $($SrcFiles.Count) source files against $($DstFiles.Count) destination files..." "INFO"

    foreach ($key in $SrcHash.Keys) {
        if (-not $DstHash.ContainsKey($key)) {
            $MissingInDst++
            Write-Log "MISSING (Failed to transfer): $key" "WARN"
            $SrcTotalBytes += $SrcHash[$key].Length
        } else {
            $sFile = $SrcHash[$key]; $dFile = $DstHash[$key]
            $SrcTotalBytes += $sFile.Length; $DstTotalBytes += $dFile.Length
            $sizeStr = Format-FileSize $sFile.Length
            if ($sFile.Length -ne $dFile.Length -or $sFile.LastWriteTime.ToString() -ne $dFile.LastWriteTime.ToString()) {
                $Mismatch++
                Write-Log "MISMATCH (Name matches, but Size or Date differs): $key [Size: $sizeStr]" "ERROR"
            } else {
                $Perfect++
                Write-Log "Perfectly Synced: $key [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
            }
        }
    }

    foreach ($key in $DstHash.Keys) {
        if (-not $SrcHash.ContainsKey($key)) {
            $MissingInSrc++
            Write-Log "ORPHANED (Exists in Dest, but not Source): $key [Size: $(Format-FileSize $DstHash[$key].Length)]" "WARN"
        }
    }

    $sw.Stop()
    Write-Host ""
    Write-SummaryBox -Title "LOCAL VERIFICATION RESULTS" -Lines @(
        @{ Label = "Source";              Value = "$($SrcDirs.Count) folders  $($SrcFiles.Count) files  $(Format-FileSize $SrcTotalBytes)"; Level = "INFO" }
        @{ Label = "Destination";         Value = "$($DstDirs.Count) folders  $($DstFiles.Count) files  $(Format-FileSize $DstTotalBytes)"; Level = "INFO" }
        @{ Label = "Perfectly Synced";    Value = "$Perfect";      Level = "SUCCESS" }
        @{ Label = "Mismatched";          Value = "$Mismatch";     Level = $(if($Mismatch    -gt 0){"ERROR"}else{"SUCCESS"}) }
        @{ Label = "Missing in Dest";     Value = "$MissingInDst"; Level = $(if($MissingInDst -gt 0){"ERROR"}else{"SUCCESS"}) }
        @{ Label = "Orphaned in Dest";    Value = "$MissingInSrc"; Level = $(if($MissingInSrc -gt 0){"WARN" }else{"SUCCESS"}) }
        @{ Label = "Verification Time";   Value = "$([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s"; Level = "INFO" }
    ) -BoxColor "Green"

    return @{ Perfect = $Perfect; Mismatch = $Mismatch; MissingInDst = $MissingInDst; MissingInSrc = $MissingInSrc; SrcBytes = $SrcTotalBytes; DstBytes = $DstTotalBytes }
}

# --- FUNCTION: Perform-LocalBatch ---
function Perform-LocalBatch {
    param([string]$Src, [string]$Dst)
    Write-Log "--- ENTERING LOCAL BATCH MODE ---"
    Write-Log "SOURCE FOLDER: $(Format-PathWithStorageTag $Src)"
    if ($Dst) { Write-Log "DESTINATION FOLDER: $(Format-PathWithStorageTag $Dst)" }

    while ($true) {
        Write-PathBanner -Label1 "Source" -Path1 $Src -Label2 "Dest" -Path2 $(if($Dst){$Dst}else{""}) -BoxColor "Green"
        $Items = @(Get-ChildItem -LiteralPath $Src -Force | Sort-Object @{Expression={$_.PSIsContainer};Descending=$true}, @{Expression={$_.Name};Ascending=$true})

        if ($Items.Count -eq 0) {
            Write-Log "Source directory is completely empty. Exiting Batch Mode..." "WARN"
            break
        }

        Write-Host ""
        Write-Host "  ┌─ LOCAL DIRECTORY INDEX ──────────────────────────────" -ForegroundColor Green
        Write-Host "  │  Items = top-level count  │ Size = total recursive size" -ForegroundColor DarkGray
        Write-Host " Calculating local folder sizes..." -ForegroundColor DarkGray

        $MenuAudit = [System.Collections.Generic.List[string]]::new()
        $MenuAudit.Add("--- TOP-LEVEL DIRECTORY INDEX ---")

        for ($i = 0; $i -lt $Items.Count; $i++) {
            $item = $Items[$i]
            if ($item.PSIsContainer) {
                $metrics = Get-LocalDirectoryMetrics -Path $item.FullName
                if ($metrics) {
                    $sizeStr = $metrics.Size
                    try { $iCount = [System.IO.Directory]::GetFileSystemEntries((Get-LongPath $item.FullName)).Length }
                    catch { $iCount = 0 }
                    $logStr = "[$($i + 1)] [FOLDER] $($item.Name) [Size: $sizeStr | Items: $iCount]"
                } else {
                    $logStr = "[$($i + 1)] [FOLDER] $($item.Name) [Size: Empty/Locked | Items: 0]"
                }
            } else {
                $logStr = "[$($i + 1)] [FILE]   $($item.Name) [Size: $(Format-FileSize $item.Length)]"
            }
            Write-Host " $logStr"
            $MenuAudit.Add($logStr)
        }

        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogPayload = ""
        foreach ($line in $MenuAudit) { $LogPayload += "[$ts] [MENU] $line`r`n" }
        [System.IO.File]::AppendAllText($LogFile, $LogPayload)

        Write-Host ""
        Write-Host " [A] Process ALL items in this directory" -ForegroundColor Green
        Write-Host "  │  [S] Search this directory" -ForegroundColor Cyan
        Write-Host "  └─ [Q] Back to Main Menu ─────────────────────────────────" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Selection: 3  │  2-5  │  1,3,7  │  A (all)  │  ENTER (refresh)" -ForegroundColor DarkGray
        Write-Host "  Ranges auto-sort (10-3 = same as 3-10). Duplicates removed." -ForegroundColor DarkGray

        $selectedIndices = @()
        $UserWantsToQuit = $false
        $RefreshList     = $false

        while ($true) {
            $range = Read-Host "Enter choice (Numbers, 'A', 'S', 'Q', or ENTER)"
            Write-Log "[USER INPUT] Local Batch Range Selection: '$range'" "INFO" -FileOnly

            if ([string]::IsNullOrWhiteSpace($range)) { Write-Log "Refreshing directory inventory..." "INFO"; $RefreshList = $true; break }
            if ($range -match "^[Qq]$") { Write-Log "Exiting Local Batch Mode..." "INFO"; $UserWantsToQuit = $true; break }

            if ($range -match "^[Ss]$") {
                while ($true) {
                    $searchTerm = Read-Host "Enter FILE NAME or EXTENSION to search for (or 'Q' to exit search)"
                    Write-Log "[USER INPUT] Search Term: '$searchTerm'" "INFO" -FileOnly
                    if ($searchTerm -match "^[Qq]$") { break }
                    if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                        Perform-LocalSearch -SearchPath $Src -SearchTerm $searchTerm
                        Write-Host "`nPress ENTER to return to the batch menu, or type a new search term..." -ForegroundColor DarkGray
                    }
                }
                $RefreshList = $true; break
            }

            $invalidInput = $false; $tempIndices = @()

            if ($range -match "^[Aa]$") {
                if ($Items.Count -eq 0) { Write-Log "Nothing to process." "WARN"; $invalidInput = $true }
                else { $tempIndices = 0..($Items.Count - 1); Write-Log "Selected ALL items (1 through $($Items.Count))." "INFO" }
            } else {
                $parts = $range -split ','
                foreach ($part in $parts) {
                    $part = $part.Trim()
                    if ([string]::IsNullOrWhiteSpace($part)) { continue }
                    if ($part -match "^(\d+)\s*-\s*(\d+)$") {
                        $s = [int]$matches[1] - 1; $e = [int]$matches[2] - 1
                        if ($s -gt $e) { $temp = $s; $s = $e; $e = $temp }
                        if ($s -lt 0 -or $e -ge $Items.Count) { Write-Host "ERROR: Range '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red; $invalidInput = $true; break }
                        $tempIndices += $s..$e
                    } elseif ($part -match "^(\d+)$") {
                        $idx = [int]$matches[1] - 1
                        if ($idx -lt 0 -or $idx -ge $Items.Count) { Write-Host "ERROR: Item '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red; $invalidInput = $true; break }
                        $tempIndices += $idx
                    } else {
                        Write-Host "ERROR: Invalid format '$part'. Use numbers or ranges (e.g., 1, 3-5)." -ForegroundColor Red; $invalidInput = $true; break
                    }
                }
            }

            if ($invalidInput -or $tempIndices.Count -eq 0) { Write-Host "Please try entering your selection again.`n" -ForegroundColor Yellow; continue }
            $selectedIndices = $tempIndices | Select-Object -Unique | Sort-Object
            Write-Log "Selected $($selectedIndices.Count) items for processing." "INFO"
            break
        }

        if ($UserWantsToQuit) { break }
        if ($RefreshList) { continue }

        $action = Read-Host "Type 'C' (Copy), 'M' (Move), 'D' (Delete), or 'V' (Verify) these items"
        Write-Log "[USER INPUT] Local Batch Action Selection: '$action'" "INFO" -FileOnly

        if ($action -notmatch "^[CcMmDdVv]$") { Write-Log "Invalid action. Aborting this chunk." "ERROR"; continue }

        $isMove     = ($action -match "^[Mm]$")
        $isDelete   = ($action -match "^[Dd]$")
        $isVerify   = ($action -match "^[Vv]$")
        $actionWord = if ($isDelete) { "DELETE" } elseif ($isVerify) { "VERIFY" } elseif ($isMove) { "MOVE" } else { "COPY" }

        if ($isDelete) {
            Write-Host ""
            Write-Host "!!! WARNING !!! You selected to permanently DELETE items." -ForegroundColor Red
            $confirmDel = Read-Host "Type 'YES' to confirm this destruction"
            Write-Log "[USER INPUT] Local Batch Delete Confirmation: '$confirmDel'" "INFO" -FileOnly
            if ($confirmDel -cne 'YES') { Write-Log "Deletion aborted by user. Returning to index..." "WARN"; continue }
        }

        if (-not $isDelete -and -not $Dst) { Write-Log "You did not provide a DESTINATION Path! Cannot Copy, Move, or Verify." "ERROR"; continue }

        if (-not $isDelete -and -not $isVerify) {
            Write-Host "Calculating total batch payload for pre-flight storage check..." -ForegroundColor DarkGray
            [long]$TotalBatchBytes = 0
            $batchPairs = @()
            foreach ($i in $selectedIndices) {
                $item    = $Items[$i]
                $ItemDst = Join-Path $Dst $item.Name
                $batchPairs += @{ Src = $item.FullName; Dst = $ItemDst }
                if ($item.PSIsContainer) {
                    $m = Get-LocalDirectoryMetrics -Path $item.FullName
                    if ($m) { $TotalBatchBytes += $m.RawBytes }
                } else { $TotalBatchBytes += $item.Length }
            }
            $canProceed = Invoke-PreflightStorageCheck -DestRoot $Dst -TotalSourceBytes $TotalBatchBytes -BatchItems $batchPairs
            if (-not $canProceed) { Write-Log "Batch operation aborted due to insufficient storage space." "WARN"; continue }
        }

        Write-Log "Initializing Batch $actionWord..." "INFO"

        $global:JobCancelled = $false
        Clear-KeyBuffer

        $batchSw              = [System.Diagnostics.Stopwatch]::StartNew()
        $uiStopwatch          = [System.Diagnostics.Stopwatch]::StartNew()
        $jobStartTime         = Get-Date
        $startTimeStr         = $jobStartTime.ToString("HH:mm:ss")
        $totalItemsToProcess  = $selectedIndices.Count
        $currentItemProcessed = 0

        [long]$GrandSrcBytes = 0; [long]$GrandDstBytes = 0
        $GrandPerfect = 0; $GrandMismatch = 0; $GrandMissing = 0; $GrandOrphaned = 0
        $GrandSuccess = 0; $GrandSkipped  = 0; $GrandFailed  = 0
        [long]$GrandSuccessBytes = 0; [long]$GrandSkippedBytes = 0; [long]$GrandDeletedBytes = 0

        $BatchProcessedPaths = @()
        $BatchFailedList     = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($i in $selectedIndices) {
            $item                = $Items[$i]
            $ItemSrc             = $item.FullName
            $CurrentFolderContext = if ($item.PSIsContainer) { $item.FullName } else { $item.DirectoryName }
            $CurrentFileContext   = $item.Name

            if (-not $isDelete) {
                $ItemDst = Join-Path $Dst $item.Name
                $BatchProcessedPaths += $ItemDst
            }

            if ($global:JobCancelled) { break }
            Check-Interrupt -CurrentFile $CurrentFileContext -CurrentFolder $CurrentFolderContext | Out-Null
            if ($global:JobCancelled) { break }

            $currentItemProcessed++
            $batchStatusText  = "Processing Batch Item $currentItemProcessed of $($totalItemsToProcess): $($item.Name)"
            $percentComplete  = [math]::Min(100, [math]::Round(($currentItemProcessed / $totalItemsToProcess) * 100))

            if ($uiStopwatch.ElapsedMilliseconds -gt 5000 -or $currentItemProcessed -eq $totalItemsToProcess) {
                $elapsed    = (Get-Date) - $jobStartTime
                $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                # v109: ETA + session total — consistent with other batch engines
                $sessionStr = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
                $eta = Get-ETA -Current $currentItemProcessed -Total $totalItemsToProcess -Elapsed $elapsed
                $etaStr = if ($eta) { " | ETA: $eta" } else { "" }
                Write-Progress -Id 0 -Activity "Local Batch Engine [Started: $startTimeStr]" -Status $batchStatusText -PercentComplete $percentComplete -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
                $uiStopwatch.Restart()
            }

            Write-Host ""
            Write-Host "  ── [↑] LOCAL [$currentItemProcessed/$totalItemsToProcess] $($item.Name)" -ForegroundColor Magenta
            Write-Log "LOCAL ITEM $currentItemProcessed/$totalItemsToProcess : $($item.Name)" "INFO" -FileOnly

            if ($isVerify) {
                if ($item.PSIsContainer) {
                    if (-not [System.IO.Directory]::Exists((Get-LongPath $ItemDst))) {
                        Write-Log "MISSING: Folder '$($item.Name)' does not exist in the destination!" "ERROR"; $GrandMissing++
                    } else {
                        $vRes = Perform-LocalVerify -Src $ItemSrc -Dst $ItemDst | Select-Object -Last 1
                        if ($vRes) {
                            $GrandPerfect  += $vRes.Perfect;  $GrandMismatch += $vRes.Mismatch
                            $GrandMissing  += $vRes.MissingInDst; $GrandOrphaned += $vRes.MissingInSrc
                            $GrandSrcBytes += $vRes.SrcBytes; $GrandDstBytes += $vRes.DstBytes
                        }
                    }
                } else {
                    $LongSrc = Get-LongPath $ItemSrc; $LongDst = Get-LongPath $ItemDst
                    if (-not [System.IO.File]::Exists($LongDst)) {
                        Write-Log "MISSING: File '$($item.Name)' does not exist in the destination!" "ERROR"; $GrandMissing++; $GrandSrcBytes += $item.Length
                    } else {
                        $sInfo = [System.IO.FileInfo]::new($LongSrc); $dInfo = [System.IO.FileInfo]::new($LongDst)
                        $sizeStr = Format-FileSize $sInfo.Length; $GrandSrcBytes += $sInfo.Length; $GrandDstBytes += $dInfo.Length
                        if (($sInfo.Length -eq $dInfo.Length) -and ($sInfo.LastWriteTime.ToString() -eq $dInfo.LastWriteTime.ToString())) {
                            Write-Log "Perfectly Synced: $($item.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly; $GrandPerfect++
                        } else {
                            Write-Log "MISMATCH (Size or Date differs): $($item.Name) [Size: $sizeStr]" "ERROR"; $GrandMismatch++
                        }
                    }
                }
            }
            elseif ($isDelete) {
                $ItemDeletedBytes = 0; $ItemDeletedCount = 0; $ItemFailedCount = 0
                $delSw = [System.Diagnostics.Stopwatch]::StartNew()

                if ($item.PSIsContainer) {
                    $FilesToNuke = Get-ChildItem -LiteralPath $ItemSrc -Recurse -File -Force
                    foreach ($f in $FilesToNuke) {
                        $fSize = $f.Length; $sizeStr = Format-FileSize $fSize
                        if (Invoke-ForceDelete -Path (Get-LongPath $f.FullName) -IsFolder:$false) {
                            Write-Log "DELETED FILE: $($f.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                            $ItemDeletedBytes += $fSize; $GrandDeletedBytes += $fSize; $ItemDeletedCount++
                        } else {
                            $ItemFailedCount++; $GrandFailed++
                            $BatchFailedList.Add(@{ File = $f.Name; Path = $f.FullName; Reason = "Locked or Permissions Issue" })
                        }
                    }
                    $DirsToNuke = Get-ChildItem -LiteralPath $ItemSrc -Recurse -Directory -Force | Sort-Object FullName -Descending
                    foreach ($d in $DirsToNuke) {
                        if (Invoke-ForceDelete -Path (Get-LongPath $d.FullName) -IsFolder:$true) {
                            Write-Log "DELETED FOLDER: $($d.Name)" "SUCCESS" -ConsoleOnly; $ItemDeletedCount++
                        } else {
                            $ItemFailedCount++; $GrandFailed++
                            $BatchFailedList.Add(@{ File = $d.Name; Path = $d.FullName; Reason = "Locked or Permissions Issue" })
                        }
                    }
                }

                $LongSrc = Get-LongPath $ItemSrc
                if (Invoke-ForceDelete -Path $LongSrc -IsFolder:$item.PSIsContainer) {
                    $fSize = if (-not $item.PSIsContainer) { $item.Length } else { 0 }
                    if ($fSize -gt 0) { $ItemDeletedBytes += $fSize; $GrandDeletedBytes += $fSize }
                    $ItemDeletedCount++; $GrandSuccess++
                } else {
                    Write-Log "Failed to delete root: $($item.Name)" "ERROR"; $ItemFailedCount++; $GrandFailed++
                    $BatchFailedList.Add(@{ File = $item.Name; Path = $LongSrc; Reason = "Locked or Permissions Issue" })
                }
                $delSw.Stop()
                Write-Host ""
                Write-Log "--- LOCAL DELETE COMPLETE: $($item.Name) ---" "INFO"
                Write-Log "Successfully Deleted: $ItemDeletedCount | Payload: $(Format-FileSize $ItemDeletedBytes)" "SUCCESS"
                Write-Log "Failed: $ItemFailedCount" $(if($ItemFailedCount -gt 0){"ERROR"}else{"SUCCESS"})
                Write-Log "Total Time Elapsed: $([math]::Floor($delSw.Elapsed.TotalHours))h $($delSw.Elapsed.Minutes)m $($delSw.Elapsed.Seconds)s." "INFO"
            }
            else {
                if ($item.PSIsContainer) {
                    if (-not [System.IO.Directory]::Exists((Get-LongPath $ItemDst))) {
                        try { [System.IO.Directory]::CreateDirectory((Get-LongPath $ItemDst)) | Out-Null } catch {}
                    }
                    $tRes = Perform-LocalTransfer -Src $ItemSrc -Dst $ItemDst -IsMove:$isMove -InheritedStartTime $jobStartTime -BatchStatus $batchStatusText -BatchPercent $percentComplete -SkipStorageCheck | Select-Object -Last 1
                    if ($tRes) {
                        $GrandSuccess += $tRes.Success; $GrandSkipped += $tRes.Skipped; $GrandFailed += $tRes.Failed
                        $GrandSuccessBytes += $tRes.SuccessBytes; $GrandSkippedBytes += $tRes.SkippedBytes
                        if ($tRes.FailedItems) { $BatchFailedList.AddRange($tRes.FailedItems) }
                    }
                } else {
                    $LongSrc = Get-LongPath $ItemSrc; $LongDst = Get-LongPath $ItemDst
                    $ShouldTransfer = $true
                    $fSize   = $item.Length; $sizeStr = Format-FileSize $fSize

                    if ([System.IO.File]::Exists($LongDst)) {
                        $sInfo = [System.IO.FileInfo]::new($LongSrc); $dInfo = [System.IO.FileInfo]::new($LongDst)
                        if (($sInfo.Length -eq $dInfo.Length) -and ($sInfo.LastWriteTime.ToString() -eq $dInfo.LastWriteTime.ToString())) {
                            $ShouldTransfer = $false
                            Write-Log "SKIPPED (Already Synced): $($item.Name) [Size: $sizeStr]" "WARN" -ConsoleOnly
                            $GrandSkipped++; $GrandSkippedBytes += $fSize
                            if ($isMove) { Invoke-ForceDelete -Path $LongSrc | Out-Null }
                        } else {
                            if ($isMove) { Invoke-ForceDelete -Path $LongDst | Out-Null }
                        }
                    }

                    if ($ShouldTransfer) {
                        try {
                            if ($isMove) { [System.IO.File]::Move($LongSrc, $LongDst) }
                            else         { [System.IO.File]::Copy($LongSrc, $LongDst, $true) }
                            $oDates = [System.IO.FileInfo]::new($LongSrc)
                            if (-not $oDates.Exists -and $isMove) { $oDates = [System.IO.FileInfo]::new($LongDst) }
                            [System.IO.File]::SetCreationTime($LongDst, $oDates.CreationTime)
                            [System.IO.File]::SetLastWriteTime($LongDst, $oDates.LastWriteTime)
                            Write-Log "${actionWord}: $($item.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                            $GrandSuccess++; $GrandSuccessBytes += $fSize
                        } catch {
                            Write-Log "Failed: $($item.Name) [Reason: $($_.Exception.Message)]" "ERROR"
                            $GrandFailed++
                            $BatchFailedList.Add(@{ File = $item.Name; Path = $ItemSrc; Reason = $_.Exception.Message })
                        }
                    }
                }
            }
        }

        Write-Progress -Id 0 -Activity "Local Batch Engine [Started: $startTimeStr]" -Completed
        $batchSw.Stop()

        if ($BatchFailedList.Count -gt 0) {
            Write-Host ""
            Write-FailedReport -FailedList $BatchFailedList
        }

        if (-not $isDelete -and $BatchProcessedPaths.Count -gt 0) {
            Write-Host "Calculating local directory metrics for this batch..." -ForegroundColor DarkGray
            foreach ($processedPath in $BatchProcessedPaths) {
                if ([System.IO.Directory]::Exists((Get-LongPath $processedPath))) {
                    $metrics = Get-LocalDirectoryMetrics -Path $processedPath
                    if ($metrics) {
                        Write-Host ""
                        Write-Log "=========================================================================" "INFO"
                        Write-Log "   LOCAL DIRECTORY METRICS" "INFO"
                        Write-Log "=========================================================================" "INFO"
                        Write-Log "TARGET NAME:  $(Split-Path $processedPath -Leaf)" "INFO"
                        Write-Log "TOTAL SIZE:   $($metrics.Size)" "INFO"
                        Write-Log "FILE COUNT:   $($metrics.FileCount)" "INFO"
                        Write-Log "FOLDER COUNT: $($metrics.FolderCount)" "INFO"
                    }
                }
            }
        }

        if ($global:JobCancelled) { Write-Log "--- LOCAL BATCH JOB ABORTED BY USER ---" "WARN" }
        else                      { Write-Log "--- BATCH COMPLETION ---" "INFO" }

        Write-Host ""
        $summaryLines = [System.Collections.Generic.List[hashtable]]::new()
        if ($isVerify) {
            $summaryLines.Add(@{ Label = "Perfectly Synced";          Value = "$GrandPerfect";                                    Level = "SUCCESS" })
            $summaryLines.Add(@{ Label = "Mismatched";                Value = "$GrandMismatch";                                   Level = $(if($GrandMismatch -gt 0){"ERROR"}else{"SUCCESS"}) })
            $summaryLines.Add(@{ Label = "Missing in Destination";    Value = "$GrandMissing";                                    Level = $(if($GrandMissing  -gt 0){"ERROR"}else{"SUCCESS"}) })
            $summaryLines.Add(@{ Label = "Orphaned in Destination";   Value = "$GrandOrphaned";                                   Level = $(if($GrandOrphaned -gt 0){"WARN" }else{"SUCCESS"}) })
            $summaryLines.Add(@{ Label = "Source Size Verified";      Value = "$(Format-FileSize $GrandSrcBytes)";                Level = "INFO" })
            $summaryLines.Add(@{ Label = "Destination Size Verified"; Value = "$(Format-FileSize $GrandDstBytes)";                Level = "INFO" })
            $matchLevel = if($GrandSrcBytes -eq $GrandDstBytes -and $GrandSrcBytes -gt 0){"SUCCESS"} elseif($GrandSrcBytes -eq 0 -and $GrandDstBytes -eq 0){"WARN"} else{"ERROR"}
            $matchVal   = if($GrandSrcBytes -eq $GrandDstBytes -and $GrandSrcBytes -gt 0){"YES"} elseif($GrandSrcBytes -eq 0 -and $GrandDstBytes -eq 0){"N/A (Empty)"} else{"NO — MISMATCH"}
            $summaryLines.Add(@{ Label = "Size Match";                Value = $matchVal;                                          Level = $matchLevel })
        } elseif ($isDelete) {
            $summaryLines.Add(@{ Label = "Items Deleted";             Value = "$GrandSuccess  |  $(Format-FileSize $GrandDeletedBytes)"; Level = "SUCCESS" })
            $summaryLines.Add(@{ Label = "Deletion Failures";         Value = "$GrandFailed";                                    Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) })
        } else {
            $summaryLines.Add(@{ Label = "Successfully Processed";    Value = "$GrandSuccess  |  $(Format-FileSize $GrandSuccessBytes)"; Level = "SUCCESS" })
            $summaryLines.Add(@{ Label = "Skipped (already synced)";  Value = "$GrandSkipped  |  $(Format-FileSize $GrandSkippedBytes)"; Level = "INFO" })
            $summaryLines.Add(@{ Label = "Failed to Transfer";        Value = "$GrandFailed";                                    Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) })
        }
        $summaryLines.Add(@{ Label = "Total Time";                    Value = "$([math]::Floor($batchSw.Elapsed.TotalHours))h $($batchSw.Elapsed.Minutes)m $($batchSw.Elapsed.Seconds)s"; Level = "INFO" })
        # v109: average throughput + session total for at-a-glance context
        $avgSpeed = Format-AvgSpeed -Bytes $GrandSuccessBytes -Elapsed $batchSw.Elapsed
        if ($avgSpeed) {
            $summaryLines.Add(@{ Label = "Average Speed";             Value = $avgSpeed; Level = "INFO" })
        }
        # v110: files/min — most useful for batches with many small items
        $filesPerMin = Format-FilesPerMin -Count $GrandSuccess -Elapsed $batchSw.Elapsed
        if ($filesPerMin -and $GrandSuccess -ge 50) {
            $summaryLines.Add(@{ Label = "Throughput";                Value = $filesPerMin; Level = "INFO" })
        }
        if ($global:SessionTransferredBytes -gt 0) {
            $summaryLines.Add(@{ Label = "Session Total";             Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
        }
        Write-SummaryBox -Title "LOCAL BATCH COMPLETE" -Lines $summaryLines.ToArray() -BoxColor "Green"
        # v111: failure-rate alert + ALL distinct errors recap
        Show-OperationDiagnostics -SuccessCount $GrandSuccess -FailedCount $GrandFailed -FailedList $BatchFailedList -ProcessedItems $null
        Write-Host ""
        Write-Host "  Scanning remaining items for next batch..." -ForegroundColor DarkGray
    }
}

# --- FUNCTION: Perform-LocalWipe ---
function Perform-LocalWipe {
    param([string]$TargetFolder)
    Write-Log "--- STARTING LOCAL DIRECTORY WIPE ---" "WARN"
    Write-Log "TARGET FOLDER: $TargetFolder"
    Write-Host "!!! DANGER ZONE !!!" -ForegroundColor Red

    $confirm = Read-Host "Type 'YES' to confirm destruction"
    Write-Log "[USER INPUT] Local Wipe Confirmation: '$confirm'" "INFO" -FileOnly

    if ($confirm -cne 'YES') { Write-Log "Wipe operation aborted by user." "SUCCESS"; return }

    $global:JobCancelled = $false
    Clear-KeyBuffer

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "Executing Local Deletion..." "INFO"

    [long]$WipedBytes = 0
    $Files = Get-ChildItem -LiteralPath $TargetFolder -Recurse -File -Force
    foreach ($f in $Files) {
        $fSize = $f.Length; $sizeStr = Format-FileSize $fSize
        if (Invoke-ForceDelete -Path (Get-LongPath $f.FullName) -IsFolder:$false) {
            $WipedBytes += $fSize
            Write-Log "DELETED FILE: $($f.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
        }
    }

    $Dirs = Get-ChildItem -LiteralPath $TargetFolder -Recurse -Directory -Force | Sort-Object FullName -Descending
    foreach ($d in $Dirs) {
        if (Invoke-ForceDelete -Path (Get-LongPath $d.FullName) -IsFolder:$true) {
            Write-Log "DELETED FOLDER: $($d.Name)" "SUCCESS" -ConsoleOnly
        }
    }

    $LongRoot = Get-LongPath $TargetFolder
    if (Invoke-ForceDelete -Path $LongRoot -IsFolder:$true) {
        $sw.Stop()
        Write-Log "SUCCESSFULLY WIPED LOCAL DIRECTORY: $TargetFolder [Total Payload: $(Format-FileSize $WipedBytes)]" "SUCCESS"
        Write-Log "Total Time Elapsed: $([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s." "SUCCESS"
    } else {
        Write-Log "Failed to wipe root directory (Check if a file is still open)." "ERROR"
    }
}

# ========================================================================================
# ENGINE SUITE D: DIRECT DOWNLOAD TOOLS
# ========================================================================================

# --- FUNCTION: Perform-SafeDownload ---
function Perform-SafeDownload {
    param([object]$Item, [string]$LongDestPath, [string]$TempDir)

    if (-not (Test-Path -LiteralPath $TempDir)) { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }

    $TempFileName = [guid]::NewGuid().ToString() + ".tmp"
    $TempFilePath = Join-Path $TempDir $TempFileName

    $retryCount = 0; $maxRetries = 5

    while ($retryCount -lt $maxRetries) {
        if ($global:JobCancelled) { return @{ Success = $false; Error = "Aborted by User" } }
        try {
            try {
                Get-PnPFile -ServerRelativeUrl $Item.ServerRelativeUrl -Path $TempDir -FileName $TempFileName -AsFile -Force -ErrorAction Stop
            } catch {
                Write-Host "   [REST Failed] Attempting CSOM UniqueId Binary Fallback..." -ForegroundColor DarkGray
                $ctx          = Get-PnPContext
                $file         = $ctx.Web.GetFileById($Item.UniqueId)
                $clientResult = $file.OpenBinaryStream()
                $ctx.ExecuteQuery()
                if ($clientResult.Value -ne $null) {
                    $fileStream = [System.IO.File]::Create($TempFilePath)
                    $clientResult.Value.CopyTo($fileStream)
                    $fileStream.Close()
                    $clientResult.Value.Close()
                } else { throw "CSOM Binary Stream returned null." }
            }

            $LongSrc = Get-LongPath $TempFilePath; $LongDst = Get-LongPath $LongDestPath

            if ("Win32" -as [type]) {
                [Win32]::DeleteFileW($LongDst) | Out-Null
                $moveSuccess = [Win32]::MoveFileW($LongSrc, $LongDst)
                if (-not $moveSuccess) {
                    if ([System.IO.File]::Exists($LongDst)) { [System.IO.File]::Delete($LongDst) }
                    [System.IO.File]::Move($LongSrc, $LongDst)
                }
            } else {
                if ([System.IO.File]::Exists($LongDst)) { [System.IO.File]::Delete($LongDst) }
                [System.IO.File]::Move($LongSrc, $LongDst)
            }
            return @{ Success = $true; Error = $null }
        } catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                if (Test-Path $TempFilePath) { Remove-Item $TempFilePath -Force -ErrorAction SilentlyContinue }
                return @{ Success = $false; Error = $_.Exception.Message }
            }
            $WaitTime = $retryCount * 5
            Write-Host "   [API Throttle] Retrying in $WaitTime seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds $WaitTime
        }
    }
}

# --- FUNCTION: Perform-DirectDownload ---
# PURPOSE: Downloads a SharePoint folder (or single file) to a local path.
# Uses a queue-based BFS (breadth-first scan) so the entire folder tree is traversed
# without hitting PowerShell's recursion depth limits on deeply nested structures.
#
# TEMP BUFFER PATTERN: Each file is downloaded to $BufferDest as a .tmp GUID-named file,
# then atomically promoted to its final destination path using Win32::MoveFileW.
# This prevents partial/corrupted files from ever appearing in the destination folder.
# If the script is interrupted, the temp file is abandoned (it does not corrupt the dest).
#
# SKIP LOGIC: If a file already exists at the destination with the exact same byte size,
# it is skipped. This enables safe resume of interrupted downloads — just re-run with
# the same source and destination paths.
#
# THROTTLE: Perform-SafeDownload wraps every file download in a 5-attempt retry loop
# with exponential backoff. An additional 150ms pause exists between folder scans.
#
# BATCH AWARE: When called from a batch parent, -BatchStatus and -BatchPercent route
# the progress display to the parent progress bar instead of creating a new one.
# The per-download summary box is also suppressed — the batch has its own summary.
function Perform-DirectDownload {
    param(
        [hashtable]$CloudData,
        [string]$LocalDest,
        [string]$BufferDest,
        [string]$BatchStatus    = "",
        [int]$BatchPercent      = -1,
        [datetime]$InheritedStartTime = [datetime]::MinValue
    )

    $FailedList  = [System.Collections.Generic.List[hashtable]]::new()
    $targetName  = Split-Path $LocalDest -Leaf

    if (-not $BatchStatus) {
        $global:JobCancelled = $false
        Clear-KeyBuffer
        Write-Log "--- STARTING DIRECT DOWNLOAD ---" "INFO"
        Write-Log "SOURCE CLOUD URL: $($CloudData.FolderPath)"
        Write-Log "LOCAL DESTINATION: $(Format-PathWithStorageTag $LocalDest)"
        Write-Log "TEMP BUFFER: $BufferDest"
        Show-DestinationFreeSpace -DestPath $LocalDest
    }

    try {
        $SitePath   = [uri]::new($CloudData.SiteUrl).AbsolutePath
        $BaseRelUrl = $CloudData.FolderPath.Substring($SitePath.Length).TrimStart('/')

        $LongLocalRoot = Get-LongPath $LocalDest
        if (-not [System.IO.Directory]::Exists($LongLocalRoot)) { [System.IO.Directory]::CreateDirectory($LongLocalRoot) | Out-Null }

        $singleFileItem = Get-PnPFile -Url $BaseRelUrl -ErrorAction SilentlyContinue

        $sw           = [System.Diagnostics.Stopwatch]::StartNew()
        $SuccessCount = 0; $SkippedCount = 0; $FailedCount = 0
        [long]$SuccessBytes = 0; [long]$SkippedBytes = 0

        $jobStartTime = if ($InheritedStartTime -ne [datetime]::MinValue) { $InheritedStartTime } else { Get-Date }
        $startTimeStr = $jobStartTime.ToString("HH:mm:ss")

        $activityName = if ($BatchStatus) { "Batch Download Engine [Started: $startTimeStr]" } else { "Direct Download Engine [Started: $startTimeStr]" }
        $statusText   = if ($BatchStatus) { $BatchStatus } else { "Target: $targetName" }
        $uiStopwatch  = [System.Diagnostics.Stopwatch]::StartNew()

        if (-not $BatchStatus) {
            Write-Host ""
            Write-Host " HOTKEYS: Press 'P' to Pause or 'C' to Cancel at any time." -ForegroundColor DarkGray
            Write-Host ""
        }

        if ($singleFileItem -and $singleFileItem.Name) {
            $targetName = $singleFileItem.Name
            Write-Host " Target identified as a single file. Bypassing directory scan..." -ForegroundColor DarkGray
            $safeName         = Get-SafeWindowsName $singleFileItem.Name
            $LongItemLocalPath = Get-LongPath (Join-Path $LocalDest $safeName)
            $fSize  = $singleFileItem.Length; $sizeStr = Format-FileSize $fSize
            $ShouldDownload = $true

            if ([System.IO.File]::Exists($LongItemLocalPath)) {
                $LocalFileInfo = [System.IO.FileInfo]::new($LongItemLocalPath)
                if ($LocalFileInfo.Length -eq $singleFileItem.Length) {
                    $ShouldDownload = $false; $SkippedCount++; $SkippedBytes += $fSize
                    Write-Log "SKIPPED (Already Synced): $($singleFileItem.Name) [Size: $sizeStr]" "WARN" -ConsoleOnly
                }
            }

            if ($ShouldDownload) {
                if ($fSize -ge $script:LargeFileBytes) {
                    Write-Host " Downloading: $($singleFileItem.Name) ($sizeStr) — large file, please wait..." -ForegroundColor Yellow
                } else {
                    Write-Host " Downloading: $($singleFileItem.Name)..." -ForegroundColor Cyan
                }
                $dlResult = Perform-SafeDownload -Item $singleFileItem -LongDestPath $LongItemLocalPath -TempDir $BufferDest
                if ($dlResult.Success) {
                    $SuccessCount++; $SuccessBytes += $fSize
                    $global:SessionTransferredBytes += $fSize
                    Write-Log "Downloaded: $($singleFileItem.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                }
                else { $FailedCount++; Write-Log "Download Failed: $($singleFileItem.Name) [Reason: $($dlResult.Error)]" "ERROR"; $FailedList.Add(@{ File = $singleFileItem.Name; Path = $BaseRelUrl; Reason = $dlResult.Error }) }
            }
        } else {
            $Queue = [System.Collections.Generic.Queue[hashtable]]::new()
            $Queue.Enqueue(@{ CloudUrl = $BaseRelUrl; LocalTarget = $LocalDest })

            :QueueLoop while ($Queue.Count -gt 0) {
                if ($global:JobCancelled) { break QueueLoop }
                $current = $Queue.Dequeue()
                Write-Host " Scanning Cloud Sub-Directory: /$($current.CloudUrl)" -ForegroundColor DarkGray
                Start-Sleep -Milliseconds 150

                $items = Get-PnPFolderItem -FolderSiteRelativeUrl $current.CloudUrl -ErrorAction SilentlyContinue
                if (-not $items) { continue }

                foreach ($item in $items) {
                    $CurrentFolderContext = $current.CloudUrl; $CurrentFileContext = $item.Name
                    Check-Interrupt -CurrentFile $CurrentFileContext -CurrentFolder $CurrentFolderContext | Out-Null
                    if ($global:JobCancelled) { break QueueLoop }

                    if ($uiStopwatch.ElapsedMilliseconds -gt 5000) {
                        $elapsed    = (Get-Date) - $jobStartTime
                        $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                        $sessionStr = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
                        # v110: Id 1 for the per-item engine, with -ParentId 0 to nest under the
                        # batch engine bar. When called standalone there is no Id 0 bar present,
                        # so PowerShell renders Id 1 at top level — same visual result as v109.
                        if ($BatchPercent -ge 0) { Write-Progress -Id 1 -ParentId 0 -Activity "  └─ $activityName" -Status $statusText -PercentComplete $BatchPercent -CurrentOperation "Elapsed: $elapsedStr$sessionStr | P=Pause C=Cancel" }
                        else                     { Write-Progress -Id 1            -Activity $activityName        -Status $statusText                                  -CurrentOperation "Elapsed: $elapsedStr$sessionStr | P=Pause C=Cancel" }
                        $uiStopwatch.Restart()
                    }

                    $safeName         = Get-SafeWindowsName $item.Name
                    $ItemLocalPath    = Join-Path $current.LocalTarget $safeName
                    $LongItemLocalPath = Get-LongPath $ItemLocalPath

                    if ($item.GetType().Name -eq "Folder") {
                        if (-not [System.IO.Directory]::Exists($LongItemLocalPath)) {
                            try { [System.IO.Directory]::CreateDirectory($LongItemLocalPath) | Out-Null }
                            catch { Write-Log "Failed to create directory: $LongItemLocalPath" "ERROR" }
                        }
                        $Queue.Enqueue(@{ CloudUrl = "$($current.CloudUrl)/$($item.Name)"; LocalTarget = $ItemLocalPath })
                    } else {
                        $ShouldDownload = $true; $fSize = $item.Length; $sizeStr = Format-FileSize $fSize

                        if ([System.IO.File]::Exists($LongItemLocalPath)) {
                            $LocalFileInfo = [System.IO.FileInfo]::new($LongItemLocalPath)
                            if ($LocalFileInfo.Length -eq $item.Length) {
                                $ShouldDownload = $false; $SkippedCount++; $SkippedBytes += $fSize
                                Write-Log "SKIPPED (Already Synced): $($item.Name) [Size: $sizeStr]" "WARN" -ConsoleOnly
                            }
                        }

                        if ($ShouldDownload) {
                            if ($fSize -ge $script:LargeFileBytes) {
                                Write-Host " Downloading: $($item.Name) ($sizeStr) — large file, please wait..." -ForegroundColor Yellow
                            } else {
                                Write-Host " Downloading: $($item.Name)..." -ForegroundColor Cyan
                            }
                            $dlResult = Perform-SafeDownload -Item $item -LongDestPath $LongItemLocalPath -TempDir $BufferDest
                            if ($dlResult.Success) {
                                $SuccessCount++; $SuccessBytes += $fSize
                                $global:SessionTransferredBytes += $fSize
                                Write-Log "Downloaded: $($item.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                            } else {
                                $FailedCount++
                                Write-Log "Download Failed: $($item.Name) [Reason: $($dlResult.Error)]" "ERROR"
                                $FailedList.Add(@{ File = $item.Name; Path = "$($current.CloudUrl)/$($item.Name)"; Reason = $dlResult.Error })
                            }
                        }
                    }
                }
            }
            if (-not $BatchStatus) { Write-Progress -Id 1 -Activity $activityName -Completed }
        }

        $sw.Stop()
        Write-Host ""

        if ($FailedList.Count -gt 0) {
            Write-FailedReport -FailedList $FailedList
        }

        if (-not $BatchStatus) {
            Write-Host "Calculating final local directory size..." -ForegroundColor DarkGray
            $metrics = Get-LocalDirectoryMetrics -Path $LocalDest
            if ($metrics) {
                Write-Host ""
                Write-Log "=========================================================================" "INFO"
                Write-Log "   LOCAL DIRECTORY METRICS" "INFO"
                Write-Log "=========================================================================" "INFO"
                Write-Log "TARGET NAME:  $(Split-Path $LocalDest -Leaf)" "INFO"
                Write-Log "TOTAL SIZE:   $($metrics.Size)" "INFO"
                Write-Log "FILE COUNT:   $($metrics.FileCount)" "INFO"
                Write-Log "FOLDER COUNT: $($metrics.FolderCount)" "INFO"
            }
        }

        $dlStatusTitle = if ($global:JobCancelled) { "DOWNLOAD ABORTED" } else { "DOWNLOAD COMPLETE" }
        $dlStatusColor = if ($global:JobCancelled) { "Yellow" } else { "Magenta" }
        if (-not $BatchStatus) {
            # v109: switched from inline-conditional @( ... $(if){@{}} ) to a typed List<hashtable>
            # The old pattern injected a $null array element when the condition was false,
            # which Write-SummaryBox would render as an empty row. The list pattern is also
            # easier to extend (avg speed, future fields) without growing the array literal.
            $dlLines = [System.Collections.Generic.List[hashtable]]::new()
            $dlLines.Add(@{ Label = "Target";                  Value = $targetName;                                    Level = "INFO" })
            # v110: storage type — local destination only (source is cloud)
            $dstTag = Get-PathStorageType $LocalDest
            if ($dstTag) {
                $dlLines.Add(@{ Label = "Storage";              Value = "Cloud -> $dstTag";                            Level = "INFO" })
            }
            $dlLines.Add(@{ Label = "Successfully Downloaded"; Value = "$SuccessCount  |  $(Format-FileSize $SuccessBytes)"; Level = "SUCCESS" })
            $dlLines.Add(@{ Label = "Skipped (already existed)"; Value = "$SkippedCount  |  $(Format-FileSize $SkippedBytes)"; Level = "INFO" })
            $dlLines.Add(@{ Label = "Failed";                  Value = "$FailedCount";                                Level = $(if($FailedCount -gt 0){"ERROR"}else{"SUCCESS"}) })
            $dlLines.Add(@{ Label = "Total Time";              Value = "$([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s"; Level = "INFO" })
            $avgSpeed = Format-AvgSpeed -Bytes $SuccessBytes -Elapsed $sw.Elapsed
            if ($avgSpeed) {
                $dlLines.Add(@{ Label = "Average Speed";       Value = $avgSpeed; Level = "INFO" })
            }
            # v110: files/min — helpful when downloading many tiny cloud files
            $filesPerMin = Format-FilesPerMin -Count $SuccessCount -Elapsed $sw.Elapsed
            if ($filesPerMin -and $SuccessCount -ge 50) {
                $dlLines.Add(@{ Label = "Throughput";          Value = $filesPerMin; Level = "INFO" })
            }
            if ($global:SessionTransferredBytes -gt 0) {
                $dlLines.Add(@{ Label = "Session Total";       Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
            }
            Write-SummaryBox -Title $dlStatusTitle -Lines $dlLines.ToArray() -BoxColor $dlStatusColor
            # v111: failure-rate alert + ALL distinct errors recap (FailedList holds File/Path/Reason hashtables;
            # passing it as ProcessedItems too gives the per-extension breakdown a hint of what was attempted)
            Show-OperationDiagnostics -SuccessCount $SuccessCount -FailedCount $FailedCount -FailedList $FailedList -ProcessedItems $null
        } else {
            Write-Log "$dlStatusTitle : $targetName | Downloaded: $SuccessCount | Skipped: $SkippedCount | Failed: $FailedCount | $(Format-FileSize $SuccessBytes)" $(if($global:JobCancelled){"WARN"}else{"INFO"})
        }

        return @{ Success = $SuccessCount; Skipped = $SkippedCount; Failed = $FailedCount; SuccessBytes = $SuccessBytes; SkippedBytes = $SkippedBytes; FailedItems = $FailedList }

    } catch {
        Write-Log "CRITICAL ERROR in Direct Download Engine: $($_.Exception.Message)" "ERROR"
        return @{ Success = 0; Skipped = 0; Failed = 1; SuccessBytes = 0; SkippedBytes = 0 }
    }
}

# --- FUNCTION: Perform-CloudToLocalBatchDownload ---
function Perform-CloudToLocalBatchDownload {
    param([hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest)
    Write-Log "--- ENTERING CLOUD TO LOCAL BATCH DOWNLOAD MODE ---"
    Write-Log "SOURCE CLOUD URL: $($CloudData.FolderPath)"
    Write-Log "LOCAL DESTINATION: $(Format-PathWithStorageTag $LocalDest)"
    Write-Log "TEMP BUFFER: $BufferDest"

    Show-DestinationFreeSpace -DestPath $LocalDest

    try {
        $SitePath       = [uri]::new($CloudData.SiteUrl).AbsolutePath
        $SiteRelativeSrc = $CloudData.FolderPath.Substring($SitePath.Length).TrimStart('/')

        $singleFileItem = Get-PnPFile -Url $SiteRelativeSrc -ErrorAction SilentlyContinue
        if ($singleFileItem -and $singleFileItem.Name) {
            Write-Log "Target is a single file, not a directory. Auto-routing to Direct Download..." "WARN"
            Perform-DirectDownload -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest
            return
        }

        while ($true) {
            Write-PathBanner -Label1 "Source" -Path1 $CloudData.FolderPath -Label2 "Dest" -Path2 $LocalDest -Label3 "Buffer" -Path3 $BufferDest -BoxColor "Magenta"
            Write-Host "Fetching directory inventory from the cloud..." -ForegroundColor DarkGray
            $RawItems = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeSrc -ErrorAction SilentlyContinue

            if (-not $RawItems -or $RawItems.Count -eq 0) {
                Write-Log "Directory '/$SiteRelativeSrc' is completely empty." "WARN"
                break
            }

            $Items = @($RawItems | Sort-Object @{Expression={$_.GetType().Name -eq "Folder"};Descending=$true}, @{Expression={$_.Name};Ascending=$true})

            Write-Host ""
            Write-Host "  ┌─ CLOUD DIRECTORY INDEX (/$SiteRelativeSrc) ────────────────" -ForegroundColor Cyan
            Write-Host "  │  Items = top-level item count │ Modified = date last changed" -ForegroundColor DarkGray

            $MenuAudit = [System.Collections.Generic.List[string]]::new()
            $MenuAudit.Add("--- CLOUD DIRECTORY INDEX (/$SiteRelativeSrc) ---")

            for ($i = 0; $i -lt $Items.Count; $i++) {
                $item = $Items[$i]
                $type = if ($item.GetType().Name -eq "Folder") { "[FOLDER]" } else { "[FILE]  " }
                if ($item.GetType().Name -eq "Folder") {
                    $modDate = if ($item.TimeLastModified) { $item.TimeLastModified.ToString("MM/dd/yyyy") } else { "Unknown" }
                    $iCount  = if ($null -ne $item.ItemCount) { $item.ItemCount } else { "0" }
                    $logStr  = "[$($i + 1)] $type $($item.Name) [Items: $iCount | Modified: $modDate]"
                } else {
                    $logStr = "[$($i + 1)] $type $($item.Name) [Size: $(Format-FileSize $item.Length)]"
                }
                Write-Host " $logStr"
                $MenuAudit.Add($logStr)
            }

            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $LogPayload = ""
            foreach ($line in $MenuAudit) { $LogPayload += "[$ts] [MENU] $line`r`n" }
            [System.IO.File]::AppendAllText($LogFile, $LogPayload)

            Write-Host ""
            Write-Host " [A] Process ALL items in this directory" -ForegroundColor Green
            Write-Host "  │  [S] Search this directory" -ForegroundColor Cyan
            Write-Host "  └─ [Q] Back to Main Menu ─────────────────────────────────" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Selection: 3  │  2-5  │  1,3,7  │  A (all)  │  ENTER (refresh)" -ForegroundColor DarkGray
            Write-Host "  Ranges auto-sort (10-3 = same as 3-10). Duplicates removed." -ForegroundColor DarkGray

            $selectedIndices = @(); $UserWantsToQuit = $false; $RefreshList = $false

            while ($true) {
                $range = Read-Host "Enter choice (Numbers, 'A', 'S', 'Q', or ENTER)"
                Write-Log "[USER INPUT] Batch Range Selection: '$range'" "INFO" -FileOnly

                if ([string]::IsNullOrWhiteSpace($range)) { Write-Log "Refreshing directory inventory..." "INFO"; $RefreshList = $true; break }
                if ($range -match "^[Qq]$") { Write-Log "Exiting Batch Mode..." "INFO"; $UserWantsToQuit = $true; break }

                if ($range -match "^[Ss]$") {
                    while ($true) {
                        $searchTerm = Read-Host "Enter FILE NAME or EXTENSION to search for (or 'Q' to exit search)"
                        Write-Log "[USER INPUT] Search Term: '$searchTerm'" "INFO" -FileOnly
                        if ($searchTerm -match "^[Qq]$") { break }
                        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                            $TargetCloudData = @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $CloudData.FolderPath }
                            Perform-CloudSearch -CloudData $TargetCloudData -SearchTerm $searchTerm -SkipAuth
                            Write-Host "`nPress ENTER to return to the batch menu, or type a new search term..." -ForegroundColor DarkGray
                        }
                    }
                    $RefreshList = $true; break
                }

                $invalidInput = $false; $tempIndices = @()

                if ($range -match "^[Aa]$") {
                    if ($Items.Count -eq 0) { Write-Log "Nothing to process." "WARN"; $invalidInput = $true }
                    else { $tempIndices = 0..($Items.Count - 1); Write-Log "Selected ALL items." "INFO" }
                } else {
                    $parts = $range -split ','
                    foreach ($part in $parts) {
                        $part = $part.Trim()
                        if ([string]::IsNullOrWhiteSpace($part)) { continue }
                        if ($part -match "^(\d+)\s*-\s*(\d+)$") {
                            $s = [int]$matches[1] - 1; $e = [int]$matches[2] - 1
                            if ($s -gt $e) { $temp = $s; $s = $e; $e = $temp }
                            if ($s -lt 0 -or $e -ge $Items.Count) { Write-Host "ERROR: Range '$part' is out of bounds." -ForegroundColor Red; $invalidInput = $true; break }
                            $tempIndices += $s..$e
                        } elseif ($part -match "^(\d+)$") {
                            $idx = [int]$matches[1] - 1
                            if ($idx -lt 0 -or $idx -ge $Items.Count) { Write-Host "ERROR: Item '$part' is out of bounds." -ForegroundColor Red; $invalidInput = $true; break }
                            $tempIndices += $idx
                        } else {
                            Write-Host "ERROR: Invalid format '$part'." -ForegroundColor Red; $invalidInput = $true; break
                        }
                    }
                }

                if ($invalidInput -or $tempIndices.Count -eq 0) { Write-Host "Please try again.`n" -ForegroundColor Yellow; continue }
                $selectedIndices = $tempIndices | Select-Object -Unique | Sort-Object
                break
            }

            if ($UserWantsToQuit) { break }
            if ($RefreshList) { continue }

            Write-Log "Initializing Batch Download..." "INFO"
            $global:JobCancelled = $false
            Clear-KeyBuffer

            Write-Host " HOTKEYS: Press 'P' to Pause or 'C' to Cancel at any time." -ForegroundColor DarkGray

            $batchSw              = [System.Diagnostics.Stopwatch]::StartNew()
            $jobStartTime         = Get-Date
            $startTimeStr         = $jobStartTime.ToString("HH:mm:ss")
            $uiStopwatch          = [System.Diagnostics.Stopwatch]::StartNew()
            $totalItemsToProcess  = $selectedIndices.Count
            $currentItemProcessed = 0

            $GrandSuccess = 0; $GrandSkipped = 0; $GrandFailed = 0
            [long]$GrandSuccessBytes = 0; [long]$GrandSkippedBytes = 0

            $BatchFailedList     = [System.Collections.Generic.List[hashtable]]::new()
            $BatchProcessedPaths = @()

            foreach ($i in $selectedIndices) {
                $item                = $Items[$i]
                $CurrentFolderContext = $SiteRelativeSrc
                $CurrentFileContext   = $item.Name

                if ($global:JobCancelled) { break }
                Check-Interrupt -CurrentFile $CurrentFileContext -CurrentFolder $CurrentFolderContext | Out-Null
                if ($global:JobCancelled) { break }

                $currentItemProcessed++

                if ($uiStopwatch.ElapsedMilliseconds -gt 5000 -or $currentItemProcessed -eq $totalItemsToProcess) {
                    $percentComplete = [math]::Min(100, [math]::Round(($currentItemProcessed / $totalItemsToProcess) * 100))
                    $elapsed    = (Get-Date) - $jobStartTime
                    $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                    # v109: consistency — ETA + session in CurrentOperation across all batch engines
                    $sessionStr = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
                    $eta = Get-ETA -Current $currentItemProcessed -Total $totalItemsToProcess -Elapsed $elapsed
                    $etaStr = if ($eta) { " | ETA: $eta" } else { "" }
                    Write-Progress -Id 0 -Activity "Batch Download Engine [Started: $startTimeStr]" -Status "Processing Item $currentItemProcessed of $($totalItemsToProcess): $($item.Name)" -PercentComplete $percentComplete -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
                    $uiStopwatch.Restart()
                }

                Write-Host ""
                Write-Host "  ── [↓] DOWNLOAD [$currentItemProcessed/$totalItemsToProcess] $($item.Name)" -ForegroundColor Magenta
                Write-Log "DOWNLOAD ITEM $currentItemProcessed/$totalItemsToProcess : $($item.Name)" "INFO" -FileOnly

                $safeName          = Get-SafeWindowsName $item.Name
                $ItemLocalPath     = Join-Path $LocalDest $safeName
                $LongItemLocalPath = Get-LongPath $ItemLocalPath
                $BatchProcessedPaths += $ItemLocalPath

                if ($item.GetType().Name -eq "Folder") {
                    $ItemSrcUrl      = "$($CloudData.FolderPath)/$($item.Name)" -replace '//', '/'
                    $TargetCloudData = @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $ItemSrcUrl }
                    $batchStatusText = "Processing Batch Item $currentItemProcessed of $($totalItemsToProcess): $($item.Name)"
                    $percentComplete = [math]::Min(100, [math]::Round(($currentItemProcessed / $totalItemsToProcess) * 100))

                    $tRes = Perform-DirectDownload -CloudData $TargetCloudData -LocalDest $ItemLocalPath -BufferDest $BufferDest -BatchStatus $batchStatusText -BatchPercent $percentComplete -InheritedStartTime $jobStartTime | Select-Object -Last 1
                    if ($tRes) {
                        $GrandSuccess += $tRes.Success; $GrandSkipped += $tRes.Skipped; $GrandFailed += $tRes.Failed
                        $GrandSuccessBytes += $tRes.SuccessBytes; $GrandSkippedBytes += $tRes.SkippedBytes
                        if ($tRes.FailedItems) { $BatchFailedList.AddRange($tRes.FailedItems) }
                    }
                } else {
                    $ShouldDownload = $true; $fSize = $item.Length; $sizeStr = Format-FileSize $fSize

                    if ([System.IO.File]::Exists($LongItemLocalPath)) {
                        $LocalFileInfo = [System.IO.FileInfo]::new($LongItemLocalPath)
                        if ($LocalFileInfo.Length -eq $item.Length) {
                            $ShouldDownload = $false
                            Write-Log "SKIPPED (Already Synced): $($item.Name) [Size: $sizeStr]" "WARN" -ConsoleOnly
                            $GrandSkipped++; $GrandSkippedBytes += $fSize
                        }
                    }

                    if ($ShouldDownload) {
                        if (-not [System.IO.Directory]::Exists((Get-LongPath $LocalDest))) {
                            [System.IO.Directory]::CreateDirectory((Get-LongPath $LocalDest)) | Out-Null
                        }
                        Write-Host " Downloading: $($item.Name)..." -ForegroundColor Cyan
                        $dlResult = Perform-SafeDownload -Item $item -LongDestPath $LongItemLocalPath -TempDir $BufferDest

                        if ($dlResult.Success) {
                            Write-Log "Downloaded: $($item.Name) [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                            $GrandSuccess++; $GrandSuccessBytes += $fSize
                        } else {
                            Write-Log "Download Failed: $($item.Name) [Reason: $($dlResult.Error)]" "ERROR"
                            $BatchFailedList.Add(@{ File = $item.Name; Path = "$SiteRelativeSrc/$($item.Name)"; Reason = $dlResult.Error })
                            $GrandFailed++
                        }
                    }
                }
            }

            Write-Progress -Id 0 -Activity "Batch Download Engine [Started: $startTimeStr]" -Completed
            $batchSw.Stop()

            if ($BatchFailedList.Count -gt 0) {
                Write-Host ""
                Write-FailedReport -FailedList $BatchFailedList
            }

            Write-Host "Calculating local directory metrics for this batch..." -ForegroundColor DarkGray
            foreach ($processedPath in $BatchProcessedPaths) {
                if ([System.IO.Directory]::Exists((Get-LongPath $processedPath))) {
                    $metrics = Get-LocalDirectoryMetrics -Path $processedPath
                    if ($metrics) {
                        Write-Host ""
                        Write-Log "=========================================================================" "INFO"
                        Write-Log "   LOCAL DIRECTORY METRICS" "INFO"
                        Write-Log "=========================================================================" "INFO"
                        Write-Log "TARGET NAME:  $(Split-Path $processedPath -Leaf)" "INFO"
                        Write-Log "TOTAL SIZE:   $($metrics.Size)" "INFO"
                        Write-Log "FILE COUNT:   $($metrics.FileCount)" "INFO"
                        Write-Log "FOLDER COUNT: $($metrics.FolderCount)" "INFO"
                    }
                }
            }

            if ($global:JobCancelled) { Write-Log "--- BATCH JOB ABORTED BY USER ---" "WARN" }
            else                      { Write-Log "--- BATCH COMPLETION ---" "INFO" }

            Write-Host ""
            # v109: list-based summary so we can conditionally append avg speed and session total
            $bdLines = [System.Collections.Generic.List[hashtable]]::new()
            # v110: storage tag — local destination
            $dstTag = Get-PathStorageType $LocalDest
            if ($dstTag) {
                $bdLines.Add(@{ Label = "Storage";              Value = "Cloud -> $dstTag";                            Level = "INFO" })
            }
            $bdLines.Add(@{ Label = "Successfully Downloaded";  Value = "$GrandSuccess  |  $(Format-FileSize $GrandSuccessBytes)"; Level = "SUCCESS" })
            $bdLines.Add(@{ Label = "Skipped (already synced)"; Value = "$GrandSkipped  |  $(Format-FileSize $GrandSkippedBytes)"; Level = "INFO" })
            $bdLines.Add(@{ Label = "Failed to Download";       Value = "$GrandFailed";                                            Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) })
            $bdLines.Add(@{ Label = "Total Time";               Value = "$([math]::Floor($batchSw.Elapsed.TotalHours))h $($batchSw.Elapsed.Minutes)m $($batchSw.Elapsed.Seconds)s"; Level = "INFO" })
            $avgSpeed = Format-AvgSpeed -Bytes $GrandSuccessBytes -Elapsed $batchSw.Elapsed
            if ($avgSpeed) {
                $bdLines.Add(@{ Label = "Average Speed";        Value = $avgSpeed; Level = "INFO" })
            }
            # v110: files/min — most useful for cloud batches with many small items
            $filesPerMin = Format-FilesPerMin -Count $GrandSuccess -Elapsed $batchSw.Elapsed
            if ($filesPerMin -and $GrandSuccess -ge 50) {
                $bdLines.Add(@{ Label = "Throughput";           Value = $filesPerMin; Level = "INFO" })
            }
            if ($global:SessionTransferredBytes -gt 0) {
                $bdLines.Add(@{ Label = "Session Total";        Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
            }
            Write-SummaryBox -Title "BATCH DOWNLOAD COMPLETE" -Lines $bdLines.ToArray() -BoxColor "Magenta"
            # v111: failure-rate alert + ALL distinct errors recap
            Show-OperationDiagnostics -SuccessCount $GrandSuccess -FailedCount $GrandFailed -FailedList $BatchFailedList -ProcessedItems $null
            Write-Host ""
            Write-Host "  Scanning directory again to refresh the menu..." -ForegroundColor DarkGray
        }
    } catch {
        Write-Log "Cloud Batch Encountered an Error: $($_.Exception.Message)" "ERROR"
    }
}


# ========================================================================================
# ENGINE SUITE F: ARCHIVE OPERATIONS — Download + Verify + Delete
# ========================================================================================
#
# PURPOSE: Provides a safe, auditable "archive and clean" workflow for leavers or
# any scenario where SharePoint storage needs to be reclaimed after data is archived.
#
# WORKFLOW:
#   1. Download the cloud folder/file to a local hard drive (using the proven download engine).
#   2. Verify by comparing the cloud inventory (file count + total bytes) against what was
#      downloaded. Name-sanitisation differences are safe because we compare counts and
#      byte totals, not filenames.
#   3. Only if verification passes AND the user types 'YES' → delete from SharePoint.
#      Deletion ALWAYS goes to the SharePoint Recycle Bin (recoverable ~93 days).
#
# SAFETY GATES:
#   - Any download failure → halt, do NOT delete.
#   - File count or byte total mismatch → halt, do NOT delete.
#   - Explicit 'YES' confirmation required at the delete prompt.
#   - Job cancel (P/C hotkeys) stops without touching the cloud source.

# --- FUNCTION: Invoke-ArchiveVerify ---
# PURPOSE: Mathematically proves that a download captured everything in the cloud source.
#
# WHY COUNTS + BYTES, NOT FILENAMES:
#   The download engine applies Get-SafeWindowsName to every file before saving it locally.
#   This replaces characters that are legal in SharePoint but illegal on Windows (like : ? *)
#   with dashes. So "Report: Final.docx" becomes "Report- Final.docx" on disk.
#   Comparing filenames after this substitution would produce false mismatches.
#   Instead this function compares TOTAL FILE COUNT and TOTAL BYTES only.
#   If every file made it locally the counts and byte totals will match regardless
#   of any character substitutions applied during the save.
#
# SINGLE FILE FALLBACK:
#   Get-CloudInventory uses Get-PnPFolderItem, which only works on folder URLs.
#   If the source URL points to a single file (not a folder), Get-CloudInventory
#   returns 0 files. The fallback uses Get-PnPFile to read the file's size directly.
# APPROACH: Compares (Success + Skipped) file count and (SuccessBytes + SkippedBytes)
# against the cloud inventory totals. This avoids filename-comparison pitfalls from
# Get-SafeWindowsName character substitutions — if every cloud byte made it locally,
# the counts and totals will match regardless of character substitutions in names.
# Always called with an active PnP connection (SkipAuth implied).
#
# SINGLE FILE FALLBACK: If Get-CloudInventory returns 0 files (source was a single
# file URL, not a folder), falls back to Get-PnPFile to get the cloud file size.
function Invoke-ArchiveVerify {
    param([hashtable]$CloudData, [hashtable]$DlResult)

    $SitePath    = [uri]::new($CloudData.SiteUrl).AbsolutePath
    $BaseSiteRel = $CloudData.FolderPath.Substring($SitePath.Length).TrimStart('/')

    $CloudFileCount    = 0
    [long]$CloudTotalBytes = 0

    # Primary: full recursive inventory for folder sources
    try {
        $CloudInv = Get-CloudInventory -SiteUrl $CloudData.SiteUrl -FolderPath $CloudData.FolderPath -Tag "Archive Verify" -SkipAuth
        $CloudFileCount = $CloudInv.Files.Count
        foreach ($f in $CloudInv.Files) { $CloudTotalBytes += $f.Length }
    } catch { }

    # Fallback: single file source (Get-CloudInventory uses Get-PnPFolderItem which fails on file URLs)
    if ($CloudFileCount -eq 0) {
        try {
            $singleFile = Get-PnPFile -Url $BaseSiteRel -ErrorAction SilentlyContinue
            if ($singleFile -and $singleFile.Length) {
                $CloudFileCount  = 1
                $CloudTotalBytes = $singleFile.Length
            }
        } catch { }
    }

    [long]$LocalTotalBytes = $DlResult.SuccessBytes + $DlResult.SkippedBytes
    $LocalFileCount        = $DlResult.Success + $DlResult.Skipped

    # All three conditions must hold before deletion is offered
    $Passed = ($DlResult.Failed -eq 0) -and
              ($LocalFileCount -eq $CloudFileCount) -and
              ($LocalTotalBytes -eq $CloudTotalBytes)

    return @{
        Passed          = $Passed
        CloudFileCount  = $CloudFileCount
        CloudTotalBytes = $CloudTotalBytes
        LocalFileCount  = $LocalFileCount
        LocalTotalBytes = $LocalTotalBytes
        Failed          = $DlResult.Failed
    }
}

# --- FUNCTION: Perform-ArchiveDownload ---
# PURPOSE: Full three-step guided archive workflow for a single cloud folder.
# Called from the [17] menu handler. Auth is already established by the menu handler.
# Step 1: Download (Perform-DirectDownload).
# Step 2: Verify (Invoke-ArchiveVerify — cloud inventory vs download stats).
# Step 3: Offer deletion — shows results, requires explicit 'YES' to proceed.
#         Deletion uses Remove-PnPFolder with -Recycle (SharePoint Recycle Bin only).
function Perform-ArchiveDownload {
    param([hashtable]$CloudData, [string]$LocalDest, [string]$BufferDest)

    $archiveSw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- STARTING ARCHIVE DOWNLOAD (Download + Verify + Delete) ---" "INFO"
    Write-Log "CLOUD SOURCE: $($CloudData.FolderPath)"
    Write-Log "LOCAL DEST:   $LocalDest"
    Write-Log "TEMP BUFFER:  $BufferDest"

    # ── STEP 1: DOWNLOAD ─────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "  ║  STEP 1 OF 3  ──  DOWNLOADING FROM SHAREPOINT           ║" -ForegroundColor Magenta
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""

    $dlResult = Perform-DirectDownload -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest

    if ($global:JobCancelled) {
        Write-Host ""
        Write-Host "  Archive aborted by user. Cloud source has NOT been touched." -ForegroundColor Yellow
        Write-Log "Archive Download: Aborted by user. Cloud source untouched." "WARN"
        return
    }

    if ($dlResult.Failed -gt 0) {
        Write-Host ""
        Write-Host "  ┌─ ARCHIVE HALTED ─── DOWNLOAD ERRORS DETECTED ──────────" -ForegroundColor Red
        Write-Host "  │  $($dlResult.Failed) file(s) failed to download." -ForegroundColor Red
        Write-Host "  │  The cloud source has NOT been modified or deleted." -ForegroundColor Red
        Write-Host "  │  Resolve errors and re-run the archive before deleting." -ForegroundColor Red
        Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Red
        Write-Log "Archive Download: HALTED — $($dlResult.Failed) download failure(s). Cloud NOT deleted." "ERROR"
        return
    }

    # ── STEP 2: VERIFY ───────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "  ║  STEP 2 OF 3  ──  VERIFYING DOWNLOAD INTEGRITY          ║" -ForegroundColor Yellow
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Cross-referencing cloud file count and byte total against local..." -ForegroundColor DarkGray

    $vResult = Invoke-ArchiveVerify -CloudData $CloudData -DlResult $dlResult

    $countLabel  = if ($vResult.CloudFileCount -eq $vResult.LocalFileCount) { "YES" } else { "NO — MISMATCH ($($vResult.LocalFileCount) local vs $($vResult.CloudFileCount) cloud)" }
    $bytesLabel  = if ($vResult.CloudTotalBytes -eq $vResult.LocalTotalBytes) { "YES" } else { "NO — MISMATCH ($(Format-FileSize $vResult.LocalTotalBytes) local vs $(Format-FileSize $vResult.CloudTotalBytes) cloud)" }

    Write-SummaryBox -Title "ARCHIVE VERIFICATION RESULTS" -Lines @(
        @{ Label = "Cloud Source";      Value = "$($vResult.CloudFileCount) files  |  $(Format-FileSize $vResult.CloudTotalBytes)"; Level = "INFO"    }
        @{ Label = "Local Destination"; Value = "$($vResult.LocalFileCount) files  |  $(Format-FileSize $vResult.LocalTotalBytes)"; Level = "INFO"    }
        @{ Label = "File Count Match";  Value = $countLabel; Level = if($vResult.CloudFileCount -eq $vResult.LocalFileCount){"SUCCESS"}else{"ERROR"}  }
        @{ Label = "Total Bytes Match"; Value = $bytesLabel; Level = if($vResult.CloudTotalBytes -eq $vResult.LocalTotalBytes){"SUCCESS"}else{"ERROR"} }
        @{ Label = "Download Failures"; Value = "$($vResult.Failed)"; Level = if($vResult.Failed -gt 0){"ERROR"}else{"SUCCESS"}                       }
        @{ Label = "Verification";      Value = if($vResult.Passed){"PASSED — all files confirmed locally"}else{"FAILED — do NOT delete cloud source"};
                                        Level = if($vResult.Passed){"SUCCESS"}else{"ERROR"}                                                           }
    ) -BoxColor (if ($vResult.Passed) { "Green" } else { "Red" })

    if (-not $vResult.Passed) {
        Write-Host ""
        Write-Host "  ┌─ ARCHIVE HALTED ─── VERIFICATION FAILED ───────────────" -ForegroundColor Red
        Write-Host "  │  File count or byte total does not match the cloud source." -ForegroundColor Red
        Write-Host "  │  The cloud source has NOT been modified or deleted." -ForegroundColor Red
        Write-Host "  │  Re-run the archive download to retry any missing files." -ForegroundColor Red
        Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Red
        Write-Log "Archive Download: HALTED — verification failed. Cloud NOT deleted." "ERROR"
        return
    }

    # ── STEP 3: DELETE ───────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║  STEP 3 OF 3  ──  DELETE CLOUD SOURCE                   ║" -ForegroundColor Cyan
    Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "  ║  Verification PASSED — proceeding with automatic delete. ║" -ForegroundColor Green
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Cloud folder : $($CloudData.FolderPath)" -ForegroundColor Yellow
    Write-Host "  Local backup : $LocalDest" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Deletion sends the folder to the SharePoint RECYCLE BIN." -ForegroundColor DarkGray
    Write-Host "  Recoverable for approximately 93 days. Deleting now..." -ForegroundColor DarkGray
    Write-Host ""

    Write-Log "Archive Download: Auto-deleting cloud source (verification passed): $($CloudData.FolderPath)" "WARN"

    try {
        $ParentPath = (Split-Path $CloudData.FolderPath -Parent) -replace '\\', '/'
        $FolderName = Split-Path $CloudData.FolderPath -Leaf
        Invoke-PnPWithRetry -OperationName "DELETE (ARCHIVE) $FolderName" -ScriptBlock {
            Remove-PnPFolder -Name $FolderName -Folder $ParentPath -Force -Recycle -ErrorAction Stop
        } | Out-Null
        $archiveSw.Stop()
        Write-Log "Archive Download: Cloud source deleted (Recycle Bin). '$($CloudData.FolderPath)'" "SUCCESS"
        Write-Host ""
        Write-SummaryBox -Title "ARCHIVE DOWNLOAD COMPLETE" -Lines @(
            @{ Label = "Source";          Value = $FolderName;                                                                Level = "INFO"    }
            @{ Label = "Files Archived";  Value = "$($vResult.LocalFileCount) files  |  $(Format-FileSize $vResult.LocalTotalBytes)"; Level = "SUCCESS" }
            @{ Label = "Verified";        Value = "PASSED — counts and bytes match";                                         Level = "SUCCESS" }
            @{ Label = "Cloud Deleted";   Value = "YES — sent to SharePoint Recycle Bin";                                   Level = "SUCCESS" }
            @{ Label = "Recoverable for"; Value = "~93 days from the SharePoint Recycle Bin";                               Level = "INFO"    }
            @{ Label = "Total Time";      Value = "$([math]::Floor($archiveSw.Elapsed.TotalHours))h $($archiveSw.Elapsed.Minutes)m $($archiveSw.Elapsed.Seconds)s"; Level = "INFO" }
        ) -BoxColor "Green"
    } catch {
        Write-Log "Archive Download: FAILED to delete cloud source! Reason: $($_.Exception.Message)" "ERROR"
        Write-Host ""
        Write-Host "  ERROR: The deletion command failed. The cloud source is still present." -ForegroundColor Red
        Write-Host "  Reason: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Your local archive is complete and verified." -ForegroundColor Green
        Write-Host "  You can manually delete the folder from SharePoint." -ForegroundColor Yellow
    }
}

# --- FUNCTION: Perform-BatchArchiveDownload ---
# PURPOSE: Shows a cloud directory as a numbered index. User selects items to archive.
# For each selected item: Download → Verify → Auto-delete if verified.
# Verification is the only safety gate — no prompts needed. Failed items are never
# deleted regardless.
# Items that pass verification are deleted automatically (no per-item prompts).
# Items that FAIL verification are skipped with no deletion — safe by default.
# [C] cancel hotkey still works mid-batch to stop after the current item finishes.
# Auth is already established by the [18] menu handler; this function uses SkipAuth.
function Perform-BatchArchiveDownload {
    param(
        [hashtable]$CloudData,
        [string]$LocalDest,
        [string]$BufferDest,
        # When set, all top-level items are selected automatically and the interactive
        # index menu is never shown. Used by Perform-LeaverOneDriveArchive so that
        # [20] and [21] run fully unattended after the user types YES.
        [switch]$AutoSelectAll
    )
    Write-Log "--- ENTERING BATCH ARCHIVE DOWNLOAD MODE ---"
    Write-Log "CLOUD SOURCE: $($CloudData.FolderPath)"
    Write-Log "LOCAL DEST:   $LocalDest"

    Show-DestinationFreeSpace -DestPath $LocalDest

    try {
        $SitePath        = [uri]::new($CloudData.SiteUrl).AbsolutePath
        $SiteRelativeSrc = $CloudData.FolderPath.Substring($SitePath.Length).TrimStart('/')

        # Auto-route single file to the direct archive workflow
        $singleFileItem = Get-PnPFile -Url $SiteRelativeSrc -ErrorAction SilentlyContinue
        if ($singleFileItem -and $singleFileItem.Name) {
            Write-Log "Target is a single file. Auto-routing to Archive Download..." "WARN"
            Perform-ArchiveDownload -CloudData $CloudData -LocalDest $LocalDest -BufferDest $BufferDest
            return
        }

        while ($true) {
            Write-PathBanner -Label1 "Source" -Path1 $CloudData.FolderPath -Label2 "Dest" -Path2 $LocalDest -Label3 "Buffer" -Path3 $BufferDest -BoxColor "Yellow"
            Write-Host "Fetching directory inventory from the cloud..." -ForegroundColor DarkGray
            $RawItems = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeSrc -ErrorAction SilentlyContinue

            if (-not $RawItems -or $RawItems.Count -eq 0) {
                Write-Log "Cloud directory is completely empty. Exiting Batch Archive..." "WARN"
                break
            }

            $Items = @($RawItems | Sort-Object @{Expression={$_.GetType().Name -eq "Folder"};Descending=$true}, @{Expression={$_.Name};Ascending=$true})

            # ── AUTO-SELECT ALL (unattended mode — skip menu entirely) ────────
            if ($AutoSelectAll) {
                $selectedIndices = @(1..$Items.Count)
                Write-Log "AutoSelectAll: Selected all $($Items.Count) item(s) in /$SiteRelativeSrc." "INFO"
                Write-Host "  Auto-selecting all $($Items.Count) item(s) in /$SiteRelativeSrc..." -ForegroundColor DarkGray
            } else {
                # ── INTERACTIVE MENU (normal [18]/[20] usage) ─────────────────
                Write-Host ""
                Write-Host "  ┌─ BATCH ARCHIVE INDEX (/$SiteRelativeSrc) ─────────────" -ForegroundColor Yellow
                Write-Host "  │  ARCHIVE MODE: each item is downloaded, verified, then offered for deletion." -ForegroundColor DarkGray
                Write-Host "  │  Items = top-level item count │ Modified = date last changed" -ForegroundColor DarkGray

                $MenuAudit = [System.Collections.Generic.List[string]]::new()
                $MenuAudit.Add("--- BATCH ARCHIVE INDEX (/$SiteRelativeSrc) ---")

                for ($i = 0; $i -lt $Items.Count; $i++) {
                    $item = $Items[$i]
                    $type = if ($item.GetType().Name -eq "Folder") { "[FOLDER]" } else { "[FILE]  " }
                    if ($item.GetType().Name -eq "Folder") {
                        $modDate = if ($item.TimeLastModified) { $item.TimeLastModified.ToString("MM/dd/yyyy") } else { "Unknown" }
                        $iCount  = if ($null -ne $item.ItemCount) { $item.ItemCount } else { "0" }
                        $logStr  = "[$($i + 1)] $type $($item.Name) [Items: $iCount | Modified: $modDate]"
                    } else {
                        $logStr = "[$($i + 1)] $type $($item.Name) [Size: $(Format-FileSize $item.Length)]"
                    }
                    Write-Host " $logStr"
                    $MenuAudit.Add($logStr)
                }

                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $LogPayload = ""
                foreach ($line in $MenuAudit) { $LogPayload += "[$ts] [MENU] $line`r`n" }
                [System.IO.File]::AppendAllText($LogFile, $LogPayload)

                Write-Host ""
                Write-Host " [A] Archive ALL items in this directory" -ForegroundColor Green
                Write-Host "  │  [S] Search this directory" -ForegroundColor Cyan
                Write-Host "  └─ [Q] Back to Main Menu ─────────────────────────────────" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Selection: 3  │  2-5  │  1,3,7  │  A (all)  │  ENTER (refresh)" -ForegroundColor DarkGray
                Write-Host "  Ranges auto-sort (10-3 = same as 3-10). Duplicates removed." -ForegroundColor DarkGray
            }

            $selectedIndices = @(); $UserWantsToQuit = $false; $RefreshList = $false

            if ($AutoSelectAll) {
                # Unattended mode: select all items using 0-based indices (matching $Items[$i] usage below)
                $selectedIndices = @(0..($Items.Count - 1))
                Write-Log "Selected all $($Items.Count) item(s) for archiving (AutoSelectAll)." "INFO"
            } else {
                # Interactive mode: keep the inner while loop so ENTER=refresh and S=search
                # can loop back without leaving the batch archive screen
                while ($true) {
                    $range = Read-Host "Enter choice (Numbers, 'A', 'S', 'Q', or ENTER)"
                    Write-Log "[USER INPUT] Batch Archive Range Selection: '$range'" "INFO" -FileOnly

                    if ([string]::IsNullOrWhiteSpace($range)) { Write-Log "Refreshing directory inventory..." "INFO"; $RefreshList = $true; break }
                    if ($range -match "^[Qq]$") { Write-Log "Exiting Batch Archive Mode..." "INFO"; $UserWantsToQuit = $true; break }

                    if ($range -match "^[Ss]$") {
                        while ($true) {
                            $searchTerm = Read-Host "Enter FILE NAME or EXTENSION to search for (or 'Q' to exit search)"
                            Write-Log "[USER INPUT] Search Term: '$searchTerm'" "INFO" -FileOnly
                            if ($searchTerm -match "^[Qq]$") { break }
                            if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                                Perform-CloudSearch -CloudData @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $CloudData.FolderPath } -SearchTerm $searchTerm -SkipAuth
                                Write-Host "`nPress ENTER to return to the archive menu..." -ForegroundColor DarkGray
                            }
                        }
                        $RefreshList = $true; break
                    }

                    $invalidInput = $false; $tempIndices = @()

                    if ($range -match "^[Aa]$") {
                        $tempIndices = 0..($Items.Count - 1)
                        Write-Log "Selected ALL items (1 through $($Items.Count))." "INFO"
                    } else {
                        $parts = $range -split ','
                        foreach ($part in $parts) {
                            $part = $part.Trim()
                            if ([string]::IsNullOrWhiteSpace($part)) { continue }
                            if ($part -match "^(\d+)\s*-\s*(\d+)$") {
                                $s = [int]$matches[1] - 1; $e = [int]$matches[2] - 1
                                if ($s -gt $e) { $temp = $s; $s = $e; $e = $temp }
                                if ($s -lt 0 -or $e -ge $Items.Count) { Write-Host "ERROR: Range '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red; $invalidInput = $true; break }
                                $tempIndices += $s..$e
                            } elseif ($part -match "^(\d+)$") {
                                $idx = [int]$matches[1] - 1
                                if ($idx -lt 0 -or $idx -ge $Items.Count) { Write-Host "ERROR: Item '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red; $invalidInput = $true; break }
                                $tempIndices += $idx
                            } else {
                                Write-Host "ERROR: Invalid format '$part'. Use numbers or ranges (e.g., 1, 3-5)." -ForegroundColor Red; $invalidInput = $true; break
                            }
                        }
                    }

                    if ($invalidInput -or $tempIndices.Count -eq 0) { Write-Host "Please try entering your selection again.`n" -ForegroundColor Yellow; continue }
                    $selectedIndices = $tempIndices | Select-Object -Unique | Sort-Object
                    Write-Log "Selected $($selectedIndices.Count) items for archiving." "INFO"
                    break
                }
            }

            if ($UserWantsToQuit) { break }
            if ($RefreshList) { continue }

            Write-Log "Initializing Batch Archive..." "INFO"
            $global:JobCancelled = $false
            Clear-KeyBuffer

            # No upfront prompt — verification IS the safety gate.
            # Items that pass verification are deleted automatically. Items that fail are never touched.
            Write-Host ""
            Write-Host "  ┌─ BATCH ARCHIVE STARTING ───────────────────────────────" -ForegroundColor Yellow
            Write-Host "  │  $($selectedIndices.Count) item(s) queued: Download → Verify → Auto-delete on pass." -ForegroundColor White
            Write-Host "  │  Items that FAIL verification will NOT be deleted." -ForegroundColor DarkGray
            Write-Host "  │  Press [C] at any time to stop cleanly after the current item." -ForegroundColor DarkGray
            Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Yellow
            Write-Host ""
            Write-Log "Batch Archive: Starting $($selectedIndices.Count) item(s) — auto-delete on verify pass." "INFO"

            $batchSw             = [System.Diagnostics.Stopwatch]::StartNew()
            $totalItemsToProcess = $selectedIndices.Count
            $currentItemProc     = 0

            # Grand batch counters
            $GrandArchived    = 0; [long]$GrandArchivedBytes = 0
            $GrandDeleted     = 0
            $GrandDlFailed    = 0; $GrandVerifyFailed = 0
            $GrandCheckedOut  = @()   # filenames blocked by checkout — shown in summary

            foreach ($i in $selectedIndices) {
                if ($global:JobCancelled) { break }

                $item = $Items[$i]
                $currentItemProc++

                Write-Host ""
                Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor Yellow
                Write-Host "  ARCHIVE ITEM $currentItemProc of $totalItemsToProcess : $($item.Name)" -ForegroundColor Yellow
                Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor Yellow
                Write-Log "ARCHIVE ITEM $currentItemProc/$totalItemsToProcess : $($item.Name)" "INFO" -FileOnly

                $safeName      = Get-SafeWindowsName $item.Name
                $ItemLocalPath = Join-Path $LocalDest $safeName
                $ItemDlResult  = $null

                # ── DOWNLOAD ─────────────────────────────────────────────────
                Write-Host "  [1/3] Downloading '$($item.Name)'..." -ForegroundColor Magenta

                if ($item.GetType().Name -eq "Folder") {
                    $ItemSrcUrl    = "$($CloudData.FolderPath)/$($item.Name)" -replace '//', '/'
                    $ItemCloudData = @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $ItemSrcUrl }
                    $batchStatus   = "Archive $currentItemProc/$totalItemsToProcess : $($item.Name)"
                    $pct           = [math]::Min(100, [math]::Round(($currentItemProc / $totalItemsToProcess) * 100))
                    $ItemDlResult  = Perform-DirectDownload -CloudData $ItemCloudData -LocalDest $ItemLocalPath -BufferDest $BufferDest -BatchStatus $batchStatus -BatchPercent $pct | Select-Object -Last 1
                } else {
                    # Single file item
                    $ItemSrcUrl    = "$($CloudData.FolderPath)/$($item.Name)" -replace '//', '/'
                    $ItemCloudData = @{ SiteUrl = $CloudData.SiteUrl; FolderPath = $ItemSrcUrl }
                    $LongItemLocal = Get-LongPath $ItemLocalPath
                    $fSize = $item.Length; $sizeStr = Format-FileSize $fSize
                    $ShouldDownload = $true

                    if ([System.IO.File]::Exists($LongItemLocal)) {
                        $li = [System.IO.FileInfo]::new($LongItemLocal)
                        if ($li.Length -eq $fSize) {
                            $ShouldDownload = $false
                            Write-Log "SKIPPED (Already Synced): $($item.Name) [Size: $sizeStr]" "WARN" -ConsoleOnly
                        }
                    }

                    if ($ShouldDownload) {
                        if (-not [System.IO.Directory]::Exists((Get-LongPath $LocalDest))) { [System.IO.Directory]::CreateDirectory((Get-LongPath $LocalDest)) | Out-Null }
                        Write-Host "  Downloading: $($item.Name)..." -ForegroundColor Cyan
                        $dlRes = Perform-SafeDownload -Item $item -LongDestPath $LongItemLocal -TempDir $BufferDest
                        $ItemDlResult = if ($dlRes.Success) {
                            @{ Success=1; Skipped=0; Failed=0; SuccessBytes=$fSize; SkippedBytes=0 }
                        } else {
                            @{ Success=0; Skipped=0; Failed=1; SuccessBytes=0; SkippedBytes=0 }
                        }
                    } else {
                        $ItemDlResult = @{ Success=0; Skipped=1; Failed=0; SuccessBytes=0; SkippedBytes=$fSize }
                    }
                }

                if ($global:JobCancelled) {
                    Write-Log "Batch Archive: Aborted by user during download of '$($item.Name)'." "WARN"; break
                }

                if ($ItemDlResult.Failed -gt 0) {
                    Write-Host "  [!] Download had $($ItemDlResult.Failed) failure(s) — skipping verify and delete." -ForegroundColor Red
                    Write-Log "Batch Archive: '$($item.Name)' had download failure(s). NOT deleting." "ERROR"
                    $GrandDlFailed++; continue
                }

                # ── VERIFY ───────────────────────────────────────────────────
                Write-Host "  [2/3] Verifying '$($item.Name)'..." -ForegroundColor Yellow

                $ItemVResult = $null
                if ($item.GetType().Name -eq "Folder") {
                    $ItemVResult = Invoke-ArchiveVerify -CloudData $ItemCloudData -DlResult $ItemDlResult
                } else {
                    # Single file verification: passed if downloaded/skipped and sizes match
                    $localOk = ($ItemDlResult.Failed -eq 0) -and (($ItemDlResult.Success + $ItemDlResult.Skipped) -eq 1)
                    $ItemVResult = @{
                        Passed          = $localOk
                        CloudFileCount  = 1
                        CloudTotalBytes = $item.Length
                        LocalFileCount  = $ItemDlResult.Success + $ItemDlResult.Skipped
                        LocalTotalBytes = $ItemDlResult.SuccessBytes + $ItemDlResult.SkippedBytes
                        Failed          = $ItemDlResult.Failed
                    }
                }

                if (-not $ItemVResult.Passed) {
                    Write-Host "  [!] Verification FAILED — skipping delete for '$($item.Name)'." -ForegroundColor Red
                    Write-Log "Batch Archive: '$($item.Name)' failed verification. NOT deleting." "ERROR"
                    $GrandVerifyFailed++; continue
                }

                $GrandArchived++
                $GrandArchivedBytes += $ItemVResult.LocalTotalBytes
                Write-Host "  [✔] Verified: $($ItemVResult.LocalFileCount) files | $(Format-FileSize $ItemVResult.LocalTotalBytes)" -ForegroundColor Green
                Write-Log "Batch Archive: '$($item.Name)' verified OK — auto-deleting." "INFO"

                # ── AUTO DELETE (verification passed — no per-item prompt) ─────
                Write-Host "  [3/3] Deleting '$($item.Name)' from SharePoint..." -ForegroundColor Cyan
                # Patterns that are permanent errors — never worth retrying
                $delFastFail = @(
                    "(?i)Forms directory of a document library",
                    "(?i)checked out for editing"
                )
                try {
                    if ($item.GetType().Name -eq "Folder") {
                        Invoke-PnPWithRetry -OperationName "DELETE (BATCH ARCHIVE) $($item.Name)" -FastFailPatterns $delFastFail -ScriptBlock {
                            Remove-PnPFolder -Name $item.Name -Folder $CloudData.FolderPath -Force -Recycle -ErrorAction Stop | Out-Null
                        } | Out-Null
                    } else {
                        $srvRelUrl = "$($CloudData.FolderPath)/$($item.Name)" -replace '//', '/'
                        Invoke-PnPWithRetry -OperationName "DELETE (BATCH ARCHIVE) $($item.Name)" -FastFailPatterns $delFastFail -ScriptBlock {
                            Remove-PnPFile -ServerRelativeUrl $srvRelUrl -Force -Recycle -ErrorAction Stop | Out-Null
                        } | Out-Null
                    }
                    Write-Host "  [✔] '$($item.Name)' sent to SharePoint Recycle Bin." -ForegroundColor Green
                    Write-Log "Batch Archive: '$($item.Name)' deleted (Recycle Bin)." "SUCCESS"
                    $GrandDeleted++
                } catch {
                    $delErr = $_.Exception.Message
                    # ── Case 1: SharePoint /Forms system folder — cannot ever be deleted ──
                    if ($delErr -match "(?i)Forms directory of a document library") {
                        Write-Host "  [i] '$($item.Name)' is a SharePoint system folder and cannot be deleted." -ForegroundColor DarkGray
                        Write-Host "      This is expected and safe to ignore — it contains no user data." -ForegroundColor DarkGray
                        Write-Log "Batch Archive: '$($item.Name)' skipped deletion — SharePoint system folder (expected, not an error)." "INFO"

                    # ── Case 2: A file inside the folder is checked out by someone ──────
                    } elseif ($delErr -match "(?i)checked out for editing") {
                        # Extract the checked-out filename from the error URL if present
                        $coFile = if ($delErr -match "(?i)The file `"([^`"]+)`"") { Split-Path $matches[1] -Leaf } else { "unknown file" }
                        # Extract the user identity if present
                        $coUser = if ($delErr -match "by\s+i:0#\.f\|membership\|([^\.\s]+\.[^\s]+)") { $matches[1] }
                                  elseif ($delErr -match "by\s+(.{5,60})\.$") { $matches[1].Trim() }
                                  else { "an unknown user" }
                        Write-Host ""
                        Write-Host "  [⚠] DELETION BLOCKED — file checked out:" -ForegroundColor Yellow
                        Write-Host "      File  : $coFile" -ForegroundColor Yellow
                        Write-Host "      By    : $coUser" -ForegroundColor Yellow
                        Write-Host "      Action: Go to SharePoint / OneDrive, discard or check in that file," -ForegroundColor Yellow
                        Write-Host "              then re-run the archive on '$($item.Name)' to delete it." -ForegroundColor Yellow
                        Write-Host ""
                        Write-Log "Batch Archive: '$($item.Name)' NOT deleted — '$coFile' is checked out by $coUser. Check it in then re-run." "WARN"
                        $GrandCheckedOut += "$($item.Name) (checked out: $coFile by $coUser)"

                    # ── Case 3: Any other unexpected error ───────────────────────────────
                    } else {
                        Write-Host "  [!] Deletion failed: $delErr" -ForegroundColor Red
                        Write-Log "Batch Archive: Failed to delete '$($item.Name)'. Reason: $delErr" "ERROR"
                    }
                }
            }

            $batchSw.Stop()

            if ($global:JobCancelled) { Write-Log "--- BATCH ARCHIVE JOB ABORTED BY USER ---" "WARN" }
            else                      { Write-Log "--- BATCH ARCHIVE COMPLETION ---" "INFO" }

            Write-Host ""
            # Build summary lines — include checkout row only when relevant
            $summaryLines = @(
                @{ Label = "Items Fully Archived";  Value = "$GrandArchived  |  $(Format-FileSize $GrandArchivedBytes)"; Level = "SUCCESS" }
                @{ Label = "Cloud Copies Deleted";  Value = "$GrandDeleted";      Level = if($GrandDeleted -eq $GrandArchived -and $GrandArchived -gt 0){"SUCCESS"}else{"WARN"} }
            )
            if ($GrandCheckedOut.Count -gt 0) {
                $summaryLines += @{ Label = "Blocked by Checkout"; Value = "$($GrandCheckedOut.Count) item(s) — see below"; Level = "WARN" }
            }
            $summaryLines += @{ Label = "Failed Verification";   Value = "$GrandVerifyFailed"; Level = if($GrandVerifyFailed -gt 0){"ERROR"}else{"SUCCESS"} }
            $summaryLines += @{ Label = "Failed Download";       Value = "$GrandDlFailed";     Level = if($GrandDlFailed -gt 0){"ERROR"}else{"SUCCESS"} }
            $summaryLines += @{ Label = "Total Time";            Value = "$([math]::Floor($batchSw.Elapsed.TotalHours))h $($batchSw.Elapsed.Minutes)m $($batchSw.Elapsed.Seconds)s"; Level = "INFO" }
            Write-SummaryBox -Title "BATCH ARCHIVE SUMMARY" -Lines $summaryLines -BoxColor "Yellow"

            # If any items were blocked by checkout, list them with action guidance
            if ($GrandCheckedOut.Count -gt 0) {
                Write-Host ""
                Write-Host "  ┌─ DELETION BLOCKED — CHECKED OUT FILES ────────────────────" -ForegroundColor Yellow
                Write-Host "  │  The following item(s) were downloaded and verified OK but" -ForegroundColor Yellow
                Write-Host "  │  could NOT be deleted because a file inside is checked out." -ForegroundColor Yellow
                Write-Host "  │" -ForegroundColor Yellow
                foreach ($coEntry in $GrandCheckedOut) {
                    Write-Host "  │  • $coEntry" -ForegroundColor Yellow
                }
                Write-Host "  │" -ForegroundColor Yellow
                Write-Host "  │  TO FIX: Open SharePoint / OneDrive for that user and either" -ForegroundColor White
                Write-Host "  │    1. Check in (or discard the checkout on) the blocked file." -ForegroundColor White
                Write-Host "  │    2. Re-run the archive — downloaded files will be skipped" -ForegroundColor White
                Write-Host "  │       (smart resume) and only deletion will be retried." -ForegroundColor White
                Write-Host "  └────────────────────────────────────────────────────────────" -ForegroundColor Yellow
            }
            Write-Host ""

            # In AutoSelectAll mode (unattended leaver archive) we are done — exit the
            # outer loop immediately. There is no menu to re-show.
            if ($AutoSelectAll) { break }

            Write-Host "  Scanning directory again to refresh the menu..." -ForegroundColor DarkGray
        }
    } catch {
        Write-Log "Batch Archive Encountered an Error: $($_.Exception.Message)" "ERROR"
    }
}


# ========================================================================================
# LEAVER RECYCLE BIN URL HELPERS (v100)
# ========================================================================================

# --- FUNCTION: Get-LeaverRecycleBinUrl ---
# PURPOSE: Converts any leaver OneDrive URL (including /_layouts/15/onedrive.aspx variants,
# ?viewid= URLs, or a plain personal site root) into the SharePoint Recycle Bin URL for
# that personal site.
#
# HOW IT WORKS:
#   SharePoint personal site Recycle Bins are always at:
#     https://<tenant>-my.sharepoint.com/personal/<username>/_layouts/15/RecycleBin.aspx
#   This function extracts the personal site root (https://…/personal/username) from any
#   leaver URL format and appends the fixed RecycleBin.aspx path.
#
# RETURNS: The recycle bin URL string, or empty string if the URL cannot be parsed.
function Get-LeaverRecycleBinUrl {
    param([string]$LeaverUrl)
    # Match the personal site root up to and including the username segment
    $m = [regex]::Match($LeaverUrl, '(?i)(https://[^/]+/personal/[^/?# ]+)')
    if ($m.Success) {
        return "$($m.Groups[1].Value.TrimEnd('/'))/_layouts/15/RecycleBin.aspx"
    }
    return ""
}

# --- FUNCTION: Write-RecycleBinLinks ---
# PURPOSE: Displays a formatted recycle bin link box after archive operations complete.
# Shows one or more recycle bin URLs so the admin can open them in a browser to review
# and permanently empty the deleted items.
# After displaying the links, offers C=Copy / O=Open in browser / B=Both.
#
# PARAMETERS:
#   Links : Array of hashtables — each entry has Label (string) and Url (string).
#           Single-leaver calls pass a one-element array; multi-leaver passes N.
function Write-RecycleBinLinks {
    param([array]$Links)
    if (-not $Links -or $Links.Count -eq 0) { return }

    Write-Host ""
    Write-Host "  ┌─ RECYCLE BIN LINK$(if($Links.Count -gt 1){'S'}) ────────────────────────────────" -ForegroundColor Cyan
    Write-Host "  │  Deleted items are in the SharePoint Recycle Bin (~93 days)." -ForegroundColor DarkGray
    Write-Host "  │  Open the link$(if($Links.Count -gt 1){'s'}) below in your browser to review and empty:" -ForegroundColor DarkGray
    Write-Host "  │" -ForegroundColor Cyan

    foreach ($link in $Links) {
        if ($Links.Count -gt 1) {
            $lbl = if ($link.Label.Length -gt 24) { $link.Label.Substring(0,21) + "..." } else { $link.Label }
            Write-Host "  │  $lbl" -ForegroundColor White
            Write-Host "  │    $($link.Url)" -ForegroundColor Cyan
        } else {
            Write-Host "  │  $($link.Url)" -ForegroundColor Cyan
        }
    }
    Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Log "Recycle Bin Links: $(($Links | ForEach-Object { $_.Url }) -join ', ')" "INFO"

    # ── C / O / B POST-ACTION PROMPT ─────────────────────────────────────────
    Write-Host ""
    Write-Host "  [C] Copy to clipboard   [O] Open in browser   [B] Both   ENTER = skip" -ForegroundColor DarkGray
    $rbAction = Read-Host "  Action"
    $rbAction = $rbAction.Trim().ToUpper()

    if ($rbAction -in @('C','B')) {
        try {
            $clipText = ($Links | ForEach-Object { $_.Url }) -join "`r`n"
            Set-Clipboard -Value $clipText -ErrorAction Stop
            Write-Host "  [✔] Copied $($Links.Count) link$(if($Links.Count -ne 1){'s'}) to clipboard." -ForegroundColor Green
            Write-Log "Recycle Bin Links: Copied $($Links.Count) link(s) to clipboard." "INFO"
        } catch {
            Write-Host "  [!] Clipboard copy failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($rbAction -in @('O','B')) {
        Write-Host "  Opening $($Links.Count) link$(if($Links.Count -ne 1){'s'}) in your browser..." -ForegroundColor DarkGray
        foreach ($link in $Links) {
            try {
                Start-Process $link.Url -ErrorAction Stop
                Write-Log "Recycle Bin Links: Opened in browser: $($link.Url)" "INFO"
                if ($Links.Count -gt 1) { Start-Sleep -Milliseconds 800 }
            } catch {
                Write-Host "  [!] Could not open: $($link.Url)" -ForegroundColor Yellow
            }
        }
        Write-Host "  [✔] Done — check your browser." -ForegroundColor Green
    }
}

# ========================================================================================
# ENGINE SUITE F: LEAVER ONEDRIVE ARCHIVE TOOLS (new in v100)
# ========================================================================================

# --- FUNCTION: Ensure-SPOModule ---
# PURPOSE: Checks for and auto-installs the Microsoft.Online.SharePoint.PowerShell module
# (the official SharePoint Online Management Shell). This module provides Set-SPOUser,
# which is the only way to grant Site Collection Admin rights on a personal OneDrive
# site without logging in interactively as the site owner.
#
# WHY THIS MODULE?
#   PnP.PowerShell alone cannot grant Site Collection Admin rights. The SPO Management
#   Shell has direct tenant-level cmdlets (Connect-SPOService, Set-SPOUser) that operate
#   against the SharePoint Admin Center — a separate admin endpoint that any SharePoint
#   Admin can access regardless of whether the target site owner's account is active.
#
# COMPATIBILITY: Works in Windows PowerShell 5.1. The module is from the PSGallery.
function Ensure-SPOModule {
    $spoMod = Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell' | Select-Object -First 1
    if (-not $spoMod) {
        Write-Log "SharePoint Online Management Shell not found. Installing..." "WARN"
        try {
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
            }
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Install-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Microsoft.Online.SharePoint.PowerShell installed successfully." "SUCCESS"
        } catch {
            Write-Log "CRITICAL: Failed to install SharePoint Online Management Shell. $($_.Exception.Message)" "ERROR"
            throw "SPO module install failed: $($_.Exception.Message)"
        }
    }
    try {
        Import-Module 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Log "Microsoft.Online.SharePoint.PowerShell loaded." "INFO"
    } catch {
        Write-Log "CRITICAL: Could not import SharePoint Online Management Shell. $($_.Exception.Message)" "ERROR"
        throw "SPO module import failed: $($_.Exception.Message)"
    }
}

# --- FUNCTION: ConvertTo-OneDriveSiteUrl ---
# PURPOSE: Converts a leaver's UPN (email address) and tenant name into the
# predictable personal site URL that SharePoint Online always uses.
#
# HOW IT WORKS:
#   SharePoint Online personal site URLs follow a fixed pattern:
#     https://<tenant>-my.sharepoint.com/personal/<safe_upn>
#   The <safe_upn> is the UPN with ALL dots, @ signs and hyphens replaced by underscores,
#   then lowercased. For example:
#     john.smith@zellis.com  →  john_smith_zellis_com
#     anna-maria.jones@moorepay.co.uk  →  anna_maria_jones_moorepay_co_uk
#
# PARAMETERS:
#   TenantName : The short tenant name (e.g. "zellis" from zellis.sharepoint.com)
#   Upn        : The leaver's full email / User Principal Name
#
# RETURNS: Full personal site URL string, e.g. https://zellis-my.sharepoint.com/personal/john_smith_zellis_com
function ConvertTo-OneDriveSiteUrl {
    param(
        [string]$TenantName,
        [string]$Upn
    )
    # Lowercase, then replace @, dots, and hyphens with underscores
    $safe = $Upn.ToLower().Trim()
    $safe = $safe -replace '[@\.\-]', '_'
    # Collapse any double-underscores that can arise from consecutive special chars
    while ($safe -match '__') { $safe = $safe -replace '__', '_' }
    $safe = $safe.Trim('_')
    return "https://$($TenantName.ToLower())-my.sharepoint.com/personal/$safe"
}

# --- FUNCTION: Grant-LeaverSiteAdmin ---
# PURPOSE: Grants the specified admin account Site Collection Admin rights on a leaver's
# personal OneDrive site using Set-SPOUser via the SharePoint Online Management Shell.
#
# WHY SET-SPOUSER INSTEAD OF PNP?
#   PnP.PowerShell's Add-PnPSiteCollectionAdmin requires you to ALREADY be a Site Collection
#   Admin on the target site. Set-SPOUser operates at the TENANT admin level — it can grant
#   rights on ANY site in the tenant, including personal OneDrives of disabled/unlicensed
#   accounts, without needing existing admin rights on that specific site.
#
# PROPAGATION DELAY: After Set-SPOUser succeeds, SharePoint Online's permission system
#   takes a few seconds to propagate the change. A 5-second sleep is included before
#   returning so the subsequent PnP connection reliably succeeds.
#
# PARAMETERS:
#   SpoAdminUrl  : Your SPO Admin Center URL (e.g. https://zellis-admin.sharepoint.com)
#   LeaverSiteUrl: The leaver's personal site URL (e.g. https://zellis-my.sharepoint.com/personal/...)
#   AdminUpn     : The UPN of the admin account to grant (usually the person running the script)
#
# RETURNS: $true if grant succeeded, $false on error.
function Grant-LeaverSiteAdmin {
    param(
        [string]$SpoAdminUrl,
        [string]$LeaverSiteUrl,
        [string]$AdminUpn
    )
    try {
        Write-Host "  [SPO] Connecting to SharePoint Admin Center..." -ForegroundColor DarkGray
        Write-Host "        $SpoAdminUrl" -ForegroundColor DarkGray
        Write-Host "  A browser sign-in window will appear — sign in with your admin account." -ForegroundColor Yellow

        $WarningPreference = "SilentlyContinue"
        Connect-SPOService -Url $SpoAdminUrl -ErrorAction Stop
        $WarningPreference = "Continue"

        Write-Host "  [SPO] Granting Site Collection Admin on: $LeaverSiteUrl" -ForegroundColor DarkGray
        Set-SPOUser -Site $LeaverSiteUrl -LoginName $AdminUpn -IsSiteCollectionAdmin $true -ErrorAction Stop
        Write-Log "Granted SiteCollectionAdmin: '$AdminUpn' on '$LeaverSiteUrl'" "SUCCESS"
        Write-Host "  [✔] Admin rights granted. Waiting 5s for permissions to propagate..." -ForegroundColor Green
        Start-Sleep -Seconds 5
        return $true
    } catch {
        Write-Log "Grant-LeaverSiteAdmin FAILED: $($_.Exception.Message)" "ERROR"
        Write-Host "  [!] Could not grant admin rights: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# --- FUNCTION: Revoke-LeaverSiteAdmin ---
# PURPOSE: Removes the temporary Site Collection Admin grant from the leaver's site.
# Called after the archive job completes (or fails) so permissions are cleaned up.
# Safe to call even if Grant-LeaverSiteAdmin was not successfully called — errors are
# caught and logged rather than thrown.
#
# PARAMETERS:
#   LeaverSiteUrl: The leaver's personal site URL
#   AdminUpn     : The UPN of the admin account to revoke (same value as Grant-)
function Revoke-LeaverSiteAdmin {
    param(
        [string]$LeaverSiteUrl,
        [string]$AdminUpn
    )
    try {
        Write-Host "  [SPO] Revoking temporary admin rights for '$AdminUpn'..." -ForegroundColor DarkGray
        Set-SPOUser -Site $LeaverSiteUrl -LoginName $AdminUpn -IsSiteCollectionAdmin $false -ErrorAction Stop
        Write-Log "Revoked SiteCollectionAdmin: '$AdminUpn' on '$LeaverSiteUrl'" "SUCCESS"
        Write-Host "  [✔] Temporary admin access revoked." -ForegroundColor Green
    } catch {
        Write-Log "Revoke-LeaverSiteAdmin WARNING: $($_.Exception.Message)" "WARN"
        Write-Host "  [!] Revoke returned an error (may already be revoked): $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "      Check the SPO Admin Center to verify access has been removed." -ForegroundColor Yellow
    } finally {
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
    }
}

# --- FUNCTION: Perform-LeaverOneDriveArchive ---
# PURPOSE: The full leaver OneDrive archive engine. After admin access is granted,
# this function connects to the leaver's personal site via PnP, enumerates all
# Document Libraries (not just /Documents), and presents them as a batch archive
# index identical to option [18].
#
# DOCUMENT LIBRARY ENUMERATION:
#   A personal OneDrive site can contain multiple document libraries:
#     - Documents  (the main OneDrive library — always present)
#     - Apps       (used by various Microsoft 365 integrations)
#     - Attachments (Teams meeting attachments etc.)
#   Enumerating all libraries ensures nothing is missed.
#
# ARCHIVE FLOW PER LIBRARY:
#   Each library uses Perform-BatchArchiveDownload — the same engine as option [18].
#   Download → Verify → Auto-delete on pass. Safety gates are identical.
#
# PARAMETERS:
#   LeaverSiteUrl: Full personal site URL (https://tenant-my.sharepoint.com/personal/...)
#   LocalDest    : Local folder to archive into — one subfolder is created per library
#   BufferDest   : Temp buffer path for downloads
# NOTE: Auth/admin grant is handled by the [20] menu handler before calling this function.
function Perform-LeaverOneDriveArchive {
    # PURPOSE: Enumerates all Document Libraries on an already-connected leaver's personal
    # OneDrive site (PnP connection must be established BEFORE calling this function),
    # then runs the full Batch Archive workflow (Download → Verify → Auto-delete) on each.
    #
    # DESIGN NOTE: This function intentionally does NOT handle authentication or admin
    # grants. The caller (option [20] menu handler) is responsible for establishing the
    # PnP connection — either via a standard UseWebLogin or after a Set-SPOUser escalation.
    # This keeps the archive engine reusable regardless of how access was obtained.
    #
    # RETURNS: A hashtable with keys Success (int), Failed (int) for the caller to log.
    param(
        [string]$LeaverSiteUrl,   # Full personal site URL — used for cloud data context
        [string]$LocalDest,       # Root local folder — one subfolder created per library
        [string]$BufferDest       # Temp buffer path for downloads
    )

    $archiveSw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- LEAVER ARCHIVE ENGINE: Enumerating and archiving libraries ---" "INFO"
    Write-Log "LEAVER SITE: $LeaverSiteUrl | LOCAL DEST: $LocalDest" "INFO"

    $GrandArchived = 0; $GrandFailed = 0
    $LibResults    = [System.Collections.Generic.List[hashtable]]::new()

    try {
        # ── ENUMERATE DOCUMENT LIBRARIES ──────────────────────────────────────
        # BaseType 1 = DocumentLibrary. Hidden=false excludes system lists (Style Library,
        # Site Assets, Preservation Hold Library etc.). We include empty libraries so the
        # user can see all libraries even if some are already empty.
        #
        # SYSTEM LIBRARY DENYLIST:
        # Personal OneDrive sites contain several non-user document libraries created by
        # SharePoint infrastructure. These have Hidden=false but contain no user data and
        # cannot be meaningfully archived:
        #   Form Templates (BaseTemplate 115) — InfoPath form template storage
        #   Site Pages     (BaseTemplate 850) — SharePoint page library
        #   Style Library  (BaseTemplate 116) — CSS/XSL/image assets
        #   Converted Forms                   — Legacy InfoPath migration artifacts
        # We exclude these by name. The user's actual OneDrive data lives in "Documents"
        # (BaseTemplate 700) and any other standard document libraries they created.
        Write-Host ""
        Write-Host "  Scanning for Document Libraries on the leaver's site..." -ForegroundColor DarkGray

        $SystemLibNames = @(
            'Form Templates', 'Site Pages', 'Style Library', 'Converted Forms',
            'Preservation Hold Library', 'Master Page Gallery', 'Composed Looks',
            'Theme Gallery', 'Site Collection Documents', 'Site Collection Images'
        )

        $AllLists = Get-PnPList -ErrorAction Stop
        $DocLibs  = @($AllLists | Where-Object {
            $_.BaseType -eq 1 -and
            -not $_.Hidden -and
            $_.Title -notin $SystemLibNames
        } | Sort-Object Title)

        if ($DocLibs.Count -eq 0) {
            Write-Log "No Document Libraries found on leaver site. Site may be empty." "WARN"
            Write-Host "  [!] No Document Libraries found. The site may already be empty." -ForegroundColor Yellow
            return @{ Success = 0; Failed = 0 }
        }

        Write-Host ""
        Write-Host "  ┌─ DOCUMENT LIBRARIES FOUND ─────────────────────────────" -ForegroundColor DarkYellow
        Write-Host "  │  Site: $($LeaverSiteUrl -replace 'https://','')" -ForegroundColor DarkGray
        Write-Host "  │" -ForegroundColor DarkYellow
        for ($li = 0; $li -lt $DocLibs.Count; $li++) {
            $lib     = $DocLibs[$li]
            $itemStr = "$($lib.ItemCount) item$(if($lib.ItemCount -ne 1){'s'})"
            Write-Host "  │  [$($li+1)] $($lib.Title.PadRight(32)) $itemStr" -ForegroundColor DarkYellow
        }
        Write-Host "  │" -ForegroundColor DarkYellow
        Write-Host "  │  Each library → subfolder under: $LocalDest" -ForegroundColor DarkGray
        Write-Host "  │" -ForegroundColor DarkYellow
        Write-Host "  │  TIP: Inside each library you can choose which items to archive." -ForegroundColor DarkGray
        Write-Host "  │  Press [A] to archive everything, or select specific items." -ForegroundColor DarkGray
        Write-Host "  │  Press [Q] when done with a library to move to the next one." -ForegroundColor DarkGray
        Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkYellow
        Write-Host ""

        # ── PROCESS EACH LIBRARY ──────────────────────────────────────────────
        for ($li = 0; $li -lt $DocLibs.Count; $li++) {
            $lib         = $DocLibs[$li]
            $libTitle    = $lib.Title
            $libRelUrl   = $lib.RootFolder.ServerRelativeUrl  # e.g. /personal/john_smith_co_uk/Documents

            # Documents library content goes directly into $LocalDest (no extra Documents\ subfolder).
            # Any other library (user-created or non-standard) gets its own named subfolder to
            # avoid name collisions. This keeps the archive structure clean and intuitive:
            #   dest\Desktop\      (from Documents/Desktop)
            #   dest\Meetings\     (from Documents/Meetings)
            # rather than:
            #   dest\Documents\Desktop\
            #   dest\Documents\Meetings\
            $libLocalDst = if ($libTitle -eq 'Documents') {
                $LocalDest
            } else {
                Join-Path $LocalDest (Get-SafeWindowsName $libTitle)
            }

            Write-Host ""
            Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor DarkYellow
            Write-Host "  LIBRARY $($li+1) of $($DocLibs.Count) : $libTitle" -ForegroundColor DarkYellow
            Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor DarkYellow

            if (-not (Test-Path -LiteralPath $libLocalDst)) {
                New-Item -ItemType Directory -Path $libLocalDst -Force | Out-Null
            }

            $LibCloudData = @{ SiteUrl = $LeaverSiteUrl; FolderPath = $libRelUrl }
            Write-Log "Leaver Archive: Processing library '$libTitle' → '$libLocalDst'" "INFO"

            try {
                Perform-BatchArchiveDownload -CloudData $LibCloudData -LocalDest $libLocalDst -BufferDest $BufferDest -AutoSelectAll
                $LibResults.Add(@{ Label = $libTitle; Status = "PROCESSED"; Level = "SUCCESS" })
                $GrandArchived++
            } catch {
                $errMsg = $_.Exception.Message
                Write-Log "Leaver Archive: Library '$libTitle' error: $errMsg" "ERROR"
                Write-Host "  [!] Library '$libTitle' error: $errMsg" -ForegroundColor Red
                $LibResults.Add(@{ Label = $libTitle; Status = "ERROR — $errMsg"; Level = "ERROR" })
                $GrandFailed++
            }
        }

        $archiveSw.Stop()

        # ── GRAND SUMMARY ─────────────────────────────────────────────────────
        Write-Host ""
        Write-SummaryBox -Title "LEAVER ONEDRIVE ARCHIVE COMPLETE" -Lines @(
            @{ Label = "Leaver Site";          Value = ($LeaverSiteUrl -replace 'https://','');  Level = "INFO"    }
            @{ Label = "Libraries Processed";  Value = "$($DocLibs.Count)";                       Level = "INFO"    }
            @{ Label = "Successfully Processed";Value = "$GrandArchived";                          Level = $(if($GrandArchived -gt 0){"SUCCESS"}else{"WARN"}) }
            @{ Label = "Libraries Failed";     Value = "$GrandFailed";                             Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) }
            @{ Label = "Total Time";           Value = "$([math]::Floor($archiveSw.Elapsed.TotalHours))h $($archiveSw.Elapsed.Minutes)m $($archiveSw.Elapsed.Seconds)s"; Level = "INFO" }
        ) -BoxColor "DarkYellow"

        Write-Host ""
        Write-Host "  ┌─ PER-LIBRARY OUTCOMES ─────────────────────────────────" -ForegroundColor DarkYellow
        foreach ($row in $LibResults) {
            $colour = switch ($row.Level) { "SUCCESS"{"Green"} "ERROR"{"Red"} "WARN"{"Yellow"} default{"Cyan"} }
            $lbl    = if ($row.Label.Length -gt 22) { $row.Label.Substring(0,19) + "..." } else { $row.Label.PadRight(22) }
            Write-Host "  │  $lbl  $($row.Status)" -ForegroundColor $colour
        }
        Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkYellow

        return @{ Success = $GrandArchived; Failed = $GrandFailed; LeaverSiteUrl = $LeaverSiteUrl }

    } catch {
        Write-Log "Leaver Archive Engine CRITICAL ERROR: $($_.Exception.Message)" "ERROR"
        Write-Host "  [!] Critical error: $($_.Exception.Message)" -ForegroundColor Red
        return @{ Success = 0; Failed = 1 }
    }
}

# ========================================================================================
# ENGINE SUITE F-2: MULTI-SOURCE LEAVER ONEDRIVE ARCHIVE
# ========================================================================================

# --- FUNCTION: Perform-MultiSourceLeaverArchive ---
# PURPOSE: Unattended multi-leaver archive engine. Accepts a pre-built list of leavers
# (each with a personal site URL and a friendly label) and archives every leaver's
# OneDrive in sequence using a standard PnP UseWebLogin connection.
#
# FLOW PER LEAVER:
#   1. Attempt standard PnP connection (UseWebLogin).
#      On first connection a browser sign-in window may appear. Subsequent leavers
#      on the same tenant usually reuse the existing session without prompting again.
#   2. If connected → call Perform-LeaverOneDriveArchive (enumerates user libraries,
#      runs batch archive on each). Content lands directly in the leaver's label subfolder.
#   3. If connection fails → record "CONNECTION FAILED" and move to the next leaver.
#      All other leavers continue unaffected.
#
# PARAMETERS:
#   Leavers     : List[hashtable] — each entry: SiteUrl, Label
#   LocalDest   : Root local folder; each leaver gets a subfolder named Label
#   BufferDest  : Temp buffer path
function Perform-MultiSourceLeaverArchive {
    param(
        [System.Collections.Generic.List[hashtable]]$Leavers,
        [string]$LocalDest,
        [string]$BufferDest
    )

    $archiveSw     = [System.Diagnostics.Stopwatch]::StartNew()
    $totalSources  = $Leavers.Count
    $GrandArchived = 0; $GrandFailed = 0; $GrandSkipped = 0
    $ResultRows    = [System.Collections.Generic.List[hashtable]]::new()

    Write-Log "--- STARTING MULTI-SOURCE LEAVER ARCHIVE ($totalSources leaver(s)) ---" "INFO"
    Write-Log "SHARED DESTINATION: $LocalDest" "INFO"

    for ($idx = 0; $idx -lt $totalSources; $idx++) {
        $leaver     = $Leavers[$idx]
        $siteUrl    = $leaver.SiteUrl
        $label      = $leaver.Label
        $leaverDest = Join-Path $LocalDest $label

        Write-Host ""
        Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor DarkYellow
        Write-Host "  LEAVER $($idx+1) of $totalSources : $label" -ForegroundColor DarkYellow
        Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor DarkYellow
        Write-Host "  Site   : $siteUrl" -ForegroundColor DarkGray
        Write-Host "  Saving : $leaverDest" -ForegroundColor DarkGray

        Write-Log "MULTI-LEAVER SOURCE $($idx+1)/$totalSources : $label => $siteUrl" "INFO"

        if (-not (Test-Path -LiteralPath $leaverDest)) {
            New-Item -ItemType Directory -Path $leaverDest -Force | Out-Null
        }

        # ── STANDARD PNP CONNECTION ───────────────────────────────────────────
        $connected     = $false
        $connectErrMsg = ""

        Write-Host ""
        Write-Host "  Connecting..." -ForegroundColor DarkGray
        Write-Host "  A browser sign-in window may appear — sign in with your account." -ForegroundColor Yellow

        try {
            $origWarn = $WarningPreference; $WarningPreference = "SilentlyContinue"
            Connect-PnPOnline -Url $siteUrl -UseWebLogin -ErrorAction Stop
            $WarningPreference = $origWarn
            # Quick access test
            $null = Get-PnPList -ErrorAction Stop | Select-Object -First 1
            $connected = $true
            Write-Host "  [✔] Connected." -ForegroundColor Green
            Write-Log "Multi-Leaver: Connection succeeded for '$label'." "SUCCESS"
        } catch {
            $WarningPreference = "Continue"
            $connectErrMsg = $_.Exception.Message
            Write-Host "  [!] Connection failed: $connectErrMsg" -ForegroundColor Red
            Write-Log "Multi-Leaver: Connection failed for '$label': $connectErrMsg" "WARN"
            try { $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
        }

        # ── ARCHIVE IF CONNECTED ──────────────────────────────────────────────
        if ($connected) {
            try {
                $result  = Perform-LeaverOneDriveArchive -LeaverSiteUrl $siteUrl -LocalDest $leaverDest -BufferDest $BufferDest
                $success = if ($result -is [hashtable]) { $result.Success } else { 1 }
                $failed  = if ($result -is [hashtable]) { $result.Failed  } else { 0 }

                if ($failed -eq 0) {
                    $GrandArchived++
                    $ResultRows.Add(@{ Label=$label; Status="ARCHIVED"; Level="SUCCESS" })
                } else {
                    $GrandFailed++
                    $ResultRows.Add(@{ Label=$label; Status="ARCHIVED WITH $failed LIBRARY ERROR(S)"; Level="WARN" })
                }
            } catch {
                $GrandFailed++
                Write-Log "Multi-Leaver: Archive engine error for '$label': $($_.Exception.Message)" "ERROR"
                $ResultRows.Add(@{ Label=$label; Status="ARCHIVE ERROR — $($_.Exception.Message)"; Level="ERROR" })
            } finally {
                try { $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
            }
        } else {
            $GrandSkipped++
            $ResultRows.Add(@{ Label=$label; Status="CONNECTION FAILED — check account access"; Level="ERROR" })
            Write-Log "Multi-Leaver: Skipping '$label' — connection failed." "ERROR"
        }
    }

    $archiveSw.Stop()

    # ── GRAND SUMMARY ─────────────────────────────────────────────────────────
    Write-Host ""
    Write-SummaryBox -Title "MULTI-SOURCE LEAVER ARCHIVE COMPLETE" -Lines @(
        @{ Label = "Leavers Processed";     Value = "$totalSources";    Level = "INFO"    }
        @{ Label = "Successfully Archived"; Value = "$GrandArchived";   Level = $(if($GrandArchived -gt 0){"SUCCESS"}else{"WARN"}) }
        @{ Label = "Failed / Skipped";      Value = "$($GrandFailed + $GrandSkipped)"; Level = $(if(($GrandFailed+$GrandSkipped) -gt 0){"ERROR"}else{"SUCCESS"}) }
        @{ Label = "Total Time";            Value = "$([math]::Floor($archiveSw.Elapsed.TotalHours))h $($archiveSw.Elapsed.Minutes)m $($archiveSw.Elapsed.Seconds)s"; Level = "INFO" }
    ) -BoxColor "DarkYellow"

    Write-Host ""
    Write-Host "  ┌─ PER-LEAVER OUTCOMES ──────────────────────────────────" -ForegroundColor DarkYellow
    foreach ($row in $ResultRows) {
        $colour = switch ($row.Level) { "SUCCESS"{"Green"} "ERROR"{"Red"} "WARN"{"Yellow"} default{"Cyan"} }
        $lbl    = if ($row.Label.Length -gt 24) { $row.Label.Substring(0,21) + "..." } else { $row.Label.PadRight(24) }
        Write-Host "  │  $lbl  $($row.Status)" -ForegroundColor $colour
    }
    Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkYellow

    # Show recycle bin links for all successfully archived leavers
    $rbLinks = @($Leavers | ForEach-Object {
        $rbUrl = Get-LeaverRecycleBinUrl -LeaverUrl $_.SiteUrl
        if ($rbUrl) { @{ Label = $_.Label; Url = $rbUrl } }
    } | Where-Object { $_ })
    if ($rbLinks.Count -gt 0) {
        Write-RecycleBinLinks -Links $rbLinks
    }

    Write-Log "Multi-Leaver Archive: Complete. Archived=$GrandArchived Failed=$GrandFailed Skipped=$GrandSkipped Time=$([math]::Floor($archiveSw.Elapsed.TotalHours))h$($archiveSw.Elapsed.Minutes)m$($archiveSw.Elapsed.Seconds)s" "INFO"
}

# ========================================================================================

# --- FUNCTION: Perform-MultiSourceArchive ---
# PURPOSE: Accepts a pre-built list of source entries (each with a parsed CloudData
# hashtable and a local subfolder name) plus a shared destination root and temp
# buffer. For each source it runs the full Archive workflow:
#   Step 1 — Download  (Perform-DirectDownload)
#   Step 2 — Verify    (Invoke-ArchiveVerify — counts + bytes, no filename matching)
#   Step 3 — Delete    (auto, only if verification passes)
#
# AUTH LIFECYCLE: Connects to each source's SiteUrl with UseWebLogin. If consecutive
# sources share the same SiteUrl the existing connection is reused (no extra browser
# prompts). When the SiteUrl changes the function disconnects and reconnects so the
# correct site context is active for Get-PnPFolderItem / Remove-PnPFolder calls.
#
# Called by the [19] menu handler after the URL-collection UI has run. Auth is
# managed internally — each source reconnects only when the site URL changes.
function Perform-MultiSourceArchive {
    param(
        [System.Collections.Generic.List[hashtable]]$Sources,
        [string]$LocalDest,
        [string]$BufferDest
    )

    $totalSources   = $Sources.Count
    $archiveSw      = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Log "--- STARTING MULTI-SOURCE ARCHIVE ($totalSources source(s)) ---" "INFO"
    Write-Log "SHARED DESTINATION: $LocalDest"

    # Grand counters across all sources
    $GrandArchived    = 0; [long]$GrandArchivedBytes = 0
    $GrandDeleted     = 0
    $GrandDlFailed    = 0; $GrandVerifyFailed = 0

    $ResultRows = [System.Collections.Generic.List[hashtable]]::new()

    $lastSiteUrl    = ""
    $currentSrcIdx  = 0

    foreach ($src in $Sources) {
        $currentSrcIdx++
        $srcLabel     = $src.Label
        $CloudData    = $src.CloudData
        $subFolder    = $src.SubFolder
        $itemLocalDst = Join-Path $LocalDest $subFolder
        $global:JobCancelled = $false

        Write-Host ""
        Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor Yellow
        Write-Host "  SOURCE $currentSrcIdx of $totalSources : $srcLabel" -ForegroundColor Yellow
        Write-Host "  ════════════════════════════════════════════════════════" -ForegroundColor Yellow
        Write-Host "  Cloud  : $($CloudData.FolderPath)" -ForegroundColor DarkGray
        Write-Host "  Saving : $itemLocalDst" -ForegroundColor DarkGray
        Write-Host ""
        Write-Log "MULTI-ARCHIVE SOURCE $currentSrcIdx/$totalSources : $srcLabel => $($CloudData.FolderPath)" "INFO"

        # ── AUTH: connect / reconnect if site changes ──────────────────────────
        if ($CloudData.SiteUrl -ne $lastSiteUrl) {
            if ($lastSiteUrl -ne "") {
                try { $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
                Write-Host "  Reconnecting to new site: $($CloudData.SiteUrl)" -ForegroundColor DarkGray
            } else {
                Write-Host "  Connecting to: $($CloudData.SiteUrl)" -ForegroundColor DarkGray
            }
            try {
                $originalWarning   = $WarningPreference
                $WarningPreference = "SilentlyContinue"
                Connect-PnPOnline -Url $CloudData.SiteUrl -UseWebLogin -ErrorAction Stop
                $WarningPreference = $originalWarning
                $lastSiteUrl = $CloudData.SiteUrl
                Write-Host "  Connected." -ForegroundColor Green
            } catch {
                Write-Host "  [!] Could not connect to $($CloudData.SiteUrl)" -ForegroundColor Red
                Write-Host "      Reason: $($_.Exception.Message)" -ForegroundColor Red
                Write-Log "Multi-Archive: Connection failed for '$srcLabel' ($($CloudData.SiteUrl)). Skipping." "ERROR"
                $GrandDlFailed++
                $ResultRows.Add(@{ Label=$srcLabel; Status="CONNECTION FAILED"; Level="ERROR" })
                continue
            }
        }

        # ── STEP 1: DOWNLOAD ──────────────────────────────────────────────────
        Write-Host "  [1/3] Downloading..." -ForegroundColor Magenta
        $dlResult = Perform-DirectDownload -CloudData $CloudData -LocalDest $itemLocalDst -BufferDest $BufferDest

        if ($global:JobCancelled) {
            Write-Log "Multi-Archive: Cancelled during download of '$srcLabel'. Remaining sources skipped." "WARN"
            $ResultRows.Add(@{ Label=$srcLabel; Status="CANCELLED"; Level="WARN" })
            break
        }

        if ($dlResult.Failed -gt 0) {
            Write-Host "  [!] $($dlResult.Failed) download failure(s) — verification skipped, NOT deleting." -ForegroundColor Red
            Write-Log "Multi-Archive: '$srcLabel' had $($dlResult.Failed) failure(s). NOT deleting." "ERROR"
            $GrandDlFailed++
            $ResultRows.Add(@{ Label=$srcLabel; Status="DOWNLOAD FAILURES ($($dlResult.Failed)) — NOT DELETED"; Level="ERROR" })
            continue
        }

        # ── STEP 2: VERIFY ────────────────────────────────────────────────────
        Write-Host "  [2/3] Verifying..." -ForegroundColor Yellow
        $vResult = Invoke-ArchiveVerify -CloudData $CloudData -DlResult $dlResult

        $localSummary = "$($vResult.LocalFileCount) files | $(Format-FileSize $vResult.LocalTotalBytes)"
        $cloudSummary = "$($vResult.CloudFileCount) files | $(Format-FileSize $vResult.CloudTotalBytes)"

        if (-not $vResult.Passed) {
            Write-Host "  [!] Verification FAILED (local: $localSummary | cloud: $cloudSummary) — NOT deleting." -ForegroundColor Red
            Write-Log "Multi-Archive: '$srcLabel' failed verification. Local=$localSummary | Cloud=$cloudSummary. NOT deleting." "ERROR"
            $GrandVerifyFailed++
            $ResultRows.Add(@{ Label=$srcLabel; Status="VERIFY FAILED (local $localSummary vs cloud $cloudSummary) — NOT DELETED"; Level="ERROR" })
            continue
        }

        Write-Host "  [✔] Verified: $localSummary matches cloud." -ForegroundColor Green
        Write-Log "Multi-Archive: '$srcLabel' verified OK — $localSummary." "SUCCESS"
        $GrandArchived++
        $GrandArchivedBytes += $vResult.LocalTotalBytes

        # ── STEP 3: AUTO-DELETE ───────────────────────────────────────────────
        Write-Host "  [3/3] Deleting from SharePoint Recycle Bin..." -ForegroundColor Cyan
        try {
            $ParentPath = (Split-Path $CloudData.FolderPath -Parent) -replace '\\', '/'
            $FolderName = Split-Path $CloudData.FolderPath -Leaf
            Invoke-PnPWithRetry -OperationName "DELETE (MULTI-ARCHIVE) $FolderName" -ScriptBlock {
                Remove-PnPFolder -Name $FolderName -Folder $ParentPath -Force -Recycle -ErrorAction Stop
            } | Out-Null
            Write-Host "  [✔] '$srcLabel' sent to SharePoint Recycle Bin." -ForegroundColor Green
            Write-Log "Multi-Archive: '$srcLabel' deleted (Recycle Bin)." "SUCCESS"
            $GrandDeleted++
            $ResultRows.Add(@{ Label=$srcLabel; Status="ARCHIVED + DELETED — $localSummary"; Level="SUCCESS" })
        } catch {
            Write-Host "  [!] Deletion failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Multi-Archive: Deletion failed for '$srcLabel'. Reason: $($_.Exception.Message)" "ERROR"
            $ResultRows.Add(@{ Label=$srcLabel; Status="ARCHIVED but DELETE FAILED — $localSummary"; Level="WARN" })
        }
    }

    # ── DISCONNECT ────────────────────────────────────────────────────────────
    try { $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}

    $archiveSw.Stop()

    # ── GRAND SUMMARY ─────────────────────────────────────────────────────────
    Write-Host ""
    $summaryLines = [System.Collections.Generic.List[hashtable]]::new()
    $summaryLines.Add(@{ Label="Sources Processed";   Value="$currentSrcIdx of $totalSources";                      Level="INFO"    })
    $summaryLines.Add(@{ Label="Successfully Archived";Value="$GrandArchived  |  $(Format-FileSize $GrandArchivedBytes)"; Level="SUCCESS" })
    $summaryLines.Add(@{ Label="Cloud Copies Deleted"; Value="$GrandDeleted";                                        Level=if($GrandDeleted -eq $GrandArchived -and $GrandArchived -gt 0){"SUCCESS"}else{"WARN"} })
    $summaryLines.Add(@{ Label="Failed Verification";  Value="$GrandVerifyFailed";                                   Level=if($GrandVerifyFailed -gt 0){"ERROR"}else{"SUCCESS"} })
    $summaryLines.Add(@{ Label="Failed Download";      Value="$GrandDlFailed";                                       Level=if($GrandDlFailed -gt 0){"ERROR"}else{"SUCCESS"} })
    $summaryLines.Add(@{ Label="Total Time";           Value="$([math]::Floor($archiveSw.Elapsed.TotalHours))h $($archiveSw.Elapsed.Minutes)m $($archiveSw.Elapsed.Seconds)s"; Level="INFO" })
    Write-SummaryBox -Title "MULTI-SOURCE ARCHIVE COMPLETE" -Lines $summaryLines.ToArray() -BoxColor "Yellow"

    # Per-source outcome table
    Write-Host ""
    Write-Host "  ┌─ PER-SOURCE OUTCOMES ─────────────────────────────────────" -ForegroundColor Cyan
    foreach ($row in $ResultRows) {
        $colour = switch ($row.Level) { "SUCCESS"{"Green"} "ERROR"{"Red"} "WARN"{"Yellow"} default{"Cyan"} }
        $label  = if ($row.Label.Length -gt 22) { $row.Label.Substring(0,19) + "..." } else { $row.Label.PadRight(22) }
        Write-Host "  │  $label  $($row.Status)" -ForegroundColor $colour
    }
    Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host ""
}


# ENGINE SUITE E: SHAREPOINT UPLOAD TOOLS (new in v89)
# ========================================================================================

# --- FUNCTION: Ensure-CloudFolder ---
# PURPOSE: Guarantees that a full SharePoint folder hierarchy exists before uploading.
#
# PRIMARY:  Resolve-PnPFolder -SiteRelativePath (PnP.PowerShell 1.12.0+) creates the
#           entire path in one API call, including all intermediate sub-folders.
# FALLBACK: If that throws, we walk each path segment manually — Get-PnPFolder to check,
#           Add-PnPFolder to create — ensuring compatibility with edge-case environments.
# CACHING:  Callers maintain a HashSet and skip this call for folders already confirmed.
#           This avoids hammering the SharePoint API with repeated folder-existence checks
#           when uploading many files into the same sub-folder structure.
#
# INPUT:    Site-relative path, e.g. "Shared Documents/Leavers/2024/John Smith"
#           Do NOT include the /sites/SiteName prefix.
function Ensure-CloudFolder {
    param([string]$SiteRelativePath)
    try {
        Resolve-PnPFolder -SiteRelativePath $SiteRelativePath -ErrorAction Stop | Out-Null
        return $true
    } catch {
        # Manual fallback: walk each segment, check and create as needed
        try {
            $parts = $SiteRelativePath -split '/'
            $built = ""
            foreach ($part in $parts) {
                $part = $part.Trim()
                if ([string]::IsNullOrWhiteSpace($part)) { continue }
                $parent = $built
                $built  = if ($built) { "$built/$part" } else { $part }
                $exists = $false
                try { Get-PnPFolder -Url $built -ErrorAction Stop | Out-Null; $exists = $true } catch {}
                if (-not $exists) {
                    try { Add-PnPFolder -Name $part -Folder $parent -ErrorAction Stop | Out-Null } catch {}
                }
            }
            return $true
        } catch { return $false }
    }
}

# --- FUNCTION: Invoke-UploadWithRetry ---
# PURPOSE: Uploads a single local file to a SharePoint folder via Add-PnPFile.
# Uses the same 5-attempt exponential backoff as Perform-SafeDownload:
# sleeps 5s / 10s / 15s / 20s / 25s between attempts to handle API throttling.
#
# PARAMETERS:
#   LocalPath        — Full local (long-path-safe) path to the file to upload.
#   DstFolderSiteRel — Site-relative URL of the destination folder in SharePoint.
#                      Example: "Shared Documents/Leavers/2024"
#   FileName         — The target filename as it should appear in SharePoint.
#                      May differ from the local name if Get-SafeWindowsName was applied.
function Invoke-UploadWithRetry {
    param(
        [string]$LocalPath,
        [string]$DstFolderSiteRel,
        [string]$FileName
    )
    $maxRetries = 5
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        if ($global:JobCancelled) { return @{ Success = $false; Error = "Aborted by User" } }
        try {
            Add-PnPFile -Path $LocalPath -Folder $DstFolderSiteRel -NewFileName $FileName -ErrorAction Stop | Out-Null
            return @{ Success = $true; Error = $null }
        } catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                return @{ Success = $false; Error = $_.Exception.Message }
            }
            $WaitTime = $retryCount * 5
            Write-Host "   [API Throttle] Retrying in $WaitTime seconds... ($retryCount/$maxRetries)" -ForegroundColor Yellow
            Start-Sleep -Seconds $WaitTime
        }
    }
    return @{ Success = $false; Error = "Max retries exceeded" }
}

# --- FUNCTION: Perform-DirectUpload ---
# PURPOSE: Uploads an entire local folder (all sub-folders and files) to a SharePoint
# destination. Mirrors Perform-DirectDownload in architecture, conventions, and UI output.
#
# SKIP LOGIC:    Get-PnPFile checks whether the file already exists on SharePoint with
#                the same byte size. Identical files are skipped — safe resume behaviour.
# FOLDER CREATE: Ensure-CloudFolder is called before each new cloud folder path is needed.
#                A HashSet ($FolderCache) prevents repeated API calls for the same folder.
# PATH GUARD:    Enforces the $CloudPathLimit (400 chars) — over-limit files are logged
#                to LongPaths_FailedToSync.txt, identical to the local copy engine.
# SINGLE FILE:   Detected via [System.IO.File]::Exists — bypasses the directory scan.
# BATCH AWARE:   When called from Perform-LocalToCloudBatchUpload, the -BatchStatus and
#                -BatchPercent parameters route progress to the parent bar and suppress
#                the per-item summary box — identical to the download engine pattern.
function Perform-DirectUpload {
    param(
        [hashtable]$CloudDst,
        [string]$LocalSrc,
        [string]$BatchStatus    = "",
        [int]$BatchPercent      = -1,
        [datetime]$InheritedStartTime = [datetime]::MinValue
    )

    $FailedList = [System.Collections.Generic.List[hashtable]]::new()
    $targetName = Split-Path $LocalSrc -Leaf

    if (-not $BatchStatus) {
        $global:JobCancelled = $false
        Clear-KeyBuffer
        Write-Log "--- STARTING DIRECT UPLOAD ---" "INFO"
        Write-Log "LOCAL SOURCE: $(Format-PathWithStorageTag $LocalSrc)"
        Write-Log "CLOUD DESTINATION: $($CloudDst.FolderPath)"
    }

    try {
        $SitePath       = [uri]::new($CloudDst.SiteUrl).AbsolutePath
        $DstSiteRelRoot = $CloudDst.FolderPath.Substring($SitePath.Length).TrimStart('/')

        $sw           = [System.Diagnostics.Stopwatch]::StartNew()
        $SuccessCount = 0; $SkippedCount = 0; $FailedCount = 0
        [long]$SuccessBytes = 0; [long]$SkippedBytes = 0

        $jobStartTime = if ($InheritedStartTime -ne [datetime]::MinValue) { $InheritedStartTime } else { Get-Date }
        $startTimeStr = $jobStartTime.ToString("HH:mm:ss")
        $uiStopwatch  = [System.Diagnostics.Stopwatch]::StartNew()
        $activityName = if ($BatchStatus) { "Batch Upload Engine [Started: $startTimeStr]" } else { "Upload Engine [Started: $startTimeStr]" }
        $LongSrc      = Get-LongPath $LocalSrc

        # ── SINGLE FILE CASE ──────────────────────────────────────────────────
        if ([System.IO.File]::Exists($LongSrc)) {
            $safeName = Get-SafeWindowsName ([System.IO.Path]::GetFileName($LocalSrc))
            $fSize    = [System.IO.FileInfo]::new($LongSrc).Length
            $sizeStr  = Format-FileSize $fSize

            $ShouldUpload = $true
            try {
                $existingFile = Get-PnPFile -Url "$DstSiteRelRoot/$safeName" -ErrorAction SilentlyContinue
                if ($existingFile -and $existingFile.Length -eq $fSize) {
                    $ShouldUpload = $false; $SkippedCount++; $SkippedBytes += $fSize
                    Write-Log "SKIPPED (Already Synced): $safeName [Size: $sizeStr]" "WARN" -ConsoleOnly
                }
            } catch {}

            if ($ShouldUpload) {
                Ensure-CloudFolder -SiteRelativePath $DstSiteRelRoot | Out-Null
                # v109: large file indicator matches download engine's "please wait..." message
                if ($fSize -ge $script:LargeFileBytes) {
                    Write-Host "  Uploading: $safeName ($sizeStr) — large file, please wait..." -ForegroundColor Yellow
                } else {
                    Write-Host "  Uploading: $safeName [$sizeStr]..." -ForegroundColor Cyan
                }
                $r = Invoke-UploadWithRetry -LocalPath $LocalSrc -DstFolderSiteRel $DstSiteRelRoot -FileName $safeName
                if ($r.Success) {
                    $SuccessCount++; $SuccessBytes += $fSize
                    $global:SessionTransferredBytes += $fSize   # v109: running session total (matches download)
                    Write-Log "Uploaded: $safeName [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                } else {
                    $FailedCount++
                    Write-Log "Upload Failed: $safeName [Reason: $($r.Error)]" "ERROR"
                    $FailedList.Add(@{ File = $safeName; Path = $LocalSrc; Reason = $r.Error })
                }
            }

        } else {
            # ── DIRECTORY CASE: Enumerate all files recursively ──────────────────
            # v110: emit a heartbeat every 2 s during enumeration so operators don't
            # think the script is hung when scanning very large local trees. The
            # heartbeat self-suppresses for fast scans (<2 s) so it never spams.
            $AllFiles    = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            $enumSw      = [System.Diagnostics.Stopwatch]::StartNew()
            $heartbeatSw = [System.Diagnostics.Stopwatch]::StartNew()
            $shownHeartbeat = $false
            try {
                $dirInfo = [System.IO.DirectoryInfo]::new($LongSrc)
                foreach ($f in $dirInfo.EnumerateFiles("*", [System.IO.SearchOption]::AllDirectories)) {
                    $AllFiles.Add($f) | Out-Null
                    if ($heartbeatSw.ElapsedMilliseconds -gt 2000) {
                        $shownHeartbeat = $true
                        $msg = "  [..] Scanning local source : {0:N0} files found ({1}s)..." -f $AllFiles.Count, [int]$enumSw.Elapsed.TotalSeconds
                        Write-Host "`r$($msg.PadRight(72))" -NoNewline -ForegroundColor DarkGray
                        $heartbeatSw.Restart()
                    }
                }
            } catch {
                # Fallback for access-restricted paths — Get-ChildItem is more tolerant
                # but doesn't stream, so no heartbeat possible here. Just one message.
                if ($shownHeartbeat) { Write-Host "" }
                Write-Log "EnumerateFiles failed, falling back to Get-ChildItem: $($_.Exception.Message)" "WARN"
                $tempFiles = Get-ChildItem -LiteralPath $LongSrc -Recurse -File -Force -ErrorAction SilentlyContinue
                foreach ($f in $tempFiles) { $AllFiles.Add([System.IO.FileInfo]::new($f.FullName)) | Out-Null }
            }
            # Finalize heartbeat line with completion message (only if we actually showed any)
            if ($shownHeartbeat) {
                $finalMsg = "  [OK] Scan complete: {0:N0} files in {1}s" -f $AllFiles.Count, [int]$enumSw.Elapsed.TotalSeconds
                Write-Host "`r$($finalMsg.PadRight(72))" -ForegroundColor DarkGray
            }

            $TotalFiles     = $AllFiles.Count
            $CurrentFileIdx = 0
            # HashSet: once a cloud folder is confirmed to exist, skip future API calls for it
            $FolderCache = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            if (-not $BatchStatus) {
                Write-Host ""
                Write-Host "  Preparing to upload $TotalFiles file(s) to SharePoint..." -ForegroundColor DarkGray
                Write-Host "  HOTKEYS: Press 'P' to Pause or 'C' to Cancel at any time." -ForegroundColor DarkGray
                Write-Host ""
            }

            foreach ($file in $AllFiles) {
                if ($global:JobCancelled) { break }
                Check-Interrupt -CurrentFile $file.Name -CurrentFolder $file.DirectoryName | Out-Null
                if ($global:JobCancelled) { break }

                $CurrentFileIdx++

                # Relative path from source root (e.g. "SubFolder\Nested\File.txt")
                $relFilePath  = $file.FullName.Substring($LongSrc.Length).TrimStart('\')
                $relDir       = [System.IO.Path]::GetDirectoryName($relFilePath)
                $safeName     = Get-SafeWindowsName $file.Name
                $fSize        = $file.Length
                $sizeStr      = Format-FileSize $fSize

                # Build site-relative URL for the target cloud folder
                $relDirCloud        = if ([string]::IsNullOrWhiteSpace($relDir)) { "" } else { $relDir.Replace('\', '/') }
                $cloudFolderSiteRel = if ([string]::IsNullOrWhiteSpace($relDirCloud)) { $DstSiteRelRoot } else { "$DstSiteRelRoot/$relDirCloud" }
                $cloudFileSiteRel   = "$cloudFolderSiteRel/$safeName"

                # Cloud path length guard — mirrors $CloudPathLimit for local copy ops
                if ($cloudFileSiteRel.Length -gt $CloudPathLimit) {
                    Write-Log "SKIPPED: Cloud path exceeds $CloudPathLimit chars ($($cloudFileSiteRel.Length)): $cloudFileSiteRel" "WARN" -ConsoleOnly
                    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "[$Timestamp] OVER CLOUD LIMIT ($($cloudFileSiteRel.Length) chars): $cloudFileSiteRel" | Out-File -FilePath $LongPathReport -Append -Encoding UTF8
                    $SkippedCount++; $SkippedBytes += $fSize
                    continue
                }

                # Progress bar
                if ($uiStopwatch.ElapsedMilliseconds -gt 5000 -or $CurrentFileIdx -eq $TotalFiles) {
                    $pct        = [math]::Min(100, [math]::Round(($CurrentFileIdx / $TotalFiles) * 100))
                    $elapsed    = (Get-Date) - $jobStartTime
                    $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                    # v109: consistency with download engine — session total + ETA in CurrentOperation
                    $sessionStr = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
                    $eta = Get-ETA -Current $CurrentFileIdx -Total $TotalFiles -Elapsed $elapsed
                    $etaStr = if ($eta) { " | ETA: $eta" } else { "" }
                    if ($BatchPercent -ge 0) {
                        # v110: nested under Id 0 batch engine — show file-level percent (not batch %)
                        # so the operator sees independent info on each bar.
                        Write-Progress -Id 1 -ParentId 0 -Activity "  └─ $activityName" -Status "File $CurrentFileIdx of $TotalFiles" -PercentComplete $pct -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
                    } else {
                        Write-Progress -Id 1 -Activity $activityName -Status "Uploading file $CurrentFileIdx of $TotalFiles" -PercentComplete $pct -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
                    }
                    $uiStopwatch.Restart()
                }

                # Ensure cloud folder exists — use cache to skip redundant API calls
                if (-not $FolderCache.Contains($cloudFolderSiteRel)) {
                    if (Ensure-CloudFolder -SiteRelativePath $cloudFolderSiteRel) {
                        $FolderCache.Add($cloudFolderSiteRel) | Out-Null
                    } else {
                        Write-Log "WARNING: Could not verify/create cloud folder: $cloudFolderSiteRel" "WARN"
                    }
                }

                # Skip-if-identical: byte-size check, same logic as download engine
                $ShouldUpload = $true
                try {
                    $existingFile = Get-PnPFile -Url $cloudFileSiteRel -ErrorAction SilentlyContinue
                    if ($existingFile -and $existingFile.Length -eq $fSize) {
                        $ShouldUpload = $false; $SkippedCount++; $SkippedBytes += $fSize
                        Write-Log "SKIPPED (Already Synced): $relFilePath [Size: $sizeStr]" "WARN" -ConsoleOnly
                    }
                } catch {}

                if ($ShouldUpload) {
                    # v109: large file indicator + session bytes tracking — parity with download
                    if ($fSize -ge $script:LargeFileBytes) {
                        Write-Host "  Uploading: $relFilePath ($sizeStr) — large file, please wait..." -ForegroundColor Yellow
                    } else {
                        Write-Host "  Uploading: $relFilePath [$sizeStr]..." -ForegroundColor Cyan
                    }
                    $r = Invoke-UploadWithRetry -LocalPath $file.FullName -DstFolderSiteRel $cloudFolderSiteRel -FileName $safeName
                    if ($r.Success) {
                        $SuccessCount++; $SuccessBytes += $fSize
                        $global:SessionTransferredBytes += $fSize   # v109: running session total
                        Write-Log "Uploaded: $relFilePath [Size: $sizeStr]" "SUCCESS" -ConsoleOnly
                    } else {
                        $FailedCount++
                        Write-Log "Upload Failed: $relFilePath [Reason: $($r.Error)]" "ERROR"
                        $FailedList.Add(@{ File = $file.Name; Path = $relFilePath; Reason = $r.Error })
                    }
                }
            }

            if (-not $BatchStatus) { Write-Progress -Id 1 -Activity $activityName -Completed }
        }

        $sw.Stop()
        Write-Host ""
        if ($FailedList.Count -gt 0) { Write-FailedReport -FailedList $FailedList }

        $statusTitle = if ($global:JobCancelled) { "UPLOAD ABORTED" } else { "UPLOAD COMPLETE" }
        $statusColor = if ($global:JobCancelled) { "Yellow" } else { "Blue" }

        if (-not $BatchStatus) {
            # v109: list pattern + avg speed + session total — parity with download summary
            $upLines = [System.Collections.Generic.List[hashtable]]::new()
            $upLines.Add(@{ Label = "Target";                    Value = $targetName;                                          Level = "INFO"    })
            # v110: storage tag — local source
            $srcTag = Get-PathStorageType $LocalSrc
            if ($srcTag) {
                $upLines.Add(@{ Label = "Storage";               Value = "$srcTag -> Cloud";                                  Level = "INFO"    })
            }
            $upLines.Add(@{ Label = "Successfully Uploaded";     Value = "$SuccessCount  |  $(Format-FileSize $SuccessBytes)"; Level = "SUCCESS" })
            $upLines.Add(@{ Label = "Skipped (already existed)"; Value = "$SkippedCount  |  $(Format-FileSize $SkippedBytes)"; Level = "INFO"    })
            $upLines.Add(@{ Label = "Failed";                    Value = "$FailedCount";                                       Level = $(if($FailedCount -gt 0){"ERROR"}else{"SUCCESS"}) })
            $upLines.Add(@{ Label = "Total Time";                Value = "$([math]::Floor($sw.Elapsed.TotalHours))h $($sw.Elapsed.Minutes)m $($sw.Elapsed.Seconds)s"; Level = "INFO" })
            $avgSpeed = Format-AvgSpeed -Bytes $SuccessBytes -Elapsed $sw.Elapsed
            if ($avgSpeed) {
                $upLines.Add(@{ Label = "Average Speed";         Value = $avgSpeed; Level = "INFO" })
            }
            # v110: files/min — useful for uploads of many tiny files
            $filesPerMin = Format-FilesPerMin -Count $SuccessCount -Elapsed $sw.Elapsed
            if ($filesPerMin -and $SuccessCount -ge 50) {
                $upLines.Add(@{ Label = "Throughput";            Value = $filesPerMin; Level = "INFO" })
            }
            if ($global:SessionTransferredBytes -gt 0) {
                $upLines.Add(@{ Label = "Session Total";         Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
            }
            Write-SummaryBox -Title $statusTitle -Lines $upLines.ToArray() -BoxColor $statusColor
            # v111: failure-rate alert + ALL distinct errors recap + per-extension breakdown
            # ($AllFiles is set in directory mode; in single-file mode it's $null and the
            # breakdown is suppressed inside Show-OperationDiagnostics)
            $uploadItems = if ($AllFiles) { $AllFiles } else { $null }
            Show-OperationDiagnostics -SuccessCount $SuccessCount -FailedCount $FailedCount -FailedList $FailedList -ProcessedItems $uploadItems
        } else {
            Write-Log "$statusTitle : $targetName | Uploaded: $SuccessCount | Skipped: $SkippedCount | Failed: $FailedCount | $(Format-FileSize $SuccessBytes)" $(if($global:JobCancelled){"WARN"}else{"INFO"})
        }

        return @{ Success = $SuccessCount; Skipped = $SkippedCount; Failed = $FailedCount; SuccessBytes = $SuccessBytes; SkippedBytes = $SkippedBytes; FailedItems = $FailedList }

    } catch {
        Write-Log "CRITICAL ERROR in Upload Engine: $($_.Exception.Message)" "ERROR"
        return @{ Success = 0; Skipped = 0; Failed = 1; SuccessBytes = 0; SkippedBytes = 0 }
    }
}

# --- FUNCTION: Perform-LocalToCloudBatchUpload ---
# PURPOSE: Shows the top-level contents of a local folder as a numbered index.
# User picks items to upload to SharePoint. Mirrors Perform-CloudToLocalBatchDownload.
#
# FOLDER items: their CONTENTS are uploaded into a new cloud sub-folder named after
#               the local folder (matching how Batch Download works in reverse).
# FILE items:   uploaded directly into the root cloud destination folder.
function Perform-LocalToCloudBatchUpload {
    param([string]$LocalSrc, [hashtable]$CloudDst)
    Write-Log "--- ENTERING LOCAL TO CLOUD BATCH UPLOAD MODE ---"
    Write-Log "LOCAL SOURCE: $(Format-PathWithStorageTag $LocalSrc)"
    Write-Log "CLOUD DESTINATION: $($CloudDst.FolderPath)"

    while ($true) {
        Write-PathBanner -Label1 "Source" -Path1 $LocalSrc -Label2 "Dest" -Path2 $CloudDst.FolderPath -BoxColor "Blue"
        $Items = @(Get-ChildItem -LiteralPath $LocalSrc -Force | Sort-Object @{Expression={$_.PSIsContainer};Descending=$true}, @{Expression={$_.Name};Ascending=$true})

        if ($Items.Count -eq 0) {
            Write-Log "Source directory is completely empty. Exiting Batch Upload Mode..." "WARN"
            break
        }

        Write-Host ""
        Write-Host "  ┌─ LOCAL DIRECTORY INDEX (for upload) ───────────────────" -ForegroundColor Cyan
        Write-Host "  │  Items = top-level count  │ Size = total recursive size" -ForegroundColor DarkGray
        Write-Host "  Calculating local folder sizes..." -ForegroundColor DarkGray

        $MenuAudit = [System.Collections.Generic.List[string]]::new()
        $MenuAudit.Add("--- TOP-LEVEL DIRECTORY INDEX (BATCH UPLOAD) ---")

        for ($i = 0; $i -lt $Items.Count; $i++) {
            $item = $Items[$i]
            if ($item.PSIsContainer) {
                $metrics = Get-LocalDirectoryMetrics -Path $item.FullName
                if ($metrics) {
                    $sizeStr = $metrics.Size
                    try { $iCount = [System.IO.Directory]::GetFileSystemEntries((Get-LongPath $item.FullName)).Length } catch { $iCount = 0 }
                    $logStr = "[$($i + 1)] [FOLDER] $($item.Name) [Size: $sizeStr | Items: $iCount]"
                } else { $logStr = "[$($i + 1)] [FOLDER] $($item.Name) [Size: Empty/Locked | Items: 0]" }
            } else {
                $logStr = "[$($i + 1)] [FILE]   $($item.Name) [Size: $(Format-FileSize $item.Length)]"
            }
            Write-Host " $logStr"
            $MenuAudit.Add($logStr)
        }

        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogPayload = ""
        foreach ($line in $MenuAudit) { $LogPayload += "[$ts] [MENU] $line`r`n" }
        [System.IO.File]::AppendAllText($LogFile, $LogPayload)

        Write-Host ""
        Write-Host " [A] Upload ALL items in this directory" -ForegroundColor Green
        Write-Host "  │  [S] Search this directory" -ForegroundColor Cyan
        Write-Host "  └─ [Q] Back to Main Menu ─────────────────────────────────" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Selection: 3  │  2-5  │  1,3,7  │  A (all)  │  ENTER (refresh)" -ForegroundColor DarkGray
        Write-Host "  Ranges auto-sort (10-3 = same as 3-10). Duplicates removed." -ForegroundColor DarkGray

        $selectedIndices = @(); $UserWantsToQuit = $false; $RefreshList = $false

        while ($true) {
            $range = Read-Host "Enter choice (Numbers, 'A', 'S', 'Q', or ENTER)"
            Write-Log "[USER INPUT] Batch Upload Range Selection: '$range'" "INFO" -FileOnly

            if ([string]::IsNullOrWhiteSpace($range)) { Write-Log "Refreshing directory inventory..." "INFO"; $RefreshList = $true; break }
            if ($range -match "^[Qq]$") { Write-Log "Exiting Batch Upload Mode..." "INFO"; $UserWantsToQuit = $true; break }

            if ($range -match "^[Ss]$") {
                while ($true) {
                    $searchTerm = Read-Host "Enter FILE NAME or EXTENSION to search for (or 'Q' to exit search)"
                    Write-Log "[USER INPUT] Search Term: '$searchTerm'" "INFO" -FileOnly
                    if ($searchTerm -match "^[Qq]$") { break }
                    if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                        Perform-LocalSearch -SearchPath $LocalSrc -SearchTerm $searchTerm
                        Write-Host "`nPress ENTER to return to the batch menu, or type a new search term..." -ForegroundColor DarkGray
                    }
                }
                $RefreshList = $true; break
            }

            $invalidInput = $false; $tempIndices = @()

            if ($range -match "^[Aa]$") {
                if ($Items.Count -eq 0) { Write-Log "Nothing to upload." "WARN"; $invalidInput = $true }
                else { $tempIndices = 0..($Items.Count - 1); Write-Log "Selected ALL items (1 through $($Items.Count))." "INFO" }
            } else {
                $parts = $range -split ','
                foreach ($part in $parts) {
                    $part = $part.Trim()
                    if ([string]::IsNullOrWhiteSpace($part)) { continue }
                    if ($part -match "^(\d+)\s*-\s*(\d+)$") {
                        $s = [int]$matches[1] - 1; $e = [int]$matches[2] - 1
                        if ($s -gt $e) { $temp = $s; $s = $e; $e = $temp }
                        if ($s -lt 0 -or $e -ge $Items.Count) { Write-Host "ERROR: Range '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red; $invalidInput = $true; break }
                        $tempIndices += $s..$e
                    } elseif ($part -match "^(\d+)$") {
                        $idx = [int]$matches[1] - 1
                        if ($idx -lt 0 -or $idx -ge $Items.Count) { Write-Host "ERROR: Item '$part' is out of bounds. Must be 1 to $($Items.Count)." -ForegroundColor Red; $invalidInput = $true; break }
                        $tempIndices += $idx
                    } else {
                        Write-Host "ERROR: Invalid format '$part'. Use numbers or ranges (e.g., 1, 3-5)." -ForegroundColor Red; $invalidInput = $true; break
                    }
                }
            }

            if ($invalidInput -or $tempIndices.Count -eq 0) { Write-Host "Please try entering your selection again.`n" -ForegroundColor Yellow; continue }
            $selectedIndices = $tempIndices | Select-Object -Unique | Sort-Object
            Write-Log "Selected $($selectedIndices.Count) items for upload." "INFO"
            break
        }

        if ($UserWantsToQuit) { break }
        if ($RefreshList) { continue }

        Write-Log "Initializing Batch Upload..." "INFO"
        $global:JobCancelled = $false
        Clear-KeyBuffer
        Write-Host " HOTKEYS: Press 'P' to Pause or 'C' to Cancel at any time." -ForegroundColor DarkGray

        $batchSw              = [System.Diagnostics.Stopwatch]::StartNew()
        $jobStartTime         = Get-Date
        $startTimeStr         = $jobStartTime.ToString("HH:mm:ss")
        $uiStopwatch          = [System.Diagnostics.Stopwatch]::StartNew()
        $totalItemsToProcess  = $selectedIndices.Count
        $currentItemProcessed = 0

        $GrandSuccess = 0; $GrandSkipped = 0; $GrandFailed = 0
        [long]$GrandSuccessBytes = 0; [long]$GrandSkippedBytes = 0
        $BatchFailedList = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($i in $selectedIndices) {
            $item                = $Items[$i]
            $CurrentFolderContext = if ($item.PSIsContainer) { $item.FullName } else { $item.DirectoryName }
            $CurrentFileContext   = $item.Name

            if ($global:JobCancelled) { break }
            Check-Interrupt -CurrentFile $CurrentFileContext -CurrentFolder $CurrentFolderContext | Out-Null
            if ($global:JobCancelled) { break }

            $currentItemProcessed++
            $batchStatusText = "Processing Batch Item $currentItemProcessed of $($totalItemsToProcess): $($item.Name)"
            $pct             = [math]::Min(100, [math]::Round(($currentItemProcessed / $totalItemsToProcess) * 100))

            if ($uiStopwatch.ElapsedMilliseconds -gt 5000 -or $currentItemProcessed -eq $totalItemsToProcess) {
                $elapsed    = (Get-Date) - $jobStartTime
                $elapsedStr = "{0:00}h {1:00}m {2:00}s" -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
                # v109: ETA + session total — consistency with other batch engines
                $sessionStr = if ($global:SessionTransferredBytes -gt 0) { " | Session: $(Format-FileSize $global:SessionTransferredBytes)" } else { "" }
                $eta = Get-ETA -Current $currentItemProcessed -Total $totalItemsToProcess -Elapsed $elapsed
                $etaStr = if ($eta) { " | ETA: $eta" } else { "" }
                Write-Progress -Id 0 -Activity "Batch Upload Engine [Started: $startTimeStr]" -Status $batchStatusText -PercentComplete $pct -CurrentOperation "Elapsed: $elapsedStr$etaStr$sessionStr | P=Pause C=Cancel"
                $uiStopwatch.Restart()
            }

            Write-Host ""
            Write-Host "  ── [↑] UPLOAD [$currentItemProcessed/$totalItemsToProcess] $($item.Name)" -ForegroundColor Cyan
            Write-Log "UPLOAD ITEM $currentItemProcessed/$totalItemsToProcess : $($item.Name)" "INFO" -FileOnly

            if ($item.PSIsContainer) {
                # FOLDER: upload its contents into a new cloud sub-folder of the same name
                $ItemCloudDst = @{
                    SiteUrl    = $CloudDst.SiteUrl
                    FolderPath = ("$($CloudDst.FolderPath)/$($item.Name)" -replace '//', '/')
                }
                $tRes = Perform-DirectUpload -CloudDst $ItemCloudDst -LocalSrc $item.FullName -BatchStatus $batchStatusText -BatchPercent $pct -InheritedStartTime $jobStartTime | Select-Object -Last 1
            } else {
                # FILE: upload directly into the cloud destination folder (no subfolder created)
                $tRes = Perform-DirectUpload -CloudDst $CloudDst -LocalSrc $item.FullName -BatchStatus $batchStatusText -BatchPercent $pct -InheritedStartTime $jobStartTime | Select-Object -Last 1
            }

            if ($tRes) {
                $GrandSuccess += $tRes.Success; $GrandSkipped += $tRes.Skipped; $GrandFailed += $tRes.Failed
                $GrandSuccessBytes += $tRes.SuccessBytes; $GrandSkippedBytes += $tRes.SkippedBytes
                if ($tRes.FailedItems) { $BatchFailedList.AddRange($tRes.FailedItems) }
            }
        }

        Write-Progress -Id 0 -Activity "Batch Upload Engine [Started: $startTimeStr]" -Completed
        $batchSw.Stop()

        if ($BatchFailedList.Count -gt 0) {
            Write-Host ""
            Write-FailedReport -FailedList $BatchFailedList
        }

        if ($global:JobCancelled) { Write-Log "--- BATCH UPLOAD JOB ABORTED BY USER ---" "WARN" }
        else                      { Write-Log "--- BATCH UPLOAD COMPLETION ---" "INFO" }

        Write-Host ""
        # v109: list pattern + avg speed + session total — full parity with Batch Download
        $buLines = [System.Collections.Generic.List[hashtable]]::new()
        # v110: storage tag — local source root
        $srcTag = Get-PathStorageType $LocalSrc
        if ($srcTag) {
            $buLines.Add(@{ Label = "Storage";             Value = "$srcTag -> Cloud";                                  Level = "INFO"    })
        }
        $buLines.Add(@{ Label = "Successfully Uploaded";   Value = "$GrandSuccess  |  $(Format-FileSize $GrandSuccessBytes)"; Level = "SUCCESS" })
        $buLines.Add(@{ Label = "Skipped (already synced)";Value = "$GrandSkipped  |  $(Format-FileSize $GrandSkippedBytes)"; Level = "INFO"    })
        $buLines.Add(@{ Label = "Failed to Upload";        Value = "$GrandFailed";                                            Level = $(if($GrandFailed -gt 0){"ERROR"}else{"SUCCESS"}) })
        $buLines.Add(@{ Label = "Total Time";              Value = "$([math]::Floor($batchSw.Elapsed.TotalHours))h $($batchSw.Elapsed.Minutes)m $($batchSw.Elapsed.Seconds)s"; Level = "INFO" })
        $avgSpeed = Format-AvgSpeed -Bytes $GrandSuccessBytes -Elapsed $batchSw.Elapsed
        if ($avgSpeed) {
            $buLines.Add(@{ Label = "Average Speed";       Value = $avgSpeed; Level = "INFO" })
        }
        # v110: files/min — useful for batches of many tiny files
        $filesPerMin = Format-FilesPerMin -Count $GrandSuccess -Elapsed $batchSw.Elapsed
        if ($filesPerMin -and $GrandSuccess -ge 50) {
            $buLines.Add(@{ Label = "Throughput";          Value = $filesPerMin; Level = "INFO" })
        }
        if ($global:SessionTransferredBytes -gt 0) {
            $buLines.Add(@{ Label = "Session Total";       Value = "$(Format-FileSize $global:SessionTransferredBytes) transferred this session"; Level = "INFO" })
        }
        Write-SummaryBox -Title "BATCH UPLOAD COMPLETE" -Lines $buLines.ToArray() -BoxColor "Blue"
        # v111: failure-rate alert + ALL distinct errors recap
        Show-OperationDiagnostics -SuccessCount $GrandSuccess -FailedCount $GrandFailed -FailedList $BatchFailedList -ProcessedItems $null
        Write-Host ""
        Write-Host "  Scanning directory again to refresh the menu..." -ForegroundColor DarkGray
    }
}


# ========================================================================================
# ENGINE SUITE G: UNLICENSED LEAVER ONEDRIVE ARCHIVE (new in v104)
# ========================================================================================
#
# PURPOSE: For leavers whose E5 / M365 licence was removed long ago, their personal
# OneDrive is frozen and inaccessible via PnP. Purchasing a small number of SharePoint
# Plan 1 licences and assigning them via an Azure AD security group re-activates the
# OneDrive long enough to download and delete the data. Once the content is archived
# the user is removed from the group immediately, freeing the licence slot for the
# next leaver. Only one licence is needed at a time.
#
# --- FUNCTION: Save-MgTokenCache / Load-MgTokenCache ---
# Persists the Graph refresh token using Windows DPAPI (ConvertFrom-SecureString without
# a key = DPAPI, tied to the current Windows user account — other accounts cannot read it).
# Allows silent reconnect on subsequent sessions without a new device code prompt.
function Save-MgTokenCache {
    param([string]$RefreshToken, [string]$AccountName)
    try {
        $enc = ConvertFrom-SecureString (ConvertTo-SecureString $RefreshToken -AsPlainText -Force)
        @{ Token = $enc; Account = $AccountName; At = (Get-Date -Format "o") } |
            ConvertTo-Json | Set-Content -Path $script:MgTokenCacheFile -Encoding UTF8 -Force
        Write-Log "MgGraph token cache saved for $AccountName." "INFO" -FileOnly
    } catch {
        Write-Log "MgGraph: could not save token cache: $($_.Exception.Message)" "WARN" -FileOnly
    }
}

function Load-MgTokenCache {
    # Returns @{ RefreshToken; AccountName } or $null
    try {
        if (-not (Test-Path -LiteralPath $script:MgTokenCacheFile)) { return $null }
        $raw = Get-Content $script:MgTokenCacheFile -Raw | ConvertFrom-Json
        $dec = [System.Net.NetworkCredential]::new("", (ConvertTo-SecureString $raw.Token)).Password
        return @{ RefreshToken = $dec; AccountName = $raw.Account }
    } catch {
        Write-Log "MgGraph: could not load token cache: $($_.Exception.Message)" "WARN" -FileOnly
        return $null
    }
}

# --- FUNCTION: Update-MgGraphToken ---
# Uses the stored refresh token to silently obtain a fresh access token.
# Called automatically by Invoke-MgRestMethod before token expiry.
# Returns $true on success, $false if the refresh token is expired or revoked.
function Update-MgGraphToken {
    if (-not $global:MgRefreshToken) { return $false }
    try {
        $body = "client_id=$($script:MgClientId)" +
                "&grant_type=refresh_token" +
                "&refresh_token=$([uri]::EscapeDataString($global:MgRefreshToken))" +
                "&scope=$([uri]::EscapeDataString($script:MgScopes))"
        $resp = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/token" `
            -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        $global:MgAccessToken    = $resp.access_token
        $global:MgRefreshToken   = if ($resp.refresh_token) { $resp.refresh_token } else { $global:MgRefreshToken }
        $global:MgTokenExpiry    = (Get-Date).AddSeconds([int]$resp.expires_in)
        $global:MgGraphConnected = $true
        Write-Log "MgGraph access token refreshed silently." "INFO" -FileOnly
        return $true
    } catch {
        Write-Log "MgGraph token refresh failed: $($_.Exception.Message)" "WARN" -FileOnly
        $global:MgGraphConnected = $false
        return $false
    }
}

# --- FUNCTION: Invoke-MgRestMethod ---
# Authenticated Microsoft Graph REST API caller — no module dependency.
# Works on PowerShell 5.1 / .NET Framework 4.x.
# Auto-refreshes the access token when it is within 5 minutes of expiry.
# Throws a plain-text error message on API failure (caller wraps in try/catch).
function Invoke-MgRestMethod {
    param(
        [string]$Method,
        [string]$Endpoint,  # Path after /v1.0/ e.g. "users/john@co.com"
        $Body = $null        # Hashtable — serialised to JSON automatically
    )
    # Proactive refresh — don't wait for a 401 mid-operation
    if ($global:MgTokenExpiry -and (Get-Date) -gt $global:MgTokenExpiry.AddMinutes(-5)) {
        if (-not (Update-MgGraphToken)) {
            throw "Graph access token expired. Run option [23] again to re-authenticate."
        }
    }
    $uri     = "https://graph.microsoft.com/v1.0/$Endpoint"
    $headers = @{ Authorization = "Bearer $global:MgAccessToken"; "Content-Type" = "application/json" }
    $params  = @{ Method = $Method; Uri = $uri; Headers = $headers; ErrorAction = "Stop" }
    if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 5 -Compress) }
    try {
        return Invoke-RestMethod @params
    } catch {
        # In PS 5.1, ErrorDetails.Message contains the response body for HTTP errors
        $detail = ""
        try { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message } catch {}
        if (-not $detail) { $detail = $_.Exception.Message }
        throw $detail
    }
}

# --- FUNCTION: Ensure-MgGraphForLicensing ---
# PURPOSE: Establishes a Microsoft Graph connection using pure REST — no Microsoft.Graph
# module required. Compatible with PowerShell 5.1 / .NET Framework 4.x.
#
# WHY NO MODULE:
#   Microsoft.Graph v2.x requires .NET 6+. PowerShell 5.1 runs on .NET Framework 4.x.
#   Attempting to import Microsoft.Graph.Authentication v2.x on PS 5.1 throws:
#     "Method 'GetTokenAsync' does not have an implementation."
#   The REST API approach uses only built-in Invoke-RestMethod and needs no external modules.
#
# FLOW:
#   Stage 1 — Try silent reconnect: if a DPAPI-encrypted refresh token is on disk, use it
#              to obtain a new access token without any user interaction.
#   Stage 2 — Device code flow: display a code, poll until the user signs in at
#              https://microsoft.com/devicelogin, then save the new refresh token to disk.
#
# Returns $true if connected, $false on failure.
function Ensure-MgGraphForLicensing {

    if ($global:MgGraphConnected) {
        if ($global:MgTokenExpiry -and (Get-Date) -gt $global:MgTokenExpiry.AddMinutes(-5)) {
            return Update-MgGraphToken
        }
        return $true
    }

    Write-Host ""
    Write-Host "  ─── MICROSOFT GRAPH CONNECTION (licence group operations) ───────" -ForegroundColor Cyan
    Write-Host "  Using native REST API — no Microsoft.Graph module required." -ForegroundColor DarkGray
    Write-Host ""

    # ── STAGE 1: SILENT RECONNECT VIA CACHED REFRESH TOKEN ────────────────────
    Write-Host "  Checking for a saved sign-in..." -ForegroundColor DarkGray
    $cached = Load-MgTokenCache
    if ($cached) {
        $global:MgRefreshToken = $cached.RefreshToken
        if (Update-MgGraphToken) {
            Write-Host "  [✔] Resumed as $($cached.AccountName) (saved session)" -ForegroundColor Green
            Save-MgTokenCache -RefreshToken $global:MgRefreshToken -AccountName $cached.AccountName
            return $true
        }
        Write-Host "  Saved session expired — please sign in." -ForegroundColor DarkGray
    } else {
        Write-Host "  No saved session found." -ForegroundColor DarkGray
    }

    # ── STAGE 2: DEVICE CODE FLOW ─────────────────────────────────────────────
    $dcBody = "client_id=$($script:MgClientId)&scope=$([uri]::EscapeDataString($script:MgScopes))"
    try {
        $dcResp = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" `
            -Body $dcBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    } catch {
        Write-Host "  [✗] Could not start device code flow: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "MgGraph: device code init failed: $($_.Exception.Message)" "ERROR"
        return $false
    }

    Write-Host ""
    Write-Host "  $($dcResp.message)" -ForegroundColor Cyan
    Write-Host ""

    $tokenBody = "client_id=$($script:MgClientId)" +
                 "&grant_type=urn:ietf:params:oauth:grant-type:device_code" +
                 "&device_code=$([uri]::EscapeDataString($dcResp.device_code))"
    $deadline  = (Get-Date).AddSeconds([int]$dcResp.expires_in)
    $interval  = [int]$dcResp.interval
    $tokenResp = $null

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tokenResp = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/token" `
                -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
            break
        } catch {
            $errJson = $null; try { $errJson = ($_.ErrorDetails.Message | ConvertFrom-Json) } catch {}
            $errCode = if ($errJson -and $errJson.error) { $errJson.error } else { "" }
            if ($errCode -eq "authorization_pending") { continue }
            if ($errCode -in @("authorization_declined","access_denied")) {
                Write-Host "  [✗] Sign-in was declined." -ForegroundColor Red
                return $false
            }
            if ($errCode -eq "expired_token") {
                Write-Host "  [✗] Code expired. Run option [23] again to retry." -ForegroundColor Red
                return $false
            }
            # Transient error — keep polling
        }
    }

    if (-not $tokenResp -or -not $tokenResp.access_token) {
        Write-Host "  [✗] Sign-in timed out. Run option [23] again to retry." -ForegroundColor Red
        return $false
    }

    $global:MgAccessToken    = $tokenResp.access_token
    $global:MgRefreshToken   = $tokenResp.refresh_token
    $global:MgTokenExpiry    = (Get-Date).AddSeconds([int]$tokenResp.expires_in)
    $global:MgGraphConnected = $true

    # Decode the JWT payload (base64url) to extract the signed-in account name
    $accountName = "unknown"
    try {
        $payload = $tokenResp.access_token.Split('.')[1]
        $pad     = (4 - ($payload.Length % 4)) % 4; $payload += "=" * $pad
        $claims  = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
        $accountName = if ($claims.upn) { $claims.upn }
                       elseif ($claims.preferred_username) { $claims.preferred_username }
                       elseif ($claims.unique_name)        { $claims.unique_name }
                       else { "unknown" }
    } catch {}

    Write-Host "  [✔] Signed in as $accountName" -ForegroundColor Green
    Write-Host "      Sign-in saved — next session will reconnect automatically." -ForegroundColor DarkGray
    Write-Log "MgGraph: connected as $accountName (device code REST)." "INFO" -FileOnly

    if ($global:MgRefreshToken) {
        Save-MgTokenCache -RefreshToken $global:MgRefreshToken -AccountName $accountName
    }

    return $true
}

# --- FUNCTION: Perform-LicensedLeaverArchive ---
# PURPOSE: Processes a list of leavers whose OneDrive needs a temporary licence
# to become accessible. For each leaver (sequentially):
#
#   Step 1 — Resolve the user in Entra ID by their email address.
#   Step 2 — Add user to the licence Azure AD group → Plan 1 licence assigned.
#   Step 3 — Poll the OneDrive URL every 30 s (up to 30 min) until Get-PnPList
#             succeeds, confirming the site is accessible.
#   Step 4 — Run Perform-LeaverOneDriveArchive (enumerate libs → download →
#             verify → delete to Recycle Bin). Content lands in LocalDest\FolderName.
#   Step 5 — Remove user from the licence group immediately → slot freed.
#
# Only one licence is consumed at any time. If a step fails the user is still
# removed from the group (to keep slots free) and the failure is noted in the
# per-leaver summary at the end.
#
# PARAMETERS:
#   Leavers   : List[hashtable] — each entry: Email, SiteUrl, FolderName
#   LocalDest : Root local folder — each leaver gets a subfolder named FolderName
#   BufferDest: Temp buffer path for downloads
#   GroupId   : Object ID of the Azure AD licence group
#   GroupName : Display name of the group (for log/display)
function Perform-LicensedLeaverArchive {
    param(
        [System.Collections.Generic.List[hashtable]]$Leavers,
        [string]$LocalDest,
        [string]$BufferDest,
        [string]$GroupId,
        [string]$GroupName
    )

    $totalSw         = [System.Diagnostics.Stopwatch]::StartNew()
    $grandSuccess    = 0; $grandFailed = 0
    $results         = [System.Collections.Generic.List[hashtable]]::new()
    # Collect site URLs of successfully archived leavers for Recycle Bin link display at the end
    $archivedSiteUrls = [System.Collections.Generic.List[hashtable]]::new()

    Write-Log "=== LICENSED LEAVER ARCHIVE START: $($Leavers.Count) leavers | Group: $GroupName ===" "INFO"

    for ($li = 0; $li -lt $Leavers.Count; $li++) {

        if ($global:JobCancelled) {
            Write-Host "  [!] Run cancelled by user. Stopping." -ForegroundColor Yellow
            Write-Log "Licensed Leaver Archive: cancelled by user at leaver $($li+1)/$($Leavers.Count)." "WARN"
            break
        }

        $leaver          = $Leavers[$li]
        $email           = $leaver.Email
        $siteUrl         = $leaver.SiteUrl
        $label           = $leaver.FolderName
        $leaverLocalDest = Join-Path $LocalDest $label

        Write-Host ""
        Write-Host "  ████████████████████████████████████████████████████████████" -ForegroundColor Cyan
        Write-Host "  LEAVER $($li+1) of $($Leavers.Count) : $label" -ForegroundColor Cyan
        Write-Host "  Email   : $email" -ForegroundColor DarkGray
        Write-Host "  Site    : $($siteUrl -replace 'https://','')" -ForegroundColor DarkGray
        Write-Host "  Dest    : $leaverLocalDest" -ForegroundColor DarkGray
        Write-Host "  ████████████████████████████████████████████████████████████" -ForegroundColor Cyan
        Write-Log "LEAVER $($li+1)/$($Leavers.Count): $label | $email | $siteUrl" "INFO"

        # ── STEP 1: RESOLVE USER IN ENTRA ID ──────────────────────────────────
        Write-Host ""
        Write-Host "  [1/5] Resolving user in Entra ID..." -ForegroundColor Magenta
        $mgUser = $null
        try {
            $mgUser = Invoke-MgRestMethod -Method "GET" -Endpoint "users/$([uri]::EscapeDataString($email))"
            Write-Host "  [✔] Found: $($mgUser.displayName) (ID: $($mgUser.id))" -ForegroundColor Green
            Write-Log "  Resolved: $($mgUser.displayName) | ID: $($mgUser.id)" "INFO" -FileOnly
        } catch {
            Write-Host "  [✗] '$email' not found in Entra ID: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "  SKIPPED: $label — user not found: $($_.Exception.Message)" "ERROR"
            $results.Add(@{ Label = $label; Status = "SKIPPED — user not found in Entra ID"; Level = "ERROR" })
            $grandFailed++
            continue
        }

        # ── STEP 2: ADD USER TO LICENCE GROUP ─────────────────────────────────
        Write-Host ""
        Write-Host "  [2/5] Adding $($mgUser.displayName) to [$GroupName]..." -ForegroundColor Magenta
        $addedToGroup   = $false
        $alreadyInGroup = $false
        try {
            $memberBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($mgUser.id)" }
            Invoke-MgRestMethod -Method "POST" -Endpoint "groups/$GroupId/members/`$ref" -Body $memberBody | Out-Null
            $addedToGroup = $true
            Write-Host "  [✔] Added to group. Licence assignment now in progress..." -ForegroundColor Green
            Write-Log "  ADDED to [$GroupName]: $($mgUser.displayName)" "SUCCESS" -FileOnly
        } catch {
            $addErr = $_.Exception.Message
            if ($addErr -match "(?i)already exist|already a member|exists|already") {
                Write-Host "  [i] Already in [$GroupName] — licence may already be active." -ForegroundColor DarkGray
                Write-Log "  Already in group (skipped add): $($mgUser.displayName)" "INFO" -FileOnly
                $alreadyInGroup = $true
                $addedToGroup   = $true   # treat as added — proceed to polling
            } else {
                Write-Host "  [✗] Could not add to group: $addErr" -ForegroundColor Red
                Write-Log "  FAILED to add $($mgUser.displayName) to group: $addErr" "ERROR"
                $results.Add(@{ Label = $label; Status = "FAILED — could not add to licence group"; Level = "ERROR" })
                $grandFailed++
                continue
            }
        }

        # ── STEP 3: POLL UNTIL ONEDRIVE IS ACCESSIBLE ─────────────────────────
        Write-Host ""
        Write-Host "  [3/5] Waiting for OneDrive to become accessible..." -ForegroundColor Magenta
        if (-not $alreadyInGroup) {
            Write-Host "  Licence propagation through SharePoint typically takes 2–15 minutes." -ForegroundColor DarkGray
            Write-Host "  Checking every 30 s — maximum wait 30 minutes." -ForegroundColor DarkGray
        }

        $pollSw       = [System.Diagnostics.Stopwatch]::StartNew()
        $maxPollMs    = 30 * 60 * 1000   # 30 minutes
        $pollInterval = 30               # seconds between attempts
        $accessible   = $false
        $firstPoll    = $true

        while ($pollSw.ElapsedMilliseconds -lt $maxPollMs) {
            if ($global:JobCancelled) { break }

            if (-not $firstPoll) {
                $remaining = [int](($maxPollMs - $pollSw.ElapsedMilliseconds) / 1000)
                Write-Host ("  [{0,4}s] Not accessible yet — retrying in ${pollInterval}s...  (${remaining}s remaining)" -f [int]$pollSw.Elapsed.TotalSeconds) -ForegroundColor DarkGray
                Start-InterruptibleSleep -Seconds $pollInterval -Label "Waiting for OneDrive licence activation"
                if ($global:JobCancelled) { break }
            }
            $firstPoll = $false

            try {
                Connect-PnPOnline -Url $siteUrl -UseWebLogin -ErrorAction Stop
                $null = Get-PnPList -ErrorAction Stop   # light call to confirm the site is readable
                $accessible = $true
                Write-Host ("  [✔] OneDrive accessible after {0}s — proceeding." -f [int]$pollSw.Elapsed.TotalSeconds) -ForegroundColor Green
                Write-Log "  OneDrive accessible at $([int]$pollSw.Elapsed.TotalSeconds)s: $siteUrl" "SUCCESS" -FileOnly
                break
            } catch {
                Write-Log "  Poll $([int]$pollSw.Elapsed.TotalSeconds)s — not accessible: $($_.Exception.Message)" "INFO" -FileOnly
            }
        }
        $pollSw.Stop()

        if ($global:JobCancelled) { break }

        if (-not $accessible) {
            Write-Host ""
            Write-Host "  [✗] OneDrive did not become accessible within 30 minutes." -ForegroundColor Red
            Write-Host "      Possible causes: licence propagation delay, SharePoint Admin throttle," -ForegroundColor DarkGray
            Write-Host "      or the account is fully deprovisioned. Check Azure AD licensing status." -ForegroundColor DarkGray
            Write-Log "  FAILED: OneDrive not accessible after $([int]$pollSw.Elapsed.TotalSeconds)s — $label" "ERROR"

            # Remove from group even on failure so the licence slot is not held
            if (-not $alreadyInGroup) {
                try {
                    Invoke-MgRestMethod -Method "DELETE" -Endpoint "groups/$GroupId/members/$($mgUser.id)/$ref" | Out-Null
                    Write-Host "  [✔] Removed from licence group (slot freed)." -ForegroundColor DarkGray
                    Write-Log "  Removed $($mgUser.displayName) from group after poll timeout." "INFO" -FileOnly
                } catch {
                    Write-Host "  [!] Could not remove from group: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Log "  WARNING: Could not remove $($mgUser.displayName) from group: $($_.Exception.Message)" "WARN"
                }
            }

            $results.Add(@{ Label = $label; Status = "FAILED — OneDrive not accessible after 30 min"; Level = "ERROR" })
            $grandFailed++
            continue
        }

        # ── STEP 4: ARCHIVE THE ONEDRIVE ──────────────────────────────────────
        Write-Host ""
        Write-Host "  [4/5] Archiving OneDrive content..." -ForegroundColor Magenta

        if (-not (Test-Path -LiteralPath $leaverLocalDest)) {
            New-Item -ItemType Directory -Path $leaverLocalDest -Force | Out-Null
        }

        $archiveResult = $null
        try {
            $archiveResult = Perform-LeaverOneDriveArchive `
                -LeaverSiteUrl $siteUrl `
                -LocalDest     $leaverLocalDest `
                -BufferDest    $BufferDest
        } catch {
            Write-Host "  [✗] Archive engine error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "  Archive engine exception for $label : $($_.Exception.Message)" "ERROR"
            $archiveResult = @{ Success = 0; Failed = 1 }
        }

        # ── STEP 5: REMOVE FROM LICENCE GROUP ─────────────────────────────────
        Write-Host ""
        Write-Host "  [5/5] Removing $($mgUser.displayName) from [$GroupName] (freeing licence slot)..." -ForegroundColor Magenta
        if (-not $alreadyInGroup) {
            try {
                Invoke-MgRestMethod -Method "DELETE" -Endpoint "groups/$GroupId/members/$($mgUser.id)/$ref" | Out-Null
                Write-Host "  [✔] Removed from group. Licence slot is now free for the next leaver." -ForegroundColor Green
                Write-Log "  REMOVED $($mgUser.displayName) from [$GroupName] (archive complete)." "SUCCESS" -FileOnly
            } catch {
                $remErr = $_.Exception.Message
                if ($remErr -match "(?i)not a member|ResourceNotFound|does not exist") {
                    Write-Host "  [i] User was already removed from the group." -ForegroundColor DarkGray
                } else {
                    Write-Host "  [!] Could not remove from group: $remErr" -ForegroundColor Yellow
                    Write-Log "  WARNING: Could not remove $($mgUser.displayName) from group: $remErr" "WARN"
                }
            }
        } else {
            Write-Host "  [i] Leaver was pre-existing in the group — not removing (not added by this run)." -ForegroundColor DarkGray
        }

        # ── RECORD OUTCOME ─────────────────────────────────────────────────────
        $archiveOk   = if ($archiveResult) { $archiveResult.Success } else { 0 }
        $archiveErr  = if ($archiveResult) { $archiveResult.Failed  } else { 1 }

        if ($archiveErr -eq 0 -and $archiveOk -gt 0) {
            $results.Add(@{ Label = $label; Status = "ARCHIVED ($archiveOk lib$(if($archiveOk -ne 1){'s'}) downloaded + deleted)"; Level = "SUCCESS" })
            $archivedSiteUrls.Add(@{ Label = $label; Url = $siteUrl })
            $grandSuccess++
        } elseif ($archiveOk -gt 0) {
            $results.Add(@{ Label = $label; Status = "PARTIAL ($archiveOk OK, $archiveErr failed — see log)"; Level = "WARN" })
            $grandFailed++
        } else {
            $results.Add(@{ Label = $label; Status = "FAILED (archive error — check log for details)"; Level = "ERROR" })
            $grandFailed++
        }

        Write-Log "LEAVER $($li+1)/$($Leavers.Count) DONE: $label | Archived=$archiveOk Failed=$archiveErr" "INFO" -FileOnly
    }

    $totalSw.Stop()

    # ── GRAND SUMMARY ──────────────────────────────────────────────────────────
    Write-Host ""
    Write-SummaryBox -Title "UNLICENSED LEAVER ARCHIVE — COMPLETE" -Lines @(
        @{ Label = "Leavers Queued";        Value = "$($Leavers.Count)";   Level = "INFO" }
        @{ Label = "Successfully Archived"; Value = "$grandSuccess";        Level = if($grandSuccess -gt 0){"SUCCESS"}else{"WARN"} }
        @{ Label = "Failed / Errors";       Value = "$grandFailed";         Level = if($grandFailed  -gt 0){"ERROR"}else{"SUCCESS"} }
        @{ Label = "Licence Group";         Value = $GroupName;              Level = "INFO" }
        @{ Label = "Total Time";            Value = "$([math]::Floor($totalSw.Elapsed.TotalHours))h $($totalSw.Elapsed.Minutes)m $($totalSw.Elapsed.Seconds)s"; Level = "INFO" }
    ) -BoxColor "Cyan"

    Write-Host ""
    Write-Host "  ┌─ PER-LEAVER OUTCOMES ──────────────────────────────────────" -ForegroundColor Cyan
    foreach ($r in $results) {
        $col = switch ($r.Level) { "SUCCESS"{"Green"} "ERROR"{"Red"} "WARN"{"Yellow"} default{"Cyan"} }
        $lbl = if ($r.Label.Length -gt 34) { $r.Label.Substring(0,31) + "..." } else { $r.Label.PadRight(34) }
        Write-Host "  │  $lbl  $($r.Status)" -ForegroundColor $col
    }
    Write-Host "  └─────────────────────────────────────────────────────────────" -ForegroundColor Cyan

    # Show Recycle Bin links for all successfully archived leavers — matching [20]/[21] behaviour
    if ($archivedSiteUrls.Count -gt 0) {
        $rbLinks = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($entry in $archivedSiteUrls) {
            $rbUrl = Get-LeaverRecycleBinUrl -LeaverUrl $entry.Url
            if ($rbUrl) { $rbLinks.Add(@{ Label = $entry.Label; Url = $rbUrl }) }
        }
        if ($rbLinks.Count -gt 0) {
            Write-RecycleBinLinks -Links $rbLinks
        }
    }

    Write-Log "=== LICENSED LEAVER ARCHIVE END: Success=$grandSuccess Failed=$grandFailed Time=$([int]$totalSw.Elapsed.TotalSeconds)s ===" "INFO"
}


# ========================================================================================
# 5. EXECUTION BLOCK & MAIN MENU (The Control Room)
# ========================================================================================

while ($true) {
    $favCount   = @(Get-Favorites).Count
    $favLine    = if ($favCount -gt 0) { "$favCount Favourite(s) saved" } else { "No Favourites yet" }

    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║         ROBUST HYBRID FILE MANAGER  (v111)              ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host "  │ User      : $ActiveUser" -ForegroundColor DarkGray
    Write-Host "  │ Favourites: $favLine  (type [F] at any prompt to use)" -ForegroundColor DarkGray
    Write-Host "  │ Log file  : $LogFile" -ForegroundColor DarkGray
    Write-Host "  └──────────────────────────────────────────────────────────" -ForegroundColor DarkGray

    Write-Host ""
    # Rotate tips — show a different tip each time the menu loads
    if (-not $script:TipIndex) { $script:TipIndex = 0 }
    $menuTips = @(
        "  GETTING A URL: Open SharePoint in your browser → navigate to the folder → copy the address bar.",
        "  FAVOURITES [F]: Save any URL or path as a Favourite to recall it instantly at any prompt.",
        "  SMART RESUME:   If a job is interrupted, re-run with the same paths — completed files are skipped.",
        "  PAUSE / CANCEL: Press [P] to pause a running job, [C] to cancel it cleanly at any time.",
        "  OPEN LOGS [L]:  Press L at the main menu to open your log folder and review audit history."
    )
    $tip = $menuTips[$script:TipIndex % $menuTips.Count]
    $script:TipIndex++
    Write-Host "  ┌─  TIP  ────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "  $tip" -ForegroundColor Gray
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  ┌─  NOT SURE WHERE TO START?  ──────────────────────────────" -ForegroundColor White
    Write-Host "  │  Moving files inside SharePoint / OneDrive?      → [1] [2] [3]" -ForegroundColor DarkGray
    Write-Host "  │  Copy or move files between folders on your PC?  → [7] [8] [9]" -ForegroundColor DarkGray
    Write-Host "  │  Network share (\\server\share) as source/dest?   → All local options support UNC" -ForegroundColor DarkGray
    Write-Host "  │  Download SharePoint files to your PC?           → [13] [14]" -ForegroundColor DarkGray
    Write-Host "  │  Confirm a copy completed correctly?             → [5] [11]" -ForegroundColor DarkGray
    Write-Host "  │  Find a file you cannot locate?                  → [6] [12]" -ForegroundColor DarkGray
    Write-Host "  │  Upload local files to SharePoint?              → [15] [16]" -ForegroundColor DarkGray
    Write-Host "  │  Transfer between 2 SharePoint sites?           → Download [13/14] then Upload [15/16]" -ForegroundColor DarkGray
    Write-Host "  │  Archive cloud folders + reclaim storage?        → [17] [18]" -ForegroundColor DarkGray
    Write-Host "  │  Archive multiple cloud sources in one run?      → [19]" -ForegroundColor DarkGray
    Write-Host "  │  Archive a LEAVER's OneDrive (disabled account)? → [20]
  │  Archive MULTIPLE leaver OneDrives unattended?   → [21]
  │  Get Recycle Bin links after a leaver archive?   → [22]
  │  Leaver's licence was removed — OneDrive frozen? → [23]" -ForegroundColor DarkYellow
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor White

    Write-Host ""
    Write-Host "  ┌─  A │ CLOUD OPERATIONS  ── same SharePoint site only ────" -ForegroundColor Yellow
    Write-Host "  │  [1] CLOUD COPY    Copy a folder within the same SharePoint site." -ForegroundColor Yellow
    Write-Host "  │  [2] CLOUD MOVE    Move a folder within the same SharePoint site." -ForegroundColor Yellow
    Write-Host "  │  [3] CLOUD BATCH   Pick specific items to Copy / Move / Delete / Verify." -ForegroundColor Yellow
    Write-Host "  │  [4] CLOUD WIPE    Delete a cloud folder (sent to Recycle Bin)." -ForegroundColor Yellow
    Write-Host "  │  [5] CLOUD VERIFY  Prove two cloud folders are byte-for-byte identical." -ForegroundColor Yellow
    Write-Host "  │  [6] CLOUD SEARCH  Find any file or folder on SharePoint." -ForegroundColor Yellow
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor Yellow

    Write-Host ""
    Write-Host "  ┌─  B │ LOCAL OPERATIONS  ── PC, mapped drives, UNC paths ─" -ForegroundColor Green
    Write-Host "  │  [7] LOCAL COPY    Copy a folder to another location on your PC / network." -ForegroundColor Green
    Write-Host "  │  [8] LOCAL MOVE    Move a folder to another location on your PC / network." -ForegroundColor Green
    Write-Host "  │  [9] LOCAL BATCH   Pick specific items to Copy / Move / Delete / Verify." -ForegroundColor Green
    Write-Host "  │ [10] LOCAL WIPE    Permanently delete a local folder (skips Recycle Bin!)." -ForegroundColor Green
    Write-Host "  │ [11] LOCAL VERIFY  Prove two local folders are byte-for-byte identical." -ForegroundColor Green
    Write-Host "  │ [12] LOCAL SEARCH  Find any file or folder on your PC or network share." -ForegroundColor Green
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor Green

    Write-Host ""
    Write-Host "  ┌─  C │ DIRECT DOWNLOAD  ── SharePoint → your PC / UNC ────" -ForegroundColor Magenta
    Write-Host "  │ [13] DIRECT DOWNLOAD  Download an entire SharePoint folder to your hard drive." -ForegroundColor Magenta
    Write-Host "  │ [14] BATCH DOWNLOAD   Pick specific SharePoint items to download." -ForegroundColor Magenta
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor Magenta

    Write-Host ""
    Write-Host "  ┌─  D │ UPLOAD OPERATIONS  ── Your PC / UNC → SharePoint ───" -ForegroundColor Cyan
    Write-Host "  │ [15] DIRECT UPLOAD   Upload an entire local folder to SharePoint." -ForegroundColor Cyan
    Write-Host "  │ [16] BATCH UPLOAD    Pick specific local items to upload to SharePoint." -ForegroundColor Cyan
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "  ┌─  E │ ARCHIVE OPERATIONS  ── Download + Verify + Delete ──" -ForegroundColor Red
    Write-Host "  │ [17] ARCHIVE DOWNLOAD  Download → verify → auto-delete one cloud folder." -ForegroundColor Red
    Write-Host "  │ [18] BATCH ARCHIVE     Pick items from a folder to archive in sequence." -ForegroundColor Red
    Write-Host "  │ [19] MULTI-SOURCE ARCHIVE  Archive multiple cloud sources to one local folder." -ForegroundColor Red
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor Red

    Write-Host ""
    Write-Host "  ┌─  F │ LEAVER OPERATIONS  ── Disabled account OneDrives ──" -ForegroundColor DarkYellow
    Write-Host "  │ [20] LEAVER ONEDRIVE ARCHIVE  Archive a single leaver's OneDrive." -ForegroundColor DarkYellow
    Write-Host "  │      Connect → enumerate Document Libraries → batch archive each." -ForegroundColor DarkGray
    Write-Host "  │ [21] MULTI-SOURCE LEAVER ARCHIVE  Queue multiple leavers — run all unattended." -ForegroundColor DarkYellow
    Write-Host "  │      Same archive engine as [20] batched across N leavers in one run." -ForegroundColor DarkGray
    Write-Host "  │ [22] LEAVER RECYCLE BIN LINKS  Generate Recycle Bin URLs from leaver links." -ForegroundColor Cyan
    Write-Host "  │      Paste URLs → get browser links to open and empty each Recycle Bin." -ForegroundColor DarkGray
    Write-Host "  │ [23] UNLICENSED LEAVER ARCHIVE  For leavers whose E5 licence was removed." -ForegroundColor Yellow
    Write-Host "  │      Assigns a Plan 1 licence via Azure AD group → archives → removes licence." -ForegroundColor DarkGray
    Write-Host "  │      Paste 3 or 6-column list (email / URL / folder). [V] clipboard paste." -ForegroundColor DarkGray
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor DarkYellow

    Write-Host ""
    Write-Host "  ┌─  UTILITIES  ──────────────────────────────────────────────" -ForegroundColor White
    Write-Host "  │  [F] FAVOURITES  Save / recall Cloud URLs and local paths." -ForegroundColor Cyan
    Write-Host "  │  [H] HELP        Quick reference card — hotkeys, syntax, and tips." -ForegroundColor White
    Write-Host "  │  [L] OPEN LOGS   Open the audit log folder in File Explorer." -ForegroundColor White
    Write-Host "  │  [Q] EXIT    Close the script." -ForegroundColor White
    Write-Host "  └───────────────────────────────────────────────────────────" -ForegroundColor White

    Write-Host ""
    Write-Host "  During any job:  [P] Pause   [C] Cancel   [F] Favourites   [R] Repeat last path" -ForegroundColor DarkGray
    Write-Host ""

    $choice = Read-Host "  Enter option number"
    Write-Log "[USER INPUT] Main Menu Selection: '$choice'" "INFO" -FileOnly
    Write-Host ""

    $global:JobCancelled = $false
    Clear-KeyBuffer

    if ($choice -match '^[Ff]$') {
        Manage-Favorites
        continue
    }

    if ($choice -match '^[Ll]$') {
        try {
            Invoke-Item $LogFileDirectory
            Write-Host "  ✔ Opened log folder: $LogFileDirectory" -ForegroundColor Green
            Write-Host "    Log file: HybridJobLog_$($ActiveUser).log" -ForegroundColor DarkGray
        } catch { Write-Host "  ✘ Could not open log folder automatically. Path: $LogFileDirectory" -ForegroundColor Red }
        Write-Host ""
        Write-Host "  Press ENTER to return to the main menu..." -ForegroundColor DarkGray
        $null = Read-Host
        continue
    }

    if ($choice -match '^[Hh]$') {
        Write-Host ""
        Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor White
        Write-Host "  ║                  QUICK REFERENCE CARD                   ║" -ForegroundColor White
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  PROMPT HOTKEYS                                          ║" -ForegroundColor White
        Write-Host "  ║  [F]  Open Favourites Picker for this prompt             ║" -ForegroundColor Cyan
        Write-Host "  ║  [R]  Repeat the last value used for this prompt type    ║" -ForegroundColor Cyan
        Write-Host "  ║  [Q]  Cancel and return to the previous menu             ║" -ForegroundColor Cyan
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  DURING AN ACTIVE JOB                                    ║" -ForegroundColor White
        Write-Host "  ║  [P]  Pause the job (press any key to resume, C to stop) ║" -ForegroundColor Yellow
        Write-Host "  ║  [C]  Cancel the job immediately                         ║" -ForegroundColor Yellow
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  BATCH SELECTION SYNTAX                                  ║" -ForegroundColor White
        Write-Host "  ║  3        Process only item 3                            ║" -ForegroundColor Green
        Write-Host "  ║  2-5      Process items 2, 3, 4, 5 (inclusive range)     ║" -ForegroundColor Green
        Write-Host "  ║  1,4,7    Process items 1, 4, and 7                      ║" -ForegroundColor Green
        Write-Host "  ║  1-3,6-8  Process items 1-3 and 6-8 combined             ║" -ForegroundColor Green
        Write-Host "  ║  10-2     Backwards range — same as 2-10, auto-sorted    ║" -ForegroundColor Green
        Write-Host "  ║  A        Process ALL items in the list                  ║" -ForegroundColor Green
        Write-Host "  ║  S        Search within the current folder               ║" -ForegroundColor Green
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  SEARCH WILDCARDS                                        ║" -ForegroundColor White
        Write-Host "  ║  *.pdf          All PDF files                            ║" -ForegroundColor Magenta
        Write-Host "  ║  *invoice*      Any name containing 'invoice'            ║" -ForegroundColor Magenta
        Write-Host "  ║  *2024*         Any name containing '2024'               ║" -ForegroundColor Magenta
        Write-Host "  ║  Payslip.docx   Exact file name match                    ║" -ForegroundColor Magenta
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  PATH FORMATS SUPPORTED                                  ║" -ForegroundColor White
        Write-Host "  ║  C:\Folder\Subfolder        Local drive path             ║" -ForegroundColor DarkGray
        Write-Host "  ║  D:\Archives\2024           Different local drive        ║" -ForegroundColor DarkGray
        Write-Host "  ║  \\Server01\Share\Folder    UNC network path             ║" -ForegroundColor DarkGray
        Write-Host "  ║  \\nas.local\data\archive   NAS device UNC path          ║" -ForegroundColor DarkGray
        Write-Host "  ║  (Paths can be pasted with quotes — they are stripped)   ║" -ForegroundColor DarkGray
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  VERIFY RESULT CODES                                     ║" -ForegroundColor White
        Write-Host "  ║  Perfectly Synced   File identical in both locations     ║" -ForegroundColor Green
        Write-Host "  ║  Mismatch           File exists but size/date differs    ║" -ForegroundColor Red
        Write-Host "  ║  Missing in Dest    File not found in destination        ║" -ForegroundColor Red
        Write-Host "  ║  Orphaned in Dest   Extra file in dest not in source     ║" -ForegroundColor Yellow
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  ARCHIVE OPERATIONS (Download + Verify + Auto-Delete)    ║" -ForegroundColor White
        Write-Host "  ║  [17] Archive Download — DL → verify → auto-delete       ║" -ForegroundColor Yellow
        Write-Host "  ║  [18] Batch Archive    — pick items, DL→verify→auto-del  ║" -ForegroundColor Yellow
        Write-Host "  ║  [19] Multi-Source     — paste N URLs, archive + delete all  ║" -ForegroundColor Yellow
        Write-Host "  ║  Delete blocked if: DL failure OR count/bytes mismatch   ║" -ForegroundColor Yellow
        Write-Host "  ║  Deletion = SharePoint Recycle Bin (~93 days recoverable) ║" -ForegroundColor Yellow
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  LEAVER OPERATIONS  [20] [21] [22] [23]                 ║" -ForegroundColor DarkYellow
        Write-Host "  ║  [20] Single leaver: connect → enumerate libs → archive ║" -ForegroundColor DarkYellow
        Write-Host "  ║  [21] Multi-leaver: queue N leavers, run all unattended  ║" -ForegroundColor DarkYellow
        Write-Host "  ║  [22] Recycle Bin links: paste leaver URL(s) → get       ║" -ForegroundColor Cyan
        Write-Host "  ║       browser links to open & empty each Recycle Bin.    ║" -ForegroundColor Cyan
        Write-Host "  ║  [23] Unlicensed leavers: assign Plan 1 licence via AAD  ║" -ForegroundColor Yellow
        Write-Host "  ║       group → archive → remove from group (1 at a time)  ║" -ForegroundColor Yellow
        Write-Host "  ║  [20]/[21]/[23] show Recycle Bin links at end of run.   ║" -ForegroundColor DarkGray
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  UPLOAD OPERATIONS & CROSS-SITE BRIDGE                  ║" -ForegroundColor White
        Write-Host "  ║  [15] Direct Upload — entire local folder → SharePoint   ║" -ForegroundColor Cyan
        Write-Host "  ║  [16] Batch Upload  — pick items → SharePoint            ║" -ForegroundColor Cyan
        Write-Host "  ║  Cross-site bridge: [13/14] download from Site A,        ║" -ForegroundColor Cyan
        Write-Host "  ║  then [15/16] upload to Site B (bypasses cloud policy)   ║" -ForegroundColor Cyan
        Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor White
        Write-Host "  ║  FAVOURITES MANAGER  [F] from Main Menu                  ║" -ForegroundColor White
        Write-Host "  ║  [A] Add   [R] Rename   [E] Edit Value   [D] Delete      ║" -ForegroundColor Cyan
        Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor White
        Write-Host ""
        Read-Host "  Press ENTER to return to the main menu" | Out-Null
        continue
    }

    if ([string]::IsNullOrWhiteSpace($choice) -or $choice -notmatch '^(23|22|21|20|19|1[0-8]|[1-9]|[QqFfHhLl])$') {
        Write-Host "  Not recognised. Type a number (1-22), [F], [H], [L], or [Q]." -ForegroundColor Red
        continue
    }

    # ============================================================================================
    # SECTION A - CLOUD OPERATIONS
    # ============================================================================================

    if ($choice -in @('1','2')) {
        $actionWord = if ($choice -eq '1') { "COPY" } else { "MOVE" }
        $actionVerb = if ($choice -eq '1') { "copy" } else { "move" }

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  CLOUD $actionWord  (Server-Side - uses 0 bytes of your hard drive)" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Tells Microsoft's servers to $actionVerb the folder for you - entirely in the cloud."
        Write-Host "   Your computer is NOT used to transfer the data, so it is near-instant regardless"
        Write-Host "   of folder size and uses zero bytes of your hard drive."
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS" -ForegroundColor White
        Write-Host "   Use this when BOTH the source and destination are inside the SAME SharePoint site"
        Write-Host "   or the SAME OneDrive account. For example: reorganising folders within your"
        Write-Host "   same SharePoint site, or tidying up your own OneDrive."
        Write-Host ""
        Write-Host "  ! SAME SITE ONLY" -ForegroundColor Red
        Write-Host "   This CANNOT $actionVerb data between two different companies or different SharePoint"
        Write-Host "   tenants. For cross-site transfers, use option [7] or [8] instead."
        Write-Host ""
        Write-Host "  ▸ HOW TO GET THE URL" -ForegroundColor White
        Write-Host "   1. Open SharePoint or OneDrive in your web browser (Edge, Chrome, etc.)."
        Write-Host "   2. Navigate into the folder you want to $actionVerb."
        Write-Host "   3. Copy the full URL from the address bar at the top of the browser."
        Write-Host "   4. Paste it below. The script strips tracking codes automatically."
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SOURCE folder URL (the folder you want to $actionVerb FROM)" -Type "Cloud" -LogLabel "Source URL"
            if ($inputSrc -eq "Q") { Write-Log "Cloud $actionWord cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the DESTINATION folder URL (the folder you want to $actionVerb INTO)" -Type "Cloud" -LogLabel "Destination URL"
            if ($inputDst -eq "Q") { Write-Log "Cloud $actionWord cancelled by user." "INFO"; break }

            $srcParsed = Parse-CloudUrl -Url $inputSrc.Trim()
            $dstParsed = Parse-CloudUrl -Url $inputDst.Trim()

            if (-not $srcParsed -or -not $dstParsed) {
                Write-Host " One of the URLs could not be understood. Make sure you copy the full URL from the browser address bar." -ForegroundColor Red
                continue   # stay in the loop; let user try again
            }

            Write-Host ""
            Write-Host " Starting... A browser login window may appear. Sign in with your Microsoft account." -ForegroundColor DarkGray

            if ($choice -eq '1') { Perform-CloudTransfer -SrcData $srcParsed -DstData $dstParsed -IsMove:$false }
            if ($choice -eq '2') { Perform-CloudTransfer -SrcData $srcParsed -DstData $dstParsed -IsMove:$true  }
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Operation complete. Enter another SOURCE URL to continue," -ForegroundColor DarkGray
            Write-Host "  or press [Q] at the next prompt to return to the main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '3') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  CLOUD BATCH - Copy, Move, Delete, or Verify by Index" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Opens the top-level contents of a SharePoint folder as a numbered list."
        Write-Host "   You type the numbers of the items you want to process (e.g. '1, 3-5, 8')"
        Write-Host "   and choose an action: Copy, Move, Delete, or Verify."
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS" -ForegroundColor White
        Write-Host "   Use this when you don't want to process an entire folder at once - for example,"
        Write-Host "   when you only need to move specific sub-folders from a large archive."
        Write-Host ""
        Write-Host "  ! SAME-SITE RULE" -ForegroundColor Red
        Write-Host "   Both source and destination must be on the same SharePoint site."
        Write-Host "   For cross-site transfers use option [9] instead."
        Write-Host ""
        Write-Host " DESTINATION URL: Only needed for Copy, Move, or Verify. Leave blank for Delete." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SOURCE folder URL (the folder whose contents you want to manage)" -Type "Cloud" -LogLabel "Source URL"
            if ($inputSrc -eq "Q") { Write-Log "Cloud Batch cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the DESTINATION folder URL (press ENTER to skip - only needed for Copy, Move, Verify)" -Type "Cloud" -LogLabel "Destination URL" -AllowEmpty
            if ($inputDst -eq "Q") { Write-Log "Cloud Batch cancelled by user." "INFO"; break }

            $srcParsed = Parse-CloudUrl -Url $inputSrc.Trim()
            $dstParsed = if ([string]::IsNullOrWhiteSpace($inputDst)) { $null } else { Parse-CloudUrl -Url $inputDst.Trim() }

            if (-not $srcParsed) {
                Write-Host " The source URL could not be understood. Please copy the full URL from the browser address bar." -ForegroundColor Red
                continue
            }
            Perform-CloudBatch -SrcData $srcParsed -DstData $dstParsed
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another SOURCE URL to continue, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '4') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  CLOUD WIPE - Delete an Entire Cloud Folder" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Deletes an entire SharePoint folder and everything inside it."
        Write-Host "   The data is sent to the cloud Recycle Bin - it is NOT immediately permanent."
        Write-Host "   Items remain in the Recycle Bin for approximately 93 days before being purged."
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS" -ForegroundColor White
        Write-Host "   Use this to clean up old or unwanted folders on SharePoint after confirming"
        Write-Host "   all data has been successfully transferred and verified."
        Write-Host ""
        Write-Host "  ! SAFETY GATE: You will need to type 'YES' in capitals to confirm before" -ForegroundColor Red
        Write-Host " anything is deleted. Type anything else to cancel safely." -ForegroundColor Red
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the URL of the SharePoint FOLDER you want to delete" -Type "Cloud" -LogLabel "Wipe URL"
            if ($inputSrc -eq "Q") { Write-Log "Cloud Wipe cancelled by user." "INFO"; break }

            $srcParsed = Parse-CloudUrl -Url $inputSrc.Trim()
            if (-not $srcParsed) {
                Write-Host " The URL could not be understood. Please copy the full URL from the browser address bar." -ForegroundColor Red
                continue
            }
            Perform-CloudWipe -TargetData $srcParsed
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another URL to delete, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '5') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  CLOUD VERIFY - Prove Two Cloud Folders Are Identical" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Scans both a source and destination folder on SharePoint and compares every"
        Write-Host "   single file without downloading anything. For each file it checks:"
        Write-Host "     - Does the file exist in both locations?     (missing = a failed transfer)"
        Write-Host "     - Is the file size in bytes EXACTLY the same? (mismatch = possible corruption)"
        Write-Host "   At the end you get a full report: Synced / Mismatched / Missing / Orphaned."
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS" -ForegroundColor White
        Write-Host "   Always run this after a large Cloud Copy or Move to confirm 100% success."
        Write-Host "   It is the only way to be mathematically certain no files were missed."
        Write-Host ""
        Write-Host "  ▸ RESULT GUIDE" -ForegroundColor White
        Write-Host "   Perfectly Synced  = file in both places, byte-for-byte identical. Good." -ForegroundColor DarkGray
        Write-Host "   Mismatch          = file exists in both but sizes differ. Re-copy recommended." -ForegroundColor DarkGray
        Write-Host "   Missing in Dest   = file is in source but missing from destination. Copy failed." -ForegroundColor DarkGray
        Write-Host "   Orphaned in Dest  = file in destination but NOT in source. Unexpected extra file." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SOURCE folder URL (the original / master copy)" -Type "Cloud" -LogLabel "Source URL"
            if ($inputSrc -eq "Q") { Write-Log "Cloud Verify cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the DESTINATION folder URL (the copy you want to check)" -Type "Cloud" -LogLabel "Destination URL"
            if ($inputDst -eq "Q") { Write-Log "Cloud Verify cancelled by user." "INFO"; break }

            $srcParsed = Parse-CloudUrl -Url $inputSrc.Trim()
            $dstParsed = Parse-CloudUrl -Url $inputDst.Trim()

            if (-not $srcParsed -or -not $dstParsed) {
                Write-Host " One of the URLs could not be understood. Please copy the full URL from the browser address bar." -ForegroundColor Red
                continue
            }

            Perform-CloudVerify -SrcData $srcParsed -DstData $dstParsed
            try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another SOURCE URL to compare, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '6') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  CLOUD SEARCH - Find Files on SharePoint" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Recursively scans every sub-folder inside a SharePoint directory and returns"
        Write-Host "   a list of every file or folder whose name matches your search term."
        Write-Host ""
        Write-Host "  ▸ HOW TO SEARCH" -ForegroundColor White
        Write-Host "   Exact name:   Type the full file name,   e.g.   Annual Report.docx"
        Write-Host "   By extension: Type a wildcard,           e.g.   *.pdf   or   *.xlsx"
        Write-Host "   Partial name: Use a wildcard,            e.g.   *invoice*   or   *2024*"
        Write-Host ""
        Write-Host "  ▸ SPEED NOTE: The search pauses 100ms between each folder to avoid Microsoft" -ForegroundColor DarkGray
        Write-Host " rate-limiting. Very large drives may take several minutes. A progress bar" -ForegroundColor DarkGray
        Write-Host " keeps you updated and [P]/[C] work throughout." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ AFTER THE SEARCH: You can open items in the browser, export to CSV, or" -ForegroundColor DarkGray
        Write-Host " type a new search term without re-entering the URL." -ForegroundColor DarkGray
        Write-Host ""

        $global:ExitSearch = $false
        $opDidRun = $false   # v111: track for press-any-key
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SharePoint folder URL to search within" -Type "Cloud" -LogLabel "Cloud Search URL"
            if ($inputSrc -eq "Q") { break }

            $srcParsed = Parse-CloudUrl -Url $inputSrc.Trim()
            if (-not $srcParsed) {
                Write-Host " That URL could not be understood. Make sure you copy the full address from the browser." -ForegroundColor Red
                continue
            }

            while ($true) {
                $searchTerm = Read-Host "Search term (e.g. *.pdf, *report*, Annual Budget.xlsx) | [S]=change folder | [Q]=quit"
                Write-Log "[USER INPUT] Cloud Search Term: '$searchTerm'" "INFO" -FileOnly

                if ($searchTerm -match "^[Qq]$") { $global:ExitSearch = $true; break }
                if ($searchTerm -match "^[Ss]$") { break }
                if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                    Write-Host " Please enter a search term. Use * as a wildcard - e.g. *.docx or *budget*" -ForegroundColor Yellow
                    continue
                }

                Perform-CloudSearch -CloudData $srcParsed -SearchTerm $searchTerm
                $opDidRun = $true
                Write-Host ""
            }
            if ($global:ExitSearch) { $global:ExitSearch = $false; break }
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    # ============================================================================================
    # SECTION B - LOCAL / CROSS-TENANT OPERATIONS
    # ============================================================================================

    elseif ($choice -in @('7','8')) {
        $actionWord = if ($choice -eq '7') { "COPY" } else { "MOVE" }
        $actionVerb = if ($choice -eq '7') { "copy" } else { "move" }

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  LOCAL $actionWord - $($actionVerb.ToUpper()) Files Between Folders on Your PC" -ForegroundColor Green
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Performs a smart $actionVerb of every file inside a folder on your hard drive."
        Write-Host "   'Smart' means it automatically SKIPS files that already exist and are identical"
        Write-Host "   at the destination - so you can safely resume an interrupted job."
        Write-Host ""
        Write-Host "  ▸ TWO MAIN USE CASES" -ForegroundColor White
        Write-Host "   1. CROSS-SITE TRANSFERS: If you need to $actionVerb data from one"
        Write-Host "      SharePoint site or tenant to another, sync both to your PC with"
        Write-Host "      desktop app first. Then use this to $actionVerb between the two local sync folders."
        Write-Host ""
        Write-Host "   2. STANDARD LOCAL TRANSFERS: Moving or copying files between folders on a"
        Write-Host "      hard drive, USB drive, or mapped network share (e.g. Z:\)."
        Write-Host ""
        Write-Host "  ▸ HOW TO FIND YOUR LOCAL PATH" -ForegroundColor White
        Write-Host "   Open File Explorer -> navigate to the folder -> click the address bar at the top"
        Write-Host "   -> the path appears highlighted. Copy and paste it here."
        Write-Host "   Examples:" -ForegroundColor DarkGray
        Write-Host "     C:\Users\$($ActiveUser)\OneDrive - Contoso\Shared Documents\Projects" -ForegroundColor DarkGray
        Write-Host "     D:\Archives\2024" -ForegroundColor DarkGray
        Write-Host "     Z:\Department\Finance" -ForegroundColor DarkGray
        Write-Host "     \\FileServer01\Shared\Archives       (UNC path — network share)" -ForegroundColor DarkGray
        Write-Host "     \\nas.local\data\Projects          (UNC path — NAS device)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ✔ UNC PATHS FULLY SUPPORTED: You can use \\server\share\... as source or" -ForegroundColor DarkGray
        Write-Host " destination. Free-space checks work on network shares too." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "    The script checks available disk space before starting and warns you if it is low." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SOURCE folder path (the folder you want to $actionVerb FROM)" -Type "Local" -LogLabel "Source Path"
            if ($inputSrc -eq "Q") { Write-Log "Local $actionWord cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the DESTINATION folder path (the folder you want to $actionVerb INTO)" -Type "Local" -LogLabel "Destination Path"
            if ($inputDst -eq "Q") { Write-Log "Local $actionWord cancelled by user." "INFO"; break }

            $s = $inputSrc.Trim('"', "'", ' ').TrimEnd('\')
            $d = $inputDst.Trim('"', "'", ' ').TrimEnd('\')

            if (-not (Test-Path -LiteralPath $s)) {
                Write-Host " ERROR: The source folder '$s' does not exist on this computer." -ForegroundColor Red
                Write-Host " Tip: Check the path in File Explorer. If it is a network or OneDrive folder, make sure it is synced and connected." -ForegroundColor Yellow
                continue
            }
            if (-not (Test-Path -LiteralPath $d)) {
                Write-Host " The destination folder does not exist yet - creating it automatically..." -ForegroundColor DarkGray
                try { New-Item -ItemType Directory -Path $d -Force | Out-Null } catch {}
            }

            if ($choice -eq '7') { Perform-LocalTransfer -Src $s -Dst $d -IsMove:$false }
            if ($choice -eq '8') { Perform-LocalTransfer -Src $s -Dst $d -IsMove:$true  }
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another SOURCE path to continue, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '9') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  LOCAL BATCH - Process Specific Items from a Folder" -ForegroundColor Green
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Shows you the top-level contents of a local folder as a numbered list."
        Write-Host "   You pick which items to process by typing their numbers."
        Write-Host "   Then choose an action: Copy, Move, Delete, or Verify."
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS" -ForegroundColor White
        Write-Host "   Use this when a folder has many sub-folders but you only want some of them."
        Write-Host "   Example: a folder with 50 sub-folders but you only want to process 10."
        Write-Host ""
        Write-Host "  ▸ HOW TO SELECT ITEMS" -ForegroundColor White
        Write-Host "   Single item  ->  type   3"
        Write-Host "   A range      ->  type   5-10   (processes items 5, 6, 7, 8, 9, 10)"
        Write-Host "   Mixed        ->  type   1, 3, 7-9   (processes 1, 3, 7, 8, 9)"
        Write-Host "   Everything   ->  type   A"
        Write-Host "   Backwards OK ->  type   10-5   (same as 5-10, automatically sorted)"
        Write-Host ""
        Write-Host "  ▸ DESTINATION PATH: Only needed for Copy, Move, or Verify. Leave blank for Delete." -ForegroundColor DarkGray
        Write-Host "  ✔ UNC PATHS are supported for both source and destination (\\server\share\folder)." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SOURCE folder path (the folder whose contents you want to manage)" -Type "Local" -LogLabel "Source Path"
            if ($inputSrc -eq "Q") { Write-Log "Local Batch cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the DESTINATION folder path (press ENTER to skip - only needed for Copy, Move, Verify)" -Type "Local" -LogLabel "Destination Path" -AllowEmpty
            if ($inputDst -eq "Q") { Write-Log "Local Batch cancelled by user." "INFO"; break }

            $s = $inputSrc.Trim('"', "'", ' ').TrimEnd('\')
            $d = if ([string]::IsNullOrWhiteSpace($inputDst)) { $null } else { $inputDst.Trim('"', "'", ' ').TrimEnd('\') }

            if (-not (Test-Path -LiteralPath $s)) {
                Write-Host " ERROR: The source folder '$s' does not exist on this computer." -ForegroundColor Red
                Write-Host " Tip: Check the path in File Explorer and try again." -ForegroundColor Yellow
                continue
            }
            if ($d -and -not (Test-Path -LiteralPath $d)) {
                Write-Host " The destination folder does not exist yet - creating it automatically..." -ForegroundColor DarkGray
                try { New-Item -ItemType Directory -Path $d -Force | Out-Null } catch {}
            }

            Perform-LocalBatch -Src $s -Dst $d
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another SOURCE path, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '10') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  LOCAL WIPE - Permanently Delete a Local Folder" -ForegroundColor Green
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Completely and permanently deletes a folder and everything inside it"
        Write-Host "   from your hard drive."
        Write-Host ""
        Write-Host "  !!! CRITICAL WARNING !!!" -ForegroundColor Red
        Write-Host "   THIS BYPASSES THE WINDOWS RECYCLE BIN." -ForegroundColor Red
        Write-Host "   Files deleted this way CANNOT be recovered from the Recycle Bin." -ForegroundColor Red
        Write-Host "   Only use this AFTER you have verified the data was successfully transferred." -ForegroundColor Red
        Write-Host ""
        Write-Host "  ▸ RECOMMENDED WORKFLOW" -ForegroundColor White
        Write-Host "   Step 1: Copy the data using [7] or [8]."
        Write-Host "   Step 2: Verify the copy is perfect using [11]."
        Write-Host "   Step 3: ONLY THEN use this option to delete the source."
        Write-Host ""
        Write-Host "  ! SAFETY GATE: Type 'YES' in capitals to confirm, anything else cancels." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the path of the LOCAL FOLDER you want to permanently delete" -Type "Local" -LogLabel "Wipe Path"
            if ($inputSrc -eq "Q") { Write-Log "Local Wipe cancelled by user." "INFO"; break }

            $s = $inputSrc.Trim('"', "'", ' ').TrimEnd('\')
            if (-not (Test-Path -LiteralPath $s)) {
                Write-Host " ERROR: The folder '$s' does not exist on this computer." -ForegroundColor Red
                Write-Host " Tip: Double-check the path by navigating to it in File Explorer first." -ForegroundColor Yellow
                continue
            }
            Perform-LocalWipe -TargetFolder $s
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another path to delete, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '11') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  LOCAL VERIFY - Prove Two Local Folders Are Identical" -ForegroundColor Green
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Compares every file inside two local folders and produces a detailed report."
        Write-Host "   For each file the script checks:"
        Write-Host "     - Is the file present in both folders?          (missing = something was not copied)"
        Write-Host "     - Are the file sizes in bytes EXACTLY identical? (mismatch = possible corruption)"
        Write-Host "     - Are the last-modified dates identical?         (date difference = file may have changed)"
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS" -ForegroundColor White
        Write-Host "   Always run this after a Local Copy or Move to guarantee 100% accuracy."
        Write-Host "   It gives you mathematical proof that nothing was missed or corrupted"
        Write-Host "   before you safely delete the source."
        Write-Host ""
        Write-Host "  ▸ RESULT GUIDE" -ForegroundColor White
        Write-Host "   Perfectly Synced  = identical in both folders. Safe to delete source." -ForegroundColor DarkGray
        Write-Host "   Mismatch          = same name but size or date differs. Re-copy recommended." -ForegroundColor DarkGray
        Write-Host "   Missing in Dest   = file is in source but absent from destination. Copy failed." -ForegroundColor DarkGray
        Write-Host "   Orphaned in Dest  = file in destination but NOT in source. Unexpected extra file." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SOURCE folder path (the original / master copy)" -Type "Local" -LogLabel "Source Path"
            if ($inputSrc -eq "Q") { Write-Log "Local Verify cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the DESTINATION folder path (the copy you want to check)" -Type "Local" -LogLabel "Destination Path"
            if ($inputDst -eq "Q") { Write-Log "Local Verify cancelled by user." "INFO"; break }

            $s = $inputSrc.Trim('"', "'", ' ').TrimEnd('\')
            $d = $inputDst.Trim('"', "'", ' ').TrimEnd('\')

            if (-not (Test-Path -LiteralPath $s)) {
                Write-Host " ERROR: The source folder '$s' does not exist. Check the path in File Explorer." -ForegroundColor Red; continue
            }
            if (-not (Test-Path -LiteralPath $d)) {
                Write-Host " ERROR: The destination folder '$d' does not exist. Check the path in File Explorer." -ForegroundColor Red; continue
            }
            Perform-LocalVerify -Src $s -Dst $d
            $opDidRun = $true

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another SOURCE path to verify, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    elseif ($choice -eq '12') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  LOCAL SEARCH - Find Files on Your PC" -ForegroundColor Green
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Scans every sub-folder inside a directory on your PC and returns a list"
        Write-Host "   of every file and folder whose name matches your search term."
        Write-Host "   The search is crash-proof - it silently skips locked or restricted folders"
        Write-Host "   instead of stopping."
        Write-Host ""
        Write-Host "  ▸ HOW TO SEARCH" -ForegroundColor White
        Write-Host "   Exact name:    Type the full file name,   e.g.   Payslip March 2024.pdf"
        Write-Host "   By extension:  Type a wildcard,           e.g.   *.pdf   or   *.xlsx"
        Write-Host "   Partial name:  Use a wildcard,            e.g.   *invoice*   or   *John Smith*"
        Write-Host ""
        Write-Host "  ▸ AFTER THE SEARCH: Type a result number to see its full path, open the file," -ForegroundColor DarkGray
        Write-Host " or open its parent folder in File Explorer. Export to CSV is also available." -ForegroundColor DarkGray
        Write-Host ""

        $global:ExitSearch = $false
        $opDidRun = $false   # v111: track for press-any-key
        while ($true) {
            $inputPath = Read-PathInput -Prompt "Paste the path of the folder you want to search within (e.g. C:\Users\$($ActiveUser)\OneDrive)" -Type "Local" -LogLabel "Search Path"
            if ($inputPath -eq "Q") { break }

            $inputPath = $inputPath.Trim('"', "'", ' ').TrimEnd('\')
            if (-not (Test-Path -LiteralPath $inputPath)) {
                Write-Host " ERROR: The folder '$inputPath' does not exist on this computer." -ForegroundColor Red
                Write-Host " Tip: Check the path in File Explorer and try again." -ForegroundColor Yellow
                continue
            }

            while ($true) {
                $searchTerm = Read-Host "Search term (e.g. *.pdf, *report*, Payslip.docx) | [S]=change folder | [Q]=quit"
                Write-Log "[USER INPUT] Local Search Term: '$searchTerm'" "INFO" -FileOnly

                if ($searchTerm -match "^[Qq]$") { $global:ExitSearch = $true; break }
                if ($searchTerm -match "^[Ss]$") { break }
                if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                    Write-Host " Please enter a search term. Use * as a wildcard - e.g. *.xlsx or *smith*" -ForegroundColor Yellow
                    continue
                }

                Perform-LocalSearch -SearchPath $inputPath -SearchTerm $searchTerm
                $opDidRun = $true
                Write-Host ""
            }
            if ($global:ExitSearch) { $global:ExitSearch = $false; break }
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    # ============================================================================================
    # SECTION C - DIRECT DOWNLOAD OPERATIONS
    # ============================================================================================

    elseif ($choice -in @('13','14')) {
        $isBatch  = ($choice -eq '14')
        $modeName = if ($isBatch) { "BATCH DOWNLOAD" } else { "DIRECT DOWNLOAD" }

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Magenta
        Write-Host "  $modeName - Pull SharePoint Files to Your PC" -ForegroundColor Magenta
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Magenta
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        if ($isBatch) {
            Write-Host "   Shows the top-level contents of a SharePoint folder as a numbered list."
            Write-Host "   You pick which items to download. The script then pulls them directly"
            Write-Host "   to your hard drive - without using the OneDrive Sync app."
        } else {
            Write-Host "   Downloads an entire SharePoint folder (all sub-folders and files)"
            Write-Host "   directly to your hard drive - without using the OneDrive Sync app."
        }
        Write-Host ""
        Write-Host "  ▸ WHEN TO USE THIS vs. OneDrive Sync" -ForegroundColor White
        Write-Host "   Use this script ->  when you need a one-time archive on a specific folder"
        Write-Host "                       or when OneDrive Sync is unavailable / too slow to configure."
        Write-Host "   Use OneDrive Sync -> when you need ongoing, automatic two-way sync."
        Write-Host ""
        Write-Host "  ✔ SMART RESUME: If interrupted, re-run with the same destination folder." -ForegroundColor DarkGray
        Write-Host " Files already downloaded will be skipped automatically." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ WHAT IS THE TEMP BUFFER?" -ForegroundColor White
        Write-Host "   Each file is downloaded to a holding area first, then moved to the final"
        Write-Host "   destination. This prevents corrupted partial downloads appearing in your folder."
        Write-Host "   Recommendation: Use a simple path like C:\Temp or D:\Temp."
        Write-Host "   (Create the folder in File Explorer first if it does not exist yet.)"
        Write-Host ""
        Write-Host "  ▸ HOW TO GET THE SHAREPOINT URL" -ForegroundColor White
        Write-Host "   Open the SharePoint folder in your browser -> copy the full URL from the"
        Write-Host "   address bar -> paste it below. Tracking codes are stripped automatically."
        Write-Host ""
        Write-Host "  ✔ UNC DESTINATION SUPPORTED: Your local destination path can be a UNC path." -ForegroundColor DarkGray
        Write-Host "   Example: \\\\FileServer01\\Shared\\Archives   or   \\\\nas\\data\\archive" -ForegroundColor DarkGray
        Write-Host "   Free-space checks work on UNC shares just like local drives." -ForegroundColor DarkGray
        Write-Host ""

        # v111: stay-on-operation loop. Previously this branch ran the engine
        # once and fell back to the main menu — the operator lost the chance
        # to chain multiple downloads without re-navigating the menu, and the
        # post-op summary scrolled away the moment the main menu redrew.
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SharePoint SOURCE folder URL (the cloud folder to download FROM)" -Type "Cloud" -LogLabel "Source URL"
            if ($inputSrc -eq "Q") { break }

            $inputDst = Read-PathInput -Prompt "Paste the LOCAL DESTINATION path (the folder on your PC to save files INTO, e.g. D:\Archives\2024)" -Type "Local" -LogLabel "Destination Path"
            if ($inputDst -eq "Q") { break }

            Write-Host ""
            Write-Host " TEMP BUFFER - Where should files be staged before being moved to the destination?" -ForegroundColor Yellow
            Write-Host " Tip: Press ENTER to automatically use your system Temp folder ($env:TEMP)." -ForegroundColor DarkGray
            Write-Host " Or paste any local/UNC path, e.g. C:\Temp or \\server\share\TempBuf" -ForegroundColor DarkGray
            $inputBufferPath = Read-PathInput -Prompt "Paste the TEMP BUFFER path (or press ENTER to use $env:TEMP)" -Type "Local" -LogLabel "Buffer Path" -AllowEmpty
            if ($inputBufferPath -eq "Q") { break }

            # Auto-default: empty answer uses the system temp folder
            if ([string]::IsNullOrWhiteSpace($inputBufferPath)) {
                $inputBufferPath = $env:TEMP
                Write-Host "  Using system Temp folder: $inputBufferPath" -ForegroundColor DarkGray
            }

            $cleanLocal  = $inputDst.Trim('"', "'", ' ').TrimEnd('\')
            $cleanBuffer = $inputBufferPath.Trim('"', "'", ' ').TrimEnd('\')
            $srcParsed   = Parse-CloudUrl -Url $inputSrc.Trim()

            if (-not $srcParsed) {
                Write-Host " ERROR: The SharePoint URL could not be understood. Copy the full URL from the browser address bar." -ForegroundColor Red
                # Stay in the loop — let the operator try again with a corrected URL
                continue
            }

            $isValidPath = $true
            try {
                if (-not (Test-Path -LiteralPath $cleanBuffer)) { New-Item -ItemType Directory -Path $cleanBuffer -Force -ErrorAction Stop | Out-Null }
                $testFile = Join-Path $cleanBuffer "write_test.tmp"
                "test" | Out-File -FilePath $testFile -ErrorAction Stop
                Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
            } catch {
                $isValidPath = $false
                Write-Host " ERROR: Cannot write to the temp buffer path '$cleanBuffer'." -ForegroundColor Red
                Write-Host " Make sure the folder exists and you have permission to write to it." -ForegroundColor Yellow
                Write-Host " Try creating C:\Temp in File Explorer and using that." -ForegroundColor Yellow
                Write-Log "Cannot write to the Buffer path '$cleanBuffer'. Check folder permissions." "ERROR"
            }

            if ($isValidPath) {
                try {
                    Write-Host ""
                    Write-Host " Connecting to SharePoint... A browser login window may appear." -ForegroundColor DarkGray
                    Write-Host " Sign in with your Microsoft account, then return to this window." -ForegroundColor DarkGray
                    $originalWarning   = $WarningPreference
                    $WarningPreference = "SilentlyContinue"
                    Connect-PnPOnline -Url $srcParsed.SiteUrl -UseWebLogin -ErrorAction Stop
                    $WarningPreference = $originalWarning

                    if ($choice -eq '13') { Perform-DirectDownload            -CloudData $srcParsed -LocalDest $cleanLocal -BufferDest $cleanBuffer }
                    if ($choice -eq '14') { Perform-CloudToLocalBatchDownload -CloudData $srcParsed -LocalDest $cleanLocal -BufferDest $cleanBuffer }
                    $opDidRun = $true
                } catch {
                    Write-Host " ERROR: Connection failed. $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host " Make sure you are connected to the internet and signed into the correct Microsoft account." -ForegroundColor Yellow
                    Write-Log "Connection Error: $($_.Exception.Message)" "ERROR"
                } finally {
                    try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
                }
            }

            # v111: brief end-of-op separator + prompt continuation hint
            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Operation complete. Enter another SharePoint URL to continue," -ForegroundColor DarkGray
            Write-Host "  or press [Q] at the next prompt to return to the main menu." -ForegroundColor DarkGray
            Write-Host ""
        }

        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    # ============================================================================================
    # SECTION D - UPLOAD OPERATIONS (new in v89)
    # ============================================================================================

    elseif ($choice -in @('15','16')) {
        $isBatch  = ($choice -eq '16')
        $modeName = if ($isBatch) { "BATCH UPLOAD" } else { "DIRECT UPLOAD" }

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "  $modeName - Push Local Files to SharePoint" -ForegroundColor Cyan
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        if ($isBatch) {
            Write-Host "   Shows the top-level contents of a LOCAL folder as a numbered list."
            Write-Host "   You pick which items to upload. The script pushes them directly to"
            Write-Host "   SharePoint - creating all sub-folder structures automatically."
        } else {
            Write-Host "   Uploads an entire local folder (all sub-folders and files) directly"
            Write-Host "   to a SharePoint destination - creating all sub-folders automatically."
        }
        Write-Host ""
        Write-Host "  ▸ TRANSFERRING BETWEEN TWO SHAREPOINT SITES (Cross-Site Bridge)" -ForegroundColor White
        Write-Host "   Direct cloud-to-cloud transfers between different sites are blocked by"
        Write-Host "   some organisations' security policies. The workaround is a two-step"
        Write-Host "   bridge that uses your PC as a temporary staging area:"
        Write-Host "    Step 1: Download from Site A using [13] Direct Download or [14] Batch Download."
        Write-Host "    Step 2: Upload that local folder to Site B using [15] or [16] here."
        Write-Host "   Both steps use standard browser login \u2014 no extra credentials needed."
        Write-Host "   security policy permits - no extra credentials or workarounds required."
        Write-Host ""
        Write-Host "  ▸ SMART RESUME" -ForegroundColor White
        Write-Host "   Files already on SharePoint with the SAME byte size are skipped."
        Write-Host "   If the upload is interrupted, re-run with the same paths to resume."
        Write-Host ""
        Write-Host "  ▸ AUTOMATIC FOLDER CREATION" -ForegroundColor White
        Write-Host "   Sub-folder structures are created automatically in SharePoint as needed."
        Write-Host "   You do NOT need to pre-create any folders at the destination."
        Write-Host ""
        Write-Host "  ▸ PATH LENGTH LIMIT" -ForegroundColor White
        Write-Host "   SharePoint enforces a $CloudPathLimit-character path limit. Files exceeding this"
        Write-Host "   are skipped and logged to LongPaths_FailedToSync.txt in your log folder."
        Write-Host ""
        Write-Host "  ▸ HOW TO GET THE DESTINATION URL" -ForegroundColor White
        Write-Host "   Open the target SharePoint folder in your browser -> copy the full URL"
        Write-Host "   from the address bar -> paste it below. Tracking codes stripped automatically."
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the LOCAL SOURCE path (the folder on your PC to upload FROM)" -Type "Local" -LogLabel "Upload Source Path"
            if ($inputSrc -eq "Q") { Write-Log "$modeName cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the SharePoint DESTINATION URL (the cloud folder to upload INTO)" -Type "Cloud" -LogLabel "Upload Destination URL"
            if ($inputDst -eq "Q") { Write-Log "$modeName cancelled by user." "INFO"; break }

            $s = $inputSrc.Trim('"', "'", ' ').TrimEnd('\')

            if (-not (Test-Path -LiteralPath $s)) {
                Write-Host " ERROR: The source folder '$s' does not exist on this computer." -ForegroundColor Red
                Write-Host " Tip: Check the path in File Explorer and try again." -ForegroundColor Yellow
                continue
            }

            $dstParsed = Parse-CloudUrl -Url $inputDst.Trim()
            if (-not $dstParsed) {
                Write-Host " ERROR: The SharePoint destination URL could not be understood." -ForegroundColor Red
                Write-Host " Make sure you copy the full URL from the browser address bar." -ForegroundColor Yellow
                continue
            }

            try {
                Write-Host ""
                Write-Host " Connecting to SharePoint... A browser login window may appear." -ForegroundColor DarkGray
                Write-Host " Sign in with your Microsoft account, then return to this window." -ForegroundColor DarkGray
                $originalWarning   = $WarningPreference
                $WarningPreference = "SilentlyContinue"
                Connect-PnPOnline -Url $dstParsed.SiteUrl -UseWebLogin -ErrorAction Stop
                $WarningPreference = $originalWarning

                if ($choice -eq '15') { Perform-DirectUpload            -CloudDst $dstParsed -LocalSrc $s }
                if ($choice -eq '16') { Perform-LocalToCloudBatchUpload -LocalSrc $s -CloudDst $dstParsed }
                $opDidRun = $true
            } catch {
                Write-Host " ERROR: Connection failed. $($_.Exception.Message)" -ForegroundColor Red
                Write-Host " Make sure you are connected to the internet and signed into the correct Microsoft account." -ForegroundColor Yellow
                Write-Log "Connection Error: $($_.Exception.Message)" "ERROR"
            } finally {
                try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
            }

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another LOCAL SOURCE path to continue, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }

    # ============================================================================================
    # SECTION E - ARCHIVE OPERATIONS
    # ============================================================================================

    elseif ($choice -in @('17','18')) {
        $isBatch  = ($choice -eq '18')
        $modeName = if ($isBatch) { "BATCH ARCHIVE" } else { "ARCHIVE DOWNLOAD" }

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  $modeName - Download, Verify, then Clean Up Cloud" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        if ($isBatch) {
            Write-Host "   Shows the top-level contents of a cloud folder as a numbered list."
            Write-Host "   For each item you select, the script runs three steps in sequence:"
        } else {
            Write-Host "   Downloads a SharePoint folder to your PC, verifies it, then"
            Write-Host "   automatically deletes the cloud original if verification passes:"
        }
        Write-Host "     Step 1: DOWNLOAD — pulls all files to your hard drive."
        Write-Host "     Step 2: VERIFY   — confirms every file arrived with the correct byte size."
        Write-Host "     Step 3: DELETE   — if verification passes, deletes from cloud automatically."
        Write-Host ""
        if ($isBatch) {
            Write-Host "  ▸ FULLY AUTOMATIC — NO PROMPTS" -ForegroundColor White
            Write-Host "   The script runs completely unattended after you start it."
            Write-Host "   Each item is downloaded, verified, and if the counts and bytes"
            Write-Host "   match, deleted — with no prompts at any point."
            Write-Host "   Press [C] at any time to stop cleanly after the current item."
            Write-Host ""
        }
        Write-Host "  ▸ IDEAL USE CASE: CLOUD STORAGE ARCHIVING" -ForegroundColor White
        Write-Host "   Download a cloud folder to your hard drive, verify every file arrived"
        Write-Host "   intact, then automatically delete the cloud original to free up storage."
        Write-Host "   The verify step gives you mathematical proof the data is safe before"
        Write-Host "   anything is removed from SharePoint."
        Write-Host ""
        Write-Host "  ▸ WHAT GETS DELETED" -ForegroundColor White
        Write-Host "   Deletion sends the folder to the SharePoint RECYCLE BIN — NOT permanently."
        Write-Host "   Items remain recoverable for approximately 93 days."
        Write-Host "   Cloud storage is freed once the item leaves the Recycle Bin."
        Write-Host ""
        Write-Host "  ▸ SAFETY GATES — the script will REFUSE to delete if:" -ForegroundColor White
        Write-Host "   • Any file failed to download (even one failure blocks deletion)."
        Write-Host "   • The local file count or total byte size doesn't match the cloud source."
        Write-Host "   • The job was cancelled mid-download by pressing [C]."
        Write-Host ""
        Write-Host "  ▸ SMART RESUME" -ForegroundColor White
        Write-Host "   Files already present locally with the same byte size are counted as"
        Write-Host "   'Skipped' — they are included in the verification totals, so re-running"
        Write-Host "   an interrupted archive will pick up exactly where it left off."
        Write-Host ""
        Write-Host "  ▸ HOW TO GET THE SHAREPOINT URL" -ForegroundColor White
        Write-Host "   Open the cloud folder in your browser -> copy the full URL from the"
        Write-Host "   address bar -> paste below. Tracking codes are stripped automatically."
        Write-Host ""

        # v111: stay-on-operation loop
        $opDidRun = $false
        while ($true) {
            $inputSrc = Read-PathInput -Prompt "Paste the SharePoint SOURCE folder URL (the cloud folder to archive FROM)" -Type "Cloud" -LogLabel "Archive Source URL"
            if ($inputSrc -eq "Q") { Write-Log "$modeName cancelled by user." "INFO"; break }

            $inputDst = Read-PathInput -Prompt "Paste the LOCAL DESTINATION path (where to save the archive on your PC / server)" -Type "Local" -LogLabel "Archive Destination Path"
            if ($inputDst -eq "Q") { Write-Log "$modeName cancelled by user." "INFO"; break }

            Write-Host ""
            Write-Host " TEMP BUFFER - Where should files be staged during download?" -ForegroundColor Yellow
            Write-Host " Tip: Press ENTER to automatically use your system Temp folder ($env:TEMP)." -ForegroundColor DarkGray
            $inputBufferPath = Read-PathInput -Prompt "Paste the TEMP BUFFER path (or press ENTER to use $env:TEMP)" -Type "Local" -LogLabel "Buffer Path" -AllowEmpty
            if ($inputBufferPath -eq "Q") { break }

            if ([string]::IsNullOrWhiteSpace($inputBufferPath)) {
                $inputBufferPath = $env:TEMP
                Write-Host "  Using system Temp folder: $inputBufferPath" -ForegroundColor DarkGray
            }

            $cleanLocal  = $inputDst.Trim('"', "'", ' ').TrimEnd('\')
            $cleanBuffer = $inputBufferPath.Trim('"', "'", ' ').TrimEnd('\')
            $srcParsed   = Parse-CloudUrl -Url $inputSrc.Trim()

            if (-not $srcParsed) {
                Write-Host " ERROR: The SharePoint URL could not be understood. Copy the full URL from the browser." -ForegroundColor Red
                continue
            }

            $isValidPath = $true
            try {
                if (-not (Test-Path -LiteralPath $cleanBuffer)) { New-Item -ItemType Directory -Path $cleanBuffer -Force -ErrorAction Stop | Out-Null }
                $testFile = Join-Path $cleanBuffer "write_test.tmp"
                "test" | Out-File -FilePath $testFile -ErrorAction Stop
                Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
            } catch {
                $isValidPath = $false
                Write-Host " ERROR: Cannot write to the temp buffer '$cleanBuffer'. Check folder permissions." -ForegroundColor Red
                Write-Log "Cannot write to Buffer path '$cleanBuffer'." "ERROR"
            }

            if ($isValidPath) {
                try {
                    Write-Host ""
                    Write-Host " Connecting to SharePoint... A browser login window may appear." -ForegroundColor DarkGray
                    Write-Host " Sign in with your Microsoft account, then return to this window." -ForegroundColor DarkGray
                    $originalWarning   = $WarningPreference
                    $WarningPreference = "SilentlyContinue"
                    Connect-PnPOnline -Url $srcParsed.SiteUrl -UseWebLogin -ErrorAction Stop
                    $WarningPreference = $originalWarning

                    # v111: removed the inline "Press ENTER to return to main menu" — replaced by the
                    # unified Show-PressAnyKeyToContinue at the end of the operation loop below.
                    if ($choice -eq '17') { Perform-ArchiveDownload      -CloudData $srcParsed -LocalDest $cleanLocal -BufferDest $cleanBuffer }
                    if ($choice -eq '18') { Perform-BatchArchiveDownload -CloudData $srcParsed -LocalDest $cleanLocal -BufferDest $cleanBuffer }
                    $opDidRun = $true
                } catch {
                    Write-Host " ERROR: Connection failed. $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host " Make sure you are connected to the internet and signed into the correct Microsoft account." -ForegroundColor Yellow
                    Write-Log "Connection Error: $($_.Exception.Message)" "ERROR"
                } finally {
                    try { $null = Disconnect-PnPOnline -ErrorAction Stop } catch {}
                }
            }

            Write-Host ""
            Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "  Enter another SOURCE URL to archive, or [Q] at the next prompt for main menu." -ForegroundColor DarkGray
            Write-Host ""
        }
        if ($opDidRun) { Show-PressAnyKeyToContinue }
    }


    elseif ($choice -eq '19') {
        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  MULTI-SOURCE ARCHIVE - Archive Multiple Cloud Sources at Once" -ForegroundColor Yellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   You paste the SharePoint URLs of multiple cloud folders one at a time."
        Write-Host "   combination of SharePoint folders) one at a time."
        Write-Host "   For each URL you can set a friendly label — this becomes the subfolder"
        Write-Host "   name under your destination, so archives are clearly organised:"
        Write-Host ""
        Write-Host "   D:\\Archives\\2024\\" -ForegroundColor DarkGray
        Write-Host "     Project Alpha\\    (downloaded from source 1)" -ForegroundColor DarkGray
        Write-Host "     Project Beta\\     (downloaded from source 2)" -ForegroundColor DarkGray
        Write-Host "     User Data\\        (downloaded from source 3)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ FULLY AUTOMATIC AFTER SETUP" -ForegroundColor White
        Write-Host "   Once you have entered all URLs the script runs completely unattended:"
        Write-Host "   Download → Verify → Auto-delete on pass, for every source in sequence."
        Write-Host "   A per-source outcome table is shown at the end."
        Write-Host ""
        Write-Host "  ▸ HANDLES DIFFERENT SITES AUTOMATICALLY" -ForegroundColor White
        Write-Host "   Each source URL may belong to a different SharePoint site collection."
        Write-Host "   The script reconnects automatically when the site changes \u2014"
        Write-Host "   you may see a browser login prompt for each different site."
        Write-Host ""
        Write-Host "  ▸ SAFETY GATES (same as Archive Download)" -ForegroundColor White
        Write-Host "   Any download failure OR count/bytes mismatch blocks deletion for that source."
        Write-Host "   Other sources in the queue continue unaffected."
        Write-Host ""
        Write-Host "  ▸ HOW TO GET ONEDRIVE URLS" -ForegroundColor White
        Write-Host "   Open the SharePoint or OneDrive folder in your browser \u2192 navigate"
        Write-Host "   access) → navigate to the folder you want → copy the address bar URL."
        Write-Host "   Tracking codes are stripped automatically."
        Write-Host ""

        # ── COLLECT DESTINATION ───────────────────────────────────────────────
        $inputDst = Read-PathInput -Prompt "Paste the LOCAL DESTINATION root (a single folder to put all archives into)" -Type "Local" -LogLabel "Multi-Archive Destination"
        if ($inputDst -eq "Q") { Write-Log "Multi-Source Archive cancelled by user." "INFO"; continue }

        Write-Host ""
        Write-Host " TEMP BUFFER - Press ENTER to use your system Temp folder ($env:TEMP)." -ForegroundColor DarkGray
        $inputBuf = Read-PathInput -Prompt "Paste the TEMP BUFFER path (or press ENTER for $env:TEMP)" -Type "Local" -LogLabel "Multi-Archive Buffer" -AllowEmpty
        if ($inputBuf -eq "Q") { continue }
        if ([string]::IsNullOrWhiteSpace($inputBuf)) { $inputBuf = $env:TEMP; Write-Host "  Using: $inputBuf" -ForegroundColor DarkGray }

        $cleanDst = $inputDst.Trim('"', "'", ' ').TrimEnd('\')
        $cleanBuf = $inputBuf.Trim('"', "'", ' ').TrimEnd('\')

        # Validate buffer is writable
        $bufValid = $true
        try {
            if (-not (Test-Path -LiteralPath $cleanBuf)) { New-Item -ItemType Directory -Path $cleanBuf -Force -ErrorAction Stop | Out-Null }
            $tf = Join-Path $cleanBuf "write_test.tmp"; "test" | Out-File -FilePath $tf -ErrorAction Stop
            Remove-Item $tf -Force -ErrorAction SilentlyContinue
        } catch {
            $bufValid = $false
            Write-Host " ERROR: Cannot write to temp buffer '$cleanBuf'. Check permissions." -ForegroundColor Red
            Write-Log "Multi-Archive: Buffer path unwritable '$cleanBuf'." "ERROR"
        }
        if (-not $bufValid) { continue }

        # ── COLLECT SOURCE URLS ───────────────────────────────────────────────
        $SourceList = [System.Collections.Generic.List[hashtable]]::new()
        Write-Host ""
        Write-Host "  ┌─ ENTER SOURCE URLs ─────────────────────────────────────" -ForegroundColor Cyan
        Write-Host "  │  Paste each SharePoint URL and press ENTER." -ForegroundColor DarkGray
        Write-Host "  │  Press ENTER on a blank line when done." -ForegroundColor DarkGray
        Write-Host "  │  Type [Q] at the URL prompt to cancel." -ForegroundColor DarkGray
        Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Cyan

        $srcNum = 0
        while ($true) {
            $srcNum++
            Write-Host ""
            $rawUrl = Read-Host "  Source $srcNum URL (or ENTER to finish)"
            Write-Log "[USER INPUT] Multi-Archive Source $srcNum URL: '$rawUrl'" "INFO" -FileOnly

            if ($rawUrl -match "^[Qq]$") {
                Write-Host "  Cancelled." -ForegroundColor Yellow
                Write-Log "Multi-Source Archive cancelled by user during URL entry." "INFO"
                $SourceList.Clear(); break
            }

            if ([string]::IsNullOrWhiteSpace($rawUrl)) {
                Write-Host "  URL entry complete — $($SourceList.Count) source(s) queued." -ForegroundColor DarkGray
                break
            }

            $parsed = Parse-CloudUrl -Url $rawUrl.Trim()
            if (-not $parsed) {
                Write-Host "  [!] Could not parse that URL. Check it starts with https:// and try again." -ForegroundColor Red
                $srcNum--; continue
            }

            # ── DERIVE DEFAULT SUBFOLDER NAME ─────────────────────────────────
            # Use the last non-trivial segment of FolderPath.
            # For personal sites strip common generic library names (Documents, Shared Documents)
            # and use the display name segment from the path instead.
            $pathSegments = ($parsed.FolderPath -split '/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $defaultName  = if ($pathSegments.Count -gt 0) { $pathSegments[-1] } else { "Source$srcNum" }

            # If path ends in a generic library name, also offer the personal site username
            $genericLibs = @('Documents','Shared Documents','Shared%20Documents','Documents%20','Forms','AllItems.aspx')
            if ($defaultName -in $genericLibs -or $defaultName -like "*.aspx") {
                # Try to extract the personal site username (e.g. john_smith_company_com from /personal/john_smith_company_com/...)
                $personalMatch = [regex]::Match($parsed.FolderPath, '(?i)/personal/([^/]+)')
                if ($personalMatch.Success) {
                    # Convert john_smith_company_com → john_smith (strip domain suffix)
                    $rawName = $personalMatch.Groups[1].Value
                    # Remove last two segments (domain + tld) if they look like a domain suffix
                    $parts = $rawName -split '_'
                    if ($parts.Count -ge 3) {
                        $defaultName = ($parts[0..($parts.Count - 3)] -join '_')
                    } else { $defaultName = $rawName }
                }
            }

            $safeDefault = Get-SafeWindowsName $defaultName

            Write-Host "  Parsed OK.  Default subfolder name: '$safeDefault'" -ForegroundColor Green
            $labelInput = Read-Host "  Friendly label (press ENTER to use '$safeDefault', or type a name like 'John Smith')"
            Write-Log "[USER INPUT] Multi-Archive Source $srcNum Label: '$labelInput'" "INFO" -FileOnly

            $finalLabel = if ([string]::IsNullOrWhiteSpace($labelInput)) { $safeDefault } else { Get-SafeWindowsName $labelInput.Trim() }

            $SourceList.Add(@{
                CloudData = $parsed
                SubFolder = $finalLabel
                Label     = $finalLabel
            })
            Write-Host "  [✔] Added: '$finalLabel' ← $($parsed.FolderPath)" -ForegroundColor Cyan
        }

        if ($SourceList.Count -eq 0) {
            if (-not ($rawUrl -match "^[Qq]$")) { Write-Host "  No sources entered. Returning to main menu." -ForegroundColor Yellow }
            continue
        }

        # ── SHOW PLAN PREVIEW ─────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  ┌─ ARCHIVE PLAN ─────────────────────────────────────────" -ForegroundColor Yellow
        Write-Host "  │  Destination root: $cleanDst" -ForegroundColor White
        Write-Host "  │" -ForegroundColor Yellow
        for ($pi = 0; $pi -lt $SourceList.Count; $pi++) {
            $e = $SourceList[$pi]
            $dispPath = if ($e.CloudData.FolderPath.Length -gt 50) { "..." + $e.CloudData.FolderPath.Substring($e.CloudData.FolderPath.Length - 47) } else { $e.CloudData.FolderPath }
            Write-Host "  │  [$($pi+1)] $($e.Label.PadRight(25)) ← $dispPath" -ForegroundColor Cyan
        }
        Write-Host "  │" -ForegroundColor Yellow
        Write-Host "  │  $($SourceList.Count) source(s) will be downloaded, verified, and if" -ForegroundColor DarkGray
        Write-Host "  │  verification passes, deleted from SharePoint (Recycle Bin)." -ForegroundColor DarkGray
        Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Yellow
        Write-Host ""

        # Create destination root if needed
        if (-not (Test-Path -LiteralPath $cleanDst)) {
            try { New-Item -ItemType Directory -Path $cleanDst -Force | Out-Null; Write-Host "  Created destination folder: $cleanDst" -ForegroundColor DarkGray }
            catch { Write-Host " ERROR: Could not create '$cleanDst'." -ForegroundColor Red; continue }
        }

        Write-Log "Multi-Source Archive: Starting $($SourceList.Count) source(s) to '$cleanDst'." "INFO"
        Perform-MultiSourceArchive -Sources $SourceList -LocalDest $cleanDst -BufferDest $cleanBuf
        # v111: use unified press-any-key helper instead of ad-hoc Read-Host
        Show-PressAnyKeyToContinue
    }


    elseif ($choice -eq '20') {

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor DarkYellow
        Write-Host "  LEAVER ONEDRIVE ARCHIVE - Archive a Single Leaver's OneDrive" -ForegroundColor DarkYellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Connects to a leaver's OneDrive personal site directly and archives ALL"
        Write-Host "   their Document Libraries to a local folder in one automated run:"
        Write-Host "   Connect → Enumerate Libraries → Download → Verify → Delete from cloud."
        Write-Host ""
        Write-Host "  ▸ HOW TO GET THE LEAVER'S ONEDRIVE URL" -ForegroundColor White
        Write-Host "   Open the leaver's OneDrive in your browser → copy the full URL from the"
        Write-Host "   address bar (any format is accepted — /_layouts/ and ?viewid= URLs are"
        Write-Host "   automatically stripped to the site root). Alternatively enter their" -ForegroundColor DarkGray
        Write-Host "   email address (UPN) and the script will build the URL for you." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ WHERE FILES ARE SAVED" -ForegroundColor White
        Write-Host "   Each leaver library's content lands directly in your destination folder:"
        Write-Host "     dest\Desktop\    dest\Meetings\    dest\Recordings\   ..." -ForegroundColor DarkGray
        Write-Host "   (No extra Documents\ wrapper folder is created.)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ SAFETY GATES" -ForegroundColor White
        Write-Host "   Deletion blocked if: any download failure OR file count/byte mismatch." -ForegroundColor DarkGray
        Write-Host "   Deleted items go to the SharePoint Recycle Bin (~93 days recoverable)." -ForegroundColor DarkGray
        Write-Host ""

        # ── LEAVER UPN OR URL ─────────────────────────────────────────────────
        Write-Host "  You can enter EITHER:" -ForegroundColor DarkGray
        Write-Host "   (a) The leaver's email / UPN — the script builds the URL automatically." -ForegroundColor DarkGray
        Write-Host "   (b) A full OneDrive URL — any /_layouts/ or ?viewid= format is accepted." -ForegroundColor DarkGray
        Write-Host ""

        $inputLeaver = Read-PathInput -Prompt "Leaver UPN (e.g. john.smith@company.com) OR OneDrive URL" -Type "Cloud" -LogLabel "Leaver UPN or URL"
        if ($inputLeaver -eq "Q") { Write-Log "Leaver Archive cancelled by user." "INFO"; continue }
        $inputLeaver = $inputLeaver.Trim()

        # ── RESOLVE LEAVER SITE URL ───────────────────────────────────────────
        $leaverSiteUrl = ""
        $leaverUpn     = ""

        if ($inputLeaver -match "(?i)^https://") {
            $parsedLeaver = Parse-CloudUrl -Url $inputLeaver
            if (-not $parsedLeaver) {
                Write-Host "  [!] Could not parse that URL. Paste the full https:// address." -ForegroundColor Red
                continue
            }
            $leaverSiteUrl = $parsedLeaver.SiteUrl
            Write-Host "  Site URL: $leaverSiteUrl" -ForegroundColor DarkGray
        } elseif ($inputLeaver -match "@") {
            $leaverUpn = $inputLeaver
        } else {
            Write-Host "  [!] Enter either a UPN (email address) or a full https:// URL." -ForegroundColor Red
            continue
        }

        # ── LOCAL DESTINATION & BUFFER ────────────────────────────────────────
        $inputDst = Read-PathInput -Prompt "Paste the LOCAL DESTINATION path (where to save the archive)" -Type "Local" -LogLabel "Leaver Archive Destination"
        if ($inputDst -eq "Q") { Write-Log "Leaver Archive cancelled by user." "INFO"; continue }

        Write-Host ""
        Write-Host "  ▸ WHAT IS THE TEMP BUFFER?" -ForegroundColor White
        Write-Host "   Each file is staged here first, then moved to your destination."
        Write-Host "   This prevents partial files appearing if a download is interrupted."
        Write-Host "   Recommendation: a fast local path like C:\Temp or D:\Temp." -ForegroundColor DarkGray
        Write-Host "   Press ENTER to use your system Temp folder automatically." -ForegroundColor Yellow
        Write-Host ""
        $inputBuf = Read-PathInput -Prompt "Paste the TEMP BUFFER path (or press ENTER for $env:TEMP)" -Type "Local" -LogLabel "Leaver Archive Buffer" -AllowEmpty
        if ($inputBuf -eq "Q") { continue }
        if ([string]::IsNullOrWhiteSpace($inputBuf)) { $inputBuf = $env:TEMP; Write-Host "  Using: $inputBuf" -ForegroundColor DarkGray }

        $cleanDst = $inputDst.Trim('"', "'", ' ').TrimEnd('\')
        $cleanBuf = $inputBuf.Trim('"', "'", ' ').TrimEnd('\')

        # Validate buffer
        $bufValid = $true
        try {
            if (-not (Test-Path -LiteralPath $cleanBuf)) { New-Item -ItemType Directory -Path $cleanBuf -Force -ErrorAction Stop | Out-Null }
            $tf = Join-Path $cleanBuf "write_test.tmp"; "test" | Out-File -FilePath $tf -ErrorAction Stop
            Remove-Item $tf -Force -ErrorAction SilentlyContinue
        } catch {
            $bufValid = $false
            Write-Host "  ERROR: Cannot write to temp buffer '$cleanBuf'. Check permissions." -ForegroundColor Red
            Write-Log "Leaver Archive: Buffer path unwritable '$cleanBuf'." "ERROR"
        }
        if (-not $bufValid) { continue }

        if (-not (Test-Path -LiteralPath $cleanDst)) {
            try { New-Item -ItemType Directory -Path $cleanDst -Force | Out-Null }
            catch { Write-Host " ERROR: Could not create '$cleanDst'." -ForegroundColor Red; continue }
        }

        # ── BUILD SITE URL FROM UPN IF NEEDED ────────────────────────────────
        if (-not $leaverSiteUrl -and $leaverUpn) {
            Write-Host ""
            Write-Host "  To build the OneDrive URL from a UPN, enter your tenant name." -ForegroundColor DarkGray
            Write-Host "  This is the prefix of your SharePoint address, e.g.:" -ForegroundColor DarkGray
            Write-Host "  'zellis'  from  zellis.sharepoint.com" -ForegroundColor DarkGray
            $tenantInput = Read-Host "  Tenant name"
            Write-Log "[USER INPUT] Tenant name: '$tenantInput'" "INFO" -FileOnly
            if ([string]::IsNullOrWhiteSpace($tenantInput)) {
                Write-Host "  [!] Tenant name is required to build the OneDrive URL." -ForegroundColor Red
                continue
            }
            $leaverSiteUrl = ConvertTo-OneDriveSiteUrl -TenantName $tenantInput.Trim() -Upn $leaverUpn
            Write-Host "  Constructed URL: $leaverSiteUrl" -ForegroundColor DarkGray
        }

        # ── CONNECT AND ARCHIVE ───────────────────────────────────────────────
        Write-Host ""
        Write-Host "  Connecting to: $leaverSiteUrl" -ForegroundColor DarkGray
        Write-Host "  A browser sign-in window may appear — sign in with your account." -ForegroundColor Yellow

        try {
            $origWarn = $WarningPreference; $WarningPreference = "SilentlyContinue"
            Connect-PnPOnline -Url $leaverSiteUrl -UseWebLogin -ErrorAction Stop
            $WarningPreference = $origWarn
            $testLists = Get-PnPList -ErrorAction Stop
            Write-Host "  [✔] Connected." -ForegroundColor Green
            Write-Log "Leaver Archive: PnP connection succeeded for $leaverSiteUrl" "SUCCESS"

            Write-Log "Leaver Archive: Running archive. Site=$leaverSiteUrl | Dest=$cleanDst" "INFO"
            $archiveResult = Perform-LeaverOneDriveArchive -LeaverSiteUrl $leaverSiteUrl -LocalDest $cleanDst -BufferDest $cleanBuf

            # Show recycle bin link after archive
            $rbUrl = Get-LeaverRecycleBinUrl -LeaverUrl $leaverSiteUrl
            if ($rbUrl) {
                Write-RecycleBinLinks -Links @(@{ Label = $leaverSiteUrl; Url = $rbUrl })
            }

        } catch {
            $WarningPreference = "Continue"
            Write-Log "Leaver Archive: Connection failed for $leaverSiteUrl — $($_.Exception.Message)" "ERROR"
            Write-Host ""
            Write-Host "  [!] Could not connect to this OneDrive." -ForegroundColor Red
            Write-Host "      $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "  This can happen if:" -ForegroundColor Yellow
            Write-Host "   • Your account does not have access to this site." -ForegroundColor DarkGray
            Write-Host "   • The leaver's OneDrive has already been deleted." -ForegroundColor DarkGray
            Write-Host "   • The URL is incorrect — check the site root is right." -ForegroundColor DarkGray
        } finally {
            try { $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
        }

        # v111: use unified press-any-key helper
        Show-PressAnyKeyToContinue
    }


    elseif ($choice -eq '21') {

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor DarkYellow
        Write-Host "  MULTI-SOURCE LEAVER ARCHIVE - Queue Multiple Leavers" -ForegroundColor DarkYellow
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Works exactly like [19] MULTI-SOURCE ARCHIVE but for LEAVER OneDrives."
        Write-Host "   Queue N leavers by UPN or URL with a friendly label, confirm the plan,"
        Write-Host "   then walk away — the script archives every leaver's OneDrive unattended:"
        Write-Host "   Connect → Enumerate Document Libraries → Download → Verify → Delete"
        Write-Host ""
        Write-Host "   D:\\Archives\\2026-03 Leavers\\"
        Write-Host "     ZZ_John_Smith\\    Desktop\   Meetings\   ..." -ForegroundColor DarkGray
        Write-Host "     ZZ_Jane_Doe\\      Desktop\   Recordings\  ..." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ BROWSER SIGN-IN" -ForegroundColor White
        Write-Host "   A sign-in window may appear for the first leaver. Subsequent leavers"
        Write-Host "   on the same tenant usually reuse the session without prompting again."
        Write-Host ""
        Write-Host "  ▸ IF A LEAVER FAILS TO CONNECT" -ForegroundColor White
        Write-Host "   It is skipped and recorded in the outcome table at the end." -ForegroundColor DarkGray
        Write-Host "   All other leavers continue running unaffected." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ SAFETY GATES (same as all archive operations)" -ForegroundColor White
        Write-Host "   Deletion blocked if: any download failure OR file count/byte mismatch." -ForegroundColor DarkGray
        Write-Host "   Deleted items go to the SharePoint Recycle Bin (~93 days recoverable)." -ForegroundColor DarkGray
        Write-Host ""

        # ── DESTINATION & BUFFER ─────────────────────────────────────────────
        $inputDst = Read-PathInput -Prompt "Paste the LOCAL DESTINATION root (one folder for all leaver archives)" -Type "Local" -LogLabel "Multi-Leaver Destination"
        if ($inputDst -eq "Q") { Write-Log "Multi-Leaver Archive cancelled by user." "INFO"; continue }

        Write-Host ""
        Write-Host "  ▸ WHAT IS THE TEMP BUFFER?" -ForegroundColor White
        Write-Host "   Files are staged here before being moved to the final destination."
        Write-Host "   Use a fast local path like C:\Temp or D:\Temp. Press ENTER for system Temp." -ForegroundColor DarkGray
        Write-Host ""
        $inputBuf = Read-PathInput -Prompt "Paste TEMP BUFFER path (or press ENTER for $env:TEMP)" -Type "Local" -LogLabel "Multi-Leaver Buffer" -AllowEmpty
        if ($inputBuf -eq "Q") { continue }
        if ([string]::IsNullOrWhiteSpace($inputBuf)) { $inputBuf = $env:TEMP; Write-Host "  Using: $inputBuf" -ForegroundColor DarkGray }

        $cleanDst = $inputDst.Trim('"', "'", ' ').TrimEnd('\')
        $cleanBuf = $inputBuf.Trim('"', "'", ' ').TrimEnd('\')

        # Validate buffer
        $bufValid = $true
        try {
            if (-not (Test-Path -LiteralPath $cleanBuf)) { New-Item -ItemType Directory -Path $cleanBuf -Force -ErrorAction Stop | Out-Null }
            $tf = Join-Path $cleanBuf "write_test.tmp"; "test" | Out-File -FilePath $tf -ErrorAction Stop
            Remove-Item $tf -Force -ErrorAction SilentlyContinue
        } catch {
            $bufValid = $false
            Write-Host "  ERROR: Cannot write to temp buffer '$cleanBuf'." -ForegroundColor Red
            Write-Log "Multi-Leaver Archive: Buffer path unwritable '$cleanBuf'." "ERROR"
        }
        if (-not $bufValid) { continue }

        # ── COLLECT LEAVER UPNs / URLs ────────────────────────────────────────
        $LeaverList = [System.Collections.Generic.List[hashtable]]::new()
        $tenantName = ""   # remembered once if a UPN needs URL conversion

        Write-Host ""
        Write-Host "  ┌─ ENTER LEAVER UPNs OR URLs ─────────────────────────────" -ForegroundColor DarkYellow
        Write-Host "  │  Paste each leaver's email (UPN) OR their full OneDrive URL." -ForegroundColor DarkGray
        Write-Host "  │  Any /_layouts/ or ?viewid= URL is accepted." -ForegroundColor DarkGray
        Write-Host "  │  Press ENTER on a blank line when done." -ForegroundColor DarkGray
        Write-Host "  │  Type [Q] to cancel." -ForegroundColor DarkGray
        Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor DarkYellow

        $srcNum     = 0
        $cancelFlag = $false
        while ($true) {
            $srcNum++
            Write-Host ""
            $rawInput = Read-Host "  Leaver $srcNum UPN or URL (or ENTER to finish)"
            Write-Log "[USER INPUT] Multi-Leaver Source $srcNum : '$rawInput'" "INFO" -FileOnly

            if ($rawInput -match "^[Qq]$") {
                Write-Host "  Cancelled." -ForegroundColor Yellow
                $cancelFlag = $true; break
            }
            if ([string]::IsNullOrWhiteSpace($rawInput)) {
                Write-Host "  Entry complete — $($LeaverList.Count) leaver(s) queued." -ForegroundColor DarkGray
                break
            }

            $rawInput = $rawInput.Trim()

            # Resolve to a personal site URL
            $resolvedSiteUrl = ""
            if ($rawInput -match "(?i)^https://") {
                $parsedL = Parse-CloudUrl -Url $rawInput
                if ($parsedL) {
                    $resolvedSiteUrl = $parsedL.SiteUrl
                    Write-Host "  Site URL: $resolvedSiteUrl" -ForegroundColor DarkGray
                } else {
                    Write-Host "  [!] Could not parse that URL. Check it starts with https:// and try again." -ForegroundColor Red
                    $srcNum--; continue
                }
            } elseif ($rawInput -match "@") {
                if (-not $tenantName) {
                    $tenantName = (Read-Host "  Tenant name needed to build URL (e.g. 'zellis' from zellis.sharepoint.com)").Trim()
                }
                if ($tenantName) {
                    $resolvedSiteUrl = ConvertTo-OneDriveSiteUrl -TenantName $tenantName -Upn $rawInput
                    Write-Host "  Site URL: $resolvedSiteUrl" -ForegroundColor DarkGray
                } else {
                    Write-Host "  [!] Tenant name required to convert UPN to URL. Try again." -ForegroundColor Red
                    $srcNum--; continue
                }
            } else {
                Write-Host "  [!] Enter a UPN (email) or full https:// URL." -ForegroundColor Red
                $srcNum--; continue
            }

            # Build default label from personal username
            $personalMatch = [regex]::Match($resolvedSiteUrl, '(?i)/personal/([^/]+)')
            $rawName     = if ($personalMatch.Success) { $personalMatch.Groups[1].Value } else { "Leaver$srcNum" }
            $parts       = $rawName -split '_'
            $defaultLbl  = if ($parts.Count -ge 3) { ($parts[0..($parts.Count - 3)] -join '_') } else { $rawName }
            $safeDefault = Get-SafeWindowsName $defaultLbl

            Write-Host "  Default label: '$safeDefault'" -ForegroundColor Green
            $labelInput  = Read-Host "  Friendly label (ENTER to use '$safeDefault', or type e.g. 'ZZ_2026.03_John_Smith')"
            Write-Log "[USER INPUT] Multi-Leaver Source $srcNum Label: '$labelInput'" "INFO" -FileOnly
            $finalLabel  = if ([string]::IsNullOrWhiteSpace($labelInput)) { $safeDefault } else { Get-SafeWindowsName $labelInput.Trim() }

            $LeaverList.Add(@{ SiteUrl = $resolvedSiteUrl; Label = $finalLabel })
            Write-Host "  [✔] Added: '$finalLabel' ← $($resolvedSiteUrl -replace 'https://','')" -ForegroundColor DarkYellow
        }

        if ($cancelFlag -or $LeaverList.Count -eq 0) {
            if (-not $cancelFlag) { Write-Host "  No leavers entered. Returning to main menu." -ForegroundColor Yellow }
            continue
        }

        # ── SHOW PLAN ─────────────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  ┌─ MULTI-LEAVER ARCHIVE PLAN ────────────────────────────" -ForegroundColor DarkYellow
        Write-Host "  │  Destination root : $cleanDst" -ForegroundColor White
        Write-Host "  │" -ForegroundColor DarkYellow
        for ($pi = 0; $pi -lt $LeaverList.Count; $pi++) {
            $e        = $LeaverList[$pi]
            $dispSite = ($e.SiteUrl -replace 'https://','')
            if ($dispSite.Length -gt 50) { $dispSite = "..." + $dispSite.Substring($dispSite.Length - 47) }
            Write-Host "  │  [$($pi+1)] $($e.Label.PadRight(25)) ← $dispSite" -ForegroundColor DarkYellow
        }
        Write-Host "  │" -ForegroundColor DarkYellow
        Write-Host "  │  $($LeaverList.Count) leaver(s) will be archived in sequence." -ForegroundColor DarkGray
        Write-Host "  │  Connect → enumerate libraries → batch archive → next leaver." -ForegroundColor DarkGray
        Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkYellow
        Write-Host ""

        $confirm21 = Read-Host "  Type YES to start unattended run (anything else cancels)"
        Write-Log "[USER INPUT] Multi-Leaver Archive Confirm: '$confirm21'" "INFO" -FileOnly
        if ($confirm21 -cne "YES") {
            Write-Host "  Cancelled." -ForegroundColor Yellow
            Write-Log "Multi-Leaver Archive: Cancelled at confirmation step." "INFO"
            continue
        }

        if (-not (Test-Path -LiteralPath $cleanDst)) {
            try { New-Item -ItemType Directory -Path $cleanDst -Force | Out-Null }
            catch { Write-Host " ERROR: Could not create '$cleanDst'." -ForegroundColor Red; continue }
        }

        Write-Log "Multi-Leaver Archive: Starting $($LeaverList.Count) leaver(s) to '$cleanDst'." "INFO"
        Perform-MultiSourceLeaverArchive `
            -Leavers    $LeaverList `
            -LocalDest  $cleanDst `
            -BufferDest $cleanBuf

        # v111: use unified press-any-key helper
        Show-PressAnyKeyToContinue
    }


    elseif ($choice -eq '22') {

        Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "  LEAVER RECYCLE BIN LINKS - Generate Browser Links" -ForegroundColor Cyan
        Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   Converts leaver OneDrive URLs into the corresponding SharePoint"
        Write-Host "   Recycle Bin URLs, then lets you copy them or open them directly"
        Write-Host "   in your browser to review and permanently empty the Recycle Bin."
        Write-Host ""
        Write-Host "  ▸ HOW TO EMPTY THE RECYCLE BIN" -ForegroundColor White
        Write-Host "   1. Open the generated URL in your browser." -ForegroundColor DarkGray
        Write-Host "   2. Select all items (Ctrl+A or top-left checkbox)." -ForegroundColor DarkGray
        Write-Host "   3. Click 'Delete' → confirm. Items are permanently removed." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ ACCEPTED INPUT FORMATS" -ForegroundColor White
        Write-Host "   • Any /_layouts/15/onedrive.aspx leaver URL" -ForegroundColor DarkGray
        Write-Host "   • A full personal site URL (https://tenant-my.sharepoint.com/personal/...)" -ForegroundColor DarkGray
        Write-Host "   • A leaver email (UPN) — the site URL is constructed automatically" -ForegroundColor DarkGray
        Write-Host ""

        # ── COLLECT URLS ──────────────────────────────────────────────────────
        Write-Host "  ┌─ ENTER LEAVER URLs OR UPNs ─────────────────────────────" -ForegroundColor Cyan
        Write-Host "  │  Paste leaver OneDrive URLs or email addresses one at a time." -ForegroundColor DarkGray
        Write-Host "  │  Press ENTER on a blank line when done.  [Q] to cancel." -ForegroundColor DarkGray
        Write-Host "  └─────────────────────────────────────────────────────────" -ForegroundColor Cyan

        $RbEntries  = [System.Collections.Generic.List[hashtable]]::new()
        $tenantName = ""
        $srcNum     = 0
        $cancelFlag = $false

        while ($true) {
            $srcNum++
            Write-Host ""
            $rawInput = Read-Host "  Leaver $srcNum URL or UPN (or ENTER to finish)"
            Write-Log "[USER INPUT] RecycleBin Source $srcNum : '$rawInput'" "INFO" -FileOnly

            if ($rawInput -match "^[Qq]$") { $cancelFlag = $true; break }
            if ([string]::IsNullOrWhiteSpace($rawInput)) {
                Write-Host "  Entry complete — $($RbEntries.Count) link$(if($RbEntries.Count -ne 1){'s'}) queued." -ForegroundColor DarkGray
                break
            }

            $rawInput        = $rawInput.Trim()
            $resolvedSiteUrl = ""

            if ($rawInput -match "(?i)^https://") {
                $m = [regex]::Match($rawInput, '(?i)(https://[^/]+/personal/[^/?# ]+)')
                if ($m.Success) {
                    $resolvedSiteUrl = $m.Groups[1].Value.TrimEnd('/')
                } else {
                    Write-Host "  [!] Could not find a /personal/ path in that URL." -ForegroundColor Red
                    $srcNum--; continue
                }
            } elseif ($rawInput -match "@") {
                if (-not $tenantName) {
                    $tenantName = (Read-Host "  Tenant name needed (e.g. 'zellis' from zellis.sharepoint.com)").Trim()
                }
                if ($tenantName) {
                    $resolvedSiteUrl = ConvertTo-OneDriveSiteUrl -TenantName $tenantName -Upn $rawInput
                } else {
                    Write-Host "  [!] Tenant name required. Try again." -ForegroundColor Red
                    $srcNum--; continue
                }
            } else {
                Write-Host "  [!] Enter a leaver URL (https://...) or email address." -ForegroundColor Red
                $srcNum--; continue
            }

            $rbUrl = Get-LeaverRecycleBinUrl -LeaverUrl $resolvedSiteUrl
            if (-not $rbUrl) {
                Write-Host "  [!] Could not build recycle bin URL. Check the URL contains /personal/." -ForegroundColor Red
                $srcNum--; continue
            }

            # Use the username segment as the label (shown in multi-link display)
            $autoLabel = ($resolvedSiteUrl -split '/')[-1]
            $RbEntries.Add(@{ Label = $autoLabel; SiteUrl = $resolvedSiteUrl; Url = $rbUrl })
            Write-Host "  [✔] $rbUrl" -ForegroundColor Cyan
        }

        if ($cancelFlag -or $RbEntries.Count -eq 0) {
            if (-not $cancelFlag) { Write-Host "  No entries. Returning to main menu." -ForegroundColor Yellow }
            continue
        }

        # ── DISPLAY RESULTS ───────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  ┌─ RECYCLE BIN LINK$(if($RbEntries.Count -gt 1){'S'}) ─────────────────────────────" -ForegroundColor Cyan
        Write-Host "  │  Select all → Delete in your browser to permanently clear each bin." -ForegroundColor DarkGray
        Write-Host "  │" -ForegroundColor Cyan
        $logUrls = @()
        foreach ($entry in $RbEntries) {
            Write-Host "  │  $($entry.Url)" -ForegroundColor Cyan
            $logUrls += $entry.Url
        }
        Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor Cyan
        Write-Log "Recycle Bin Links generated: $($logUrls -join ' | ')" "INFO"

        # ── POST-ACTIONS ──────────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  [C] Copy to clipboard   [O] Open in browser   [B] Both   ENTER = skip" -ForegroundColor DarkGray
        $rbAction = (Read-Host "  Action").Trim().ToUpper()

        if ($rbAction -in @('C','B')) {
            try {
                $clipText = ($RbEntries | ForEach-Object { $_.Url }) -join "`r`n"
                Set-Clipboard -Value $clipText -ErrorAction Stop
                Write-Host "  [✔] Copied $($RbEntries.Count) link$(if($RbEntries.Count -ne 1){'s'}) to clipboard." -ForegroundColor Green
                Write-Log "Recycle Bin Links: Copied $($RbEntries.Count) link(s) to clipboard." "INFO"
            } catch {
                Write-Host "  [!] Clipboard copy failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if ($rbAction -in @('O','B')) {
            Write-Host "  Opening $($RbEntries.Count) link$(if($RbEntries.Count -ne 1){'s'}) in your browser..." -ForegroundColor DarkGray
            foreach ($entry in $RbEntries) {
                try {
                    Start-Process $entry.Url -ErrorAction Stop
                    Write-Log "Recycle Bin Links: Opened in browser: $($entry.Url)" "INFO"
                    if ($RbEntries.Count -gt 1) { Start-Sleep -Milliseconds 800 }
                } catch {
                    Write-Host "  [!] Could not open: $($entry.Url)" -ForegroundColor Yellow
                }
            }
            Write-Host "  [✔] Done — check your browser." -ForegroundColor Green
        }

        # v111: use unified press-any-key helper
        Show-PressAnyKeyToContinue
    }


    # ════════════════════════════════════════════════════════════════════════
    # OPTION [23]: UNLICENSED LEAVER ONEDRIVE ARCHIVE
    # ════════════════════════════════════════════════════════════════════════
    # For leavers whose E5/M365 licence was removed and whose OneDrive is
    # now frozen. Temporarily assigns a Plan 1 licence via an Azure AD group,
    # waits for the OneDrive to become accessible, archives all content
    # (download → verify → delete to Recycle Bin), then removes the user
    # from the group to free the licence slot for the next leaver.
    elseif ($choice -eq '23') {

        Write-Host ""
        Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "  ║   [23] UNLICENSED LEAVER ONEDRIVE ARCHIVE               ║" -ForegroundColor Cyan
        Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  ▸ WHAT THIS DOES" -ForegroundColor White
        Write-Host "   When a leaver's E5 or M365 licence is removed, their personal OneDrive" -ForegroundColor DarkGray
        Write-Host "   becomes frozen and is no longer accessible via SharePoint admin tools." -ForegroundColor DarkGray
        Write-Host "   This option temporarily grants a SharePoint Plan 1 licence through an" -ForegroundColor DarkGray
        Write-Host "   Azure AD security group, waits for the OneDrive to become readable," -ForegroundColor DarkGray
        Write-Host "   downloads and verifies all content locally, then deletes the cloud copy" -ForegroundColor DarkGray
        Write-Host "   to the Recycle Bin. The user is removed from the licence group the moment" -ForegroundColor DarkGray
        Write-Host "   archiving is done — so you only ever consume one licence slot at a time." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ REQUIREMENTS" -ForegroundColor White
        Write-Host "   • A SharePoint Plan 1 licence assigned to an Azure AD security group." -ForegroundColor DarkGray
        Write-Host "   • You need GroupMember.ReadWrite.All Graph API access (sign-in prompted)." -ForegroundColor DarkGray
        Write-Host "   • One licence in the group is consumed per leaver being actively archived." -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  ▸ LEAVER LIST FORMAT" -ForegroundColor White
        Write-Host "   Paste tab-separated rows — one leaver per line.  Two accepted formats:" -ForegroundColor DarkGray
        Write-Host "   3-col:  Email  [TAB]  OneDrive URL  [TAB]  Folder name" -ForegroundColor DarkGray
        Write-Host "   6-col:  Email  [TAB]  URL  [TAB]  QuotaGB  [TAB]  UsedGB  [TAB]  %  [TAB]  Folder name" -ForegroundColor DarkGray
        Write-Host "   (6-column is the raw SharePoint usage report format — quota/used/% are ignored)" -ForegroundColor DarkGray
        Write-Host "   Press [V] at the leaver-list prompt to paste the entire list from the clipboard." -ForegroundColor DarkGray
        Write-Host ""

        $cancel23 = $false

        # ── INPUT 1: LOCAL DESTINATION FOLDER ─────────────────────────────────
        Write-Host "  ─── Step 1 of 4 — Local destination folder ────────────────" -ForegroundColor White
        Write-Host "  Each leaver gets a subfolder here named after their folder label." -ForegroundColor DarkGray
        Write-Host "  Example:  D:\LeaverArchives  →  D:\LeaverArchives\ZZ_2025.12.06 Bruno Cambrais\" -ForegroundColor DarkGray
        Write-Host ""
        $localDest23 = Read-PathInput -Prompt "  Local destination folder" -Type "Local" -LogLabel "Unlicensed Leaver Destination"
        if ($localDest23 -eq "Q" -or [string]::IsNullOrWhiteSpace($localDest23)) { continue }
        $localDest23 = $localDest23.Trim('"', "'", ' ').TrimEnd('\')

        if (-not (Test-Path -LiteralPath $localDest23)) {
            try {
                New-Item -ItemType Directory -Path $localDest23 -Force | Out-Null
                Write-Host "  [✔] Created: $localDest23" -ForegroundColor Green
            } catch {
                Write-Host "  [✗] Cannot create folder: $($_.Exception.Message)" -ForegroundColor Red
                continue
            }
        } else {
            Write-Host "  [✔] Folder exists: $localDest23" -ForegroundColor DarkGray
        }

        # ── INPUT 2: BUFFER / TEMP FOLDER ─────────────────────────────────────
        Write-Host ""
        Write-Host "  ─── Step 2 of 4 — Buffer / temp folder ────────────────────" -ForegroundColor White
        Write-Host "  Temporary staging area for downloads. Must be on the same drive as" -ForegroundColor DarkGray
        Write-Host "  the destination for fast atomic file promotion." -ForegroundColor DarkGray
        Write-Host ""
        $buffer23raw = Read-PathInput -Prompt "  Buffer / temp folder  (ENTER = use $env:TEMP)" -Type "Local" -LogLabel "Unlicensed Leaver Buffer" -AllowEmpty
        if ($buffer23raw -eq "Q") { continue }
        $buffer23 = if ([string]::IsNullOrWhiteSpace($buffer23raw)) { $env:TEMP } else { $buffer23raw.Trim('"', "'", ' ').TrimEnd('\') }
        Write-Host "  Using buffer: $buffer23" -ForegroundColor DarkGray

        # Verify the buffer path is actually writable before committing to a long run
        try {
            if (-not (Test-Path -LiteralPath $buffer23)) { New-Item -ItemType Directory -Path $buffer23 -Force -ErrorAction Stop | Out-Null }
            $testFile = Join-Path $buffer23 "write_test_23.tmp"
            "test" | Out-File -FilePath $testFile -ErrorAction Stop
            Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Host "  [✗] Cannot write to buffer path '$buffer23': $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "      Choose a different path or press ENTER on a fresh attempt to use $env:TEMP." -ForegroundColor DarkGray
            continue
        }

        # ── INPUT 3: LICENCE GROUP NAME ────────────────────────────────────────
        Write-Host ""
        # ── STEP 3: LICENCE GROUP (retry loop — wrong name re-prompts, no menu exit) ──
        # Graph is connected once and stays connected. The inner loop lets the user
        # correct a typo in the group name without re-entering paths or re-signing-in.
        Write-Host "  ─── Step 3 of 4 — Azure AD licence group ───────────────────" -ForegroundColor White
        Write-Host "  Enter the exact display name of the Azure AD security group" -ForegroundColor DarkGray
        Write-Host "  that has the SharePoint Plan 1 licence assigned to it." -ForegroundColor DarkGray
        Write-Host ""

        # Connect to Graph once (outside the group retry loop — signing in twice is wasteful)
        $mgOk = Ensure-MgGraphForLicensing
        if (-not $mgOk) {
            Write-Host ""
            Write-Host "  [✗] Cannot connect to Microsoft Graph. Returning to menu." -ForegroundColor Red
            Write-Host "  Tip: Check your internet connection and run option [23] again." -ForegroundColor DarkGray
            continue
        }

        $licenseGroup23     = $null
        $licenseGroupName23 = ""
        while ($null -eq $licenseGroup23 -and -not $cancel23) {

            # Prompt for group name (re-shown on every retry)
            while ([string]::IsNullOrWhiteSpace($licenseGroupName23) -and -not $cancel23) {
                $licenseGroupName23 = (Read-Host "  Licence group name  (Q to cancel)").Trim()
                if ($licenseGroupName23 -match '^[Qq]$') { $cancel23 = $true; break }
                if ([string]::IsNullOrWhiteSpace($licenseGroupName23)) {
                    Write-Host "  Please enter the group name." -ForegroundColor Yellow
                }
            }
            if ($cancel23) { break }

            Write-Host "  Resolving '$licenseGroupName23'..." -ForegroundColor DarkGray
            try {
                $safeGrp23 = [uri]::EscapeDataString("displayName eq '$($licenseGroupName23 -replace "'","''")'")
                $grpResult = Invoke-MgRestMethod -Method "GET" -Endpoint "groups?`$filter=$safeGrp23&`$select=id,displayName"
                $grpList   = @($grpResult.value)

                if ($grpList.Count -eq 0) {
                    Write-Host "  [✗] No group named '$licenseGroupName23' found in Entra ID." -ForegroundColor Red
                    Write-Host "  Tip: Group names are case-sensitive. Check the exact name in Azure AD portal." -ForegroundColor DarkGray
                    Write-Host "       Type Q to cancel, or enter the correct name below." -ForegroundColor DarkGray
                    Write-Host ""
                    $licenseGroupName23 = ""   # clear so the inner while re-prompts
                    continue
                }

                if ($grpList.Count -gt 1) {
                    Write-Host "  [!] Multiple groups match — using the first: $($grpList[0].displayName)" -ForegroundColor Yellow
                }
                $licenseGroup23 = $grpList[0]
                Write-Host "  [✔] Group: $($licenseGroup23.displayName)  (ID: $($licenseGroup23.id))" -ForegroundColor Green

            } catch {
                Write-Host "  [✗] Group look-up failed: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "      Type Q to cancel, or try again with a different name." -ForegroundColor DarkGray
                Write-Host ""
                $licenseGroupName23 = ""   # clear to re-prompt
            }
        }
        if ($cancel23 -or $null -eq $licenseGroup23) { continue }

        # ── STEP 4: LEAVER LIST (retry loop — blank first row warns, parse-fail re-enters) ─
        # The outer while keeps the user in leaver-list entry until they either provide
        # at least one valid row or press Q to cancel back to the main menu.
        Write-Host ""
        Write-Host "  ─── Step 4 of 4 — Leaver list ─────────────────────────────" -ForegroundColor White
        Write-Host "  Paste rows, one per line (Tab-separated).  Press ENTER on a blank line when done." -ForegroundColor DarkGray
        Write-Host "  Press [V] to paste the entire list from the clipboard in one go." -ForegroundColor DarkGray
        Write-Host "  Format:  Email [TAB] URL [TAB] Folder name" -ForegroundColor DarkGray
        Write-Host "  (6-column report format also accepted — quota / used / % columns are ignored)" -ForegroundColor DarkGray
        Write-Host ""

        $parsedLeavers23 = [System.Collections.Generic.List[hashtable]]::new()

        while ($parsedLeavers23.Count -eq 0 -and -not $cancel23) {

            $leaverRows = [System.Collections.Generic.List[string]]::new()
            $rowNum     = 0

            while (-not $cancel23) {
                $rowNum++
                $rawRow = (Read-Host "  Row $rowNum  (V=clipboard  ENTER=done  Q=cancel)").Trim()

                if ($rawRow -match '^[Qq]$') { $cancel23 = $true; break }

                # Blank line = done — but only if at least one row has been entered
                if ([string]::IsNullOrWhiteSpace($rawRow)) {
                    if ($leaverRows.Count -eq 0) {
                        Write-Host "  Nothing entered yet. Type a row, press [V] to paste from clipboard, or [Q] to cancel." -ForegroundColor Yellow
                        $rowNum--; continue
                    }
                    break
                }

                # ── V: CLIPBOARD PASTE ────────────────────────────────────────────
                if ($rawRow -match '^[Vv]$') {
                    try {
                        $clipData = Get-Clipboard -ErrorAction Stop
                        $clipRows = ($clipData -join "`n") -split '[\r\n]+' | Where-Object { $_.Trim() -ne '' }
                        if ($clipRows.Count -eq 0) {
                            Write-Host "  Clipboard is empty. Copy your data first, then press V." -ForegroundColor Yellow
                            $rowNum--; continue
                        }
                        $suspectRows = @($clipRows | Where-Object { $_ -match '[✓✗║╔╚═─┌┐└┘│►◄▲▼•·]' })
                        if ($suspectRows.Count -gt 0) {
                            Write-Host "  [!] Clipboard appears to contain script output, not spreadsheet data." -ForegroundColor Yellow
                            Write-Host "      Paste from Excel or your report, not from this terminal window." -ForegroundColor DarkGray
                            $rowNum--; continue
                        }
                        Write-Host ""
                        Write-Host "  Captured $($clipRows.Count) row(s) from clipboard:" -ForegroundColor Cyan
                        foreach ($cr in $clipRows) {
                            $cols    = $cr -split "`t"
                            $preview = if ($cols.Count -ge 3) {
                                "$($cols[0])  |  $($cols[1] -replace 'https://[^/]+-my\.sharepoint\.com','...')  |  $($cols[-1])"
                            } else { $cr }
                            if ($preview.Length -gt 120) { $preview = $preview.Substring(0, 117) + "..." }
                            Write-Host "    $preview" -ForegroundColor DarkGray
                        }
                        Write-Host ""
                        $clipConfirm = (Read-Host "  Use these $($clipRows.Count) row(s)?  (Y = Yes  /  N = Cancel)").Trim()
                        Write-Log "[USER INPUT] [23] Clipboard confirm: $clipConfirm" "INFO" -FileOnly
                        if ($clipConfirm -match '^[Yy]$') {
                            foreach ($cr in $clipRows) { $leaverRows.Add($cr) }
                            Write-Host "  [✔] $($clipRows.Count) row(s) added from clipboard." -ForegroundColor Green
                            break
                        } else {
                            Write-Host "  Clipboard paste cancelled. Enter rows manually below." -ForegroundColor DarkGray
                            $rowNum--; continue
                        }
                    } catch {
                        Write-Host "  Could not read clipboard: $($_.Exception.Message)" -ForegroundColor Yellow
                        $rowNum--; continue
                    }
                    continue
                }

                # Normal typed row
                $leaverRows.Add($rawRow)
            }

            if ($cancel23) { break }

            # ── PARSE ROWS ────────────────────────────────────────────────────────
            $parseErrors23 = [System.Collections.Generic.List[string]]::new()
            $parsedLeavers23.Clear()

            foreach ($row in $leaverRows) {
                $cols = $row -split "`t" | ForEach-Object { $_.Trim() }
                if ($cols.Count -lt 3) {
                    $snippet = if ($row.Length -gt 80) { $row.Substring(0,77) + "..." } else { $row }
                    $parseErrors23.Add("Fewer than 3 columns — skipped: '$snippet'")
                    continue
                }
                $pEmail  = $cols[0]
                $pUrl    = $cols[1]
                $pFolder = if ($cols.Count -ge 6) { $cols[5] } else { $cols[2] }

                if ($pEmail -notmatch '@')          { $parseErrors23.Add("Col 1 missing @ — not an email: '$pEmail'"); continue }
                if ($pUrl -notmatch '^https://')    { $parseErrors23.Add("Col 2 missing https:// — not a URL: '$pUrl'"); continue }
                if ([string]::IsNullOrWhiteSpace($pFolder)) { $parseErrors23.Add("Folder name blank for: $pEmail"); continue }

                if ($pUrl -match '(https://[^/]+-my\.sharepoint\.com/personal/[^/\?#]+)') { $pUrl = $matches[1] }
                $parsedLeavers23.Add(@{ Email = $pEmail; SiteUrl = $pUrl; FolderName = $pFolder })
            }

            # Show parse errors
            if ($parseErrors23.Count -gt 0) {
                Write-Host ""
                Write-Host "  [!] $($parseErrors23.Count) row(s) could not be parsed:" -ForegroundColor Yellow
                foreach ($pe in $parseErrors23) { Write-Host "      $pe" -ForegroundColor DarkGray }
            }

            # If nothing valid parsed, stay in the outer while to re-enter the list
            if ($parsedLeavers23.Count -eq 0) {
                Write-Host ""
                Write-Host "  No valid leaver rows found — please check the format and try again." -ForegroundColor Red
                Write-Host "  Format reminder:  Email [TAB] URL [TAB] Folder name" -ForegroundColor DarkGray
                Write-Host "  Press [V] to paste from clipboard, or [Q] to cancel." -ForegroundColor DarkGray
                Write-Host ""
                # Loop back — outer while re-enters row collection
            }
        }

        if ($cancel23 -or $parsedLeavers23.Count -eq 0) { continue }

        # ── CONFIRMATION ───────────────────────────────────────────────────────
        Write-Host ""
        Write-Host "  ┌─ CONFIRMATION ─────────────────────────────────────────────" -ForegroundColor Cyan
        Write-Host "  │  Leavers to process : $($parsedLeavers23.Count)" -ForegroundColor Cyan
        Write-Host "  │  Licence group      : $($licenseGroup23.displayName)" -ForegroundColor Cyan
        Write-Host "  │  Local destination  : $localDest23" -ForegroundColor Cyan
        Write-Host "  │  Buffer folder      : $buffer23" -ForegroundColor Cyan
        Write-Host "  │" -ForegroundColor Cyan
        for ($pi = 0; $pi -lt $parsedLeavers23.Count; $pi++) {
            $lv = $parsedLeavers23[$pi]
            Write-Host "  │  [$($pi+1)]  $($lv.FolderName)" -ForegroundColor DarkYellow
            Write-Host "  │       $($lv.Email)" -ForegroundColor DarkGray
            Write-Host "  │       $($lv.SiteUrl -replace 'https://','')" -ForegroundColor DarkGray
        }
        Write-Host "  │" -ForegroundColor Cyan
        Write-Host "  │  PER LEAVER — what will happen:" -ForegroundColor White
        Write-Host "  │   1. Added to [$($licenseGroup23.displayName)] → Plan 1 licence assigned" -ForegroundColor DarkGray
        Write-Host "  │   2. Script polls OneDrive every 30 s until accessible (max 30 min)" -ForegroundColor DarkGray
        Write-Host "  │   3. All content downloaded to LocalDest\FolderName and verified" -ForegroundColor DarkGray
        Write-Host "  │   4. Cloud content deleted to SharePoint Recycle Bin" -ForegroundColor DarkGray
        Write-Host "  │   5. Removed from group → licence slot freed for the next leaver" -ForegroundColor DarkGray
        Write-Host "  └─────────────────────────────────────────────────────────────" -ForegroundColor Cyan
        Write-Host ""

        $confirm23 = (Read-Host "  Proceed with $($parsedLeavers23.Count) leaver(s)?  (Y / N)").Trim()
        Write-Log "[USER INPUT] [23] Confirm: $confirm23 — $($parsedLeavers23.Count) leavers, Group: $($licenseGroup23.displayName)" "INFO" -FileOnly

        if ($confirm23 -notmatch '^[Yy]$') {
            Write-Host "  Cancelled — no changes made." -ForegroundColor DarkGray
            continue
        }

        # ── RUN THE ENGINE ─────────────────────────────────────────────────────
        Write-Log "[23] UNLICENSED LEAVER ARCHIVE START: $($parsedLeavers23.Count) leavers | Group: $($licenseGroup23.displayName) | Dest: $localDest23" "INFO"

        Perform-LicensedLeaverArchive `
            -Leavers   $parsedLeavers23 `
            -LocalDest  $localDest23 `
            -BufferDest $buffer23 `
            -GroupId    $licenseGroup23.id `
            -GroupName  $licenseGroup23.displayName

        # v111: use unified press-any-key helper
        Show-PressAnyKeyToContinue
    }


    elseif ($choice -match '^[Qq]$') {
        Write-Host " Closing the script. All actions have been recorded in:" -ForegroundColor Cyan
        Write-Host " $LogFile" -ForegroundColor Cyan
        Write-Host " Goodbye, $ActiveUser!" -ForegroundColor Cyan
        break
    }
}