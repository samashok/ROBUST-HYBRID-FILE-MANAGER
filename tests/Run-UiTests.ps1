<#
    Screen-flow tests for Cross-Tenant-Cloud-FileManager.ps1.
    Drives the menu and every main option with scripted keyboard input against temp folders,
    the in-memory fake SharePoint (MockPnP.ps1) and a fake Microsoft Graph.
    Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-UiTests.ps1
#>
param([string]$Script = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Cross-Tenant-Cloud-FileManager.ps1'))
$root = Join-Path ([System.IO.Path]::GetTempPath()) 'rhfm-ui-tests'
if (Test-Path $root) { Remove-Item $root -Recurse -Force }
$null = New-Item -ItemType Directory -Path $root
. (Join-Path $PSScriptRoot 'MockPnP.ps1')
. $Script -WorkDir "$root/wd" -LibraryMode
$script:PnPReady = $true
$script:Settings.DefaultBuffer = "$root/buf"

$global:Pass = 0; $global:Fail = 0
function Assert { param([bool]$Cond, [string]$Name) if ($Cond) { $global:Pass++; Write-Output "PASS  $Name" } else { $global:Fail++; Write-Output "FAIL  $Name" } }

# Scripted keyboard
$global:Inputs = New-Object System.Collections.Generic.Queue[string]
$global:Prompts = New-Object System.Collections.Generic.List[string]
function Read-Line { param([string]$Prompt = '  > ', [string]$LogLabel = '', [switch]$NoLog, [string]$Color = 'Cyan')
    $global:Prompts.Add($Prompt)
    if ($global:Inputs.Count -eq 0) { throw "INPUT EXHAUSTED at prompt: $Prompt" }
    return $global:Inputs.Dequeue() }
function Show-Pause { param([string]$Message) }
function Feed { param([string[]]$Lines) $global:Inputs.Clear(); $global:Prompts.Clear(); foreach ($l in $Lines) { $global:Inputs.Enqueue($l) } }
function Run { param([scriptblock]$B) try { & $B 6>$null; return $null } catch { return $_.Exception.Message } }

# --- Local copy through the option screen ---
$src = "$root/src"; $null = New-Item -ItemType Directory "$src/a" -Force; [IO.File]::WriteAllText("$src/a/x.txt", 'x'); [IO.File]::WriteAllText("$src/y.txt", 'yy')
Feed @("`"$src`"", "$root/dst", 'm')
$e = Run { Invoke-OpLocalCopyMove }
Assert (-not $e -and (Test-Path "$root/dst/a/x.txt") -and $global:Inputs.Count -eq 0) "option 7 local copy via prompts $e"

# --- Main menu: keyword search, number, quit ---
Feed @('verify', '11', 'q', 'checksum', 'n', '7', 'q', 'zzz', 'h', 'q')
$e = Run { Start-MainMenu }
Assert (-not $e -and $global:Inputs.Count -eq 0) "main menu keyword search + quit $e"

# --- Local batch: pick item 1, verify ---
Feed @($src, "$root/dst", '1-2', 'V', 'q', 'm')
$e = Run { Invoke-OpLocalBatch }
Assert (-not $e -and $global:Inputs.Count -eq 0) "option 9 local batch verify $e"

# --- Local batch delete needs YES ---
$del = "$root/del"; $null = New-Item -ItemType Directory "$del/k" -Force; [IO.File]::WriteAllText("$del/k/f.txt", 'f'); [IO.File]::WriteAllText("$del/g.txt", 'g')
Feed @($del, '', 'A', 'D', 'yes', 'A', 'D', 'YES', 'm')
$e = Run { Invoke-OpLocalBatch }
Assert (-not $e -and -not (Test-Path "$del/k") -and -not (Test-Path "$del/g.txt")) "option 9 delete requires exact YES $e"

# --- Local wipe refuses protected, then wipes ---
$wp = "$root/wipe"; $null = New-Item -ItemType Directory "$wp/z" -Force; [IO.File]::WriteAllText("$wp/z/q.txt", 'q')
Feed @("$root/wd", 'x', $wp, 'YES', 'm')
$e = Run { Invoke-OpLocalWipe }
Assert (-not $e -and -not (Test-Path $wp) -and (Test-Path "$root/wd")) "option 10 wipe (protected refused) $e"

# --- Favourites manager add / rename / move / delete, then [F] at a prompt ---
Feed @('A', '2', $src, '', 'A', '1', 'https://contoso.sharepoint.com/sites/HR/Documents/A', 'HR A', 'R', '1', 'Source', 'N', '1', 'q')
$e = Run { Manage-Favorites }
$f = Get-Favorites
Assert (-not $e -and $f.Count -eq 2 -and $f[0].Name -eq 'HR A' -and $f[1].Name -eq 'Source') "favourites manager $e"
Feed @('f', '1')
$v = $null; $e = Run { $script:v = Read-PathInput -Prompt 'x' -Type Cloud }
Assert ($script:v -eq 'https://contoso.sharepoint.com/sites/HR/Documents/A') "F picker at a prompt ($($script:v))"
Feed @('r')
$e = Run { $script:v = Read-PathInput -Prompt 'x' -Type Cloud }
Assert ($script:v -eq 'https://contoso.sharepoint.com/sites/HR/Documents/A') 'R recent at a prompt'
Feed @('C:\local', 'https://contoso.sharepoint.com/sites/HR/Documents', 'q')
$e = Run { $script:v = Read-PathInput -Prompt 'x' -Type Cloud }
Assert ($script:v -eq 'https://contoso.sharepoint.com/sites/HR/Documents') 'cloud prompt rejects local path then accepts URL'

# --- Cloud download (13) and batch download (14) ---
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFolder '/sites/HR/Documents/P'; Add-FakeFile '/sites/HR/Documents/P/r.txt' 'rr'; Add-FakeFile '/sites/HR/Documents/t.txt' 't'
Feed @('https://contoso.sharepoint.com/sites/HR/Documents/Forms/AllItems.aspx', "$root/down", '', 'm')
$e = Run { Invoke-OpDownload }
Assert (-not $e -and (Test-Path "$root/down/P/r.txt") -and (Test-Path "$root/down/t.txt")) "option 13 download $e"
Feed @('https://contoso.sharepoint.com/sites/HR/Documents', "$root/down2", '', '2', 'q', 'm')
$e = Run { Invoke-OpDownload -Batch }
Assert (-not $e -and (Test-Path "$root/down2/t.txt") -and -not (Test-Path "$root/down2/P")) "option 14 batch download pick 2 $e"

# --- Buffer is asked every time; a typed buffer is used and offered next time ---
$custom = "$root/mybuffer"
Feed @('https://contoso.sharepoint.com/sites/HR/Documents', "$root/down3", $custom, 'm')
$e = Run { Invoke-OpDownload }
Assert (-not $e -and (Test-Path $custom) -and ($global:Prompts -match 'Temp buffer|> ').Count -ge 3 -and $script:LastBuffer -eq $custom) "buffer asked and custom buffer used $e"

# --- Cloud batch copy with destination asked later ---
Feed @('https://contoso.sharepoint.com/sites/HR/Documents', '', '1', 'C', 'https://contoso.sharepoint.com/sites/HR/Documents/Copies', 'q', 'm')
$e = Run { Invoke-OpCloudBatch }
Assert (-not $e -and $global:Fake.ContainsKey('/sites/hr/documents/copies/p/r.txt')) "option 3 cloud batch copy folder INTO destination $e"

# --- Cloud wipe of a library root needs YES and keeps the library ---
Feed @('https://contoso.sharepoint.com/sites/HR/Documents', 'YES', 'm')
$e = Run { Invoke-OpCloudWipe }
Assert (-not $e -and $global:Fake.ContainsKey('/sites/hr/documents') -and -not $global:Fake.ContainsKey('/sites/hr/documents/t.txt')) "option 4 wipe library root keeps library $e"

# --- Archive (17) with YES ---
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFolder '/sites/HR/Documents/Old'; Add-FakeFile '/sites/HR/Documents/Old/o.txt' 'old'
Feed @('https://contoso.sharepoint.com/sites/HR/Documents/Old', "$root/arch", '', 'm')
$e = Run { Invoke-OpArchive }
Assert (-not $e -and (Test-Path "$root/arch/o.txt") -and -not $global:Fake.ContainsKey('/sites/hr/documents/old')) "option 17 archive $e"
# Cancelling at the buffer question changes nothing
Add-FakeFolder '/sites/HR/Documents/Keep'; Add-FakeFile '/sites/HR/Documents/Keep/k.txt' 'k'
Feed @('https://contoso.sharepoint.com/sites/HR/Documents/Keep', "$root/arch2", 'q')
$e = Run { Invoke-OpArchive }
Assert (-not $e -and $global:Fake.ContainsKey('/sites/hr/documents/keep/k.txt') -and $global:Inputs.Count -eq 0) "option 17 cancelled at buffer leaves cloud untouched $e"

# --- Batch archive (18) ---
Add-FakeFolder '/sites/HR/Documents/B1'; Add-FakeFile '/sites/HR/Documents/B1/b.txt' 'b'
Add-FakeFile '/sites/HR/Documents/loose.txt' 'loose'
Feed @('https://contoso.sharepoint.com/sites/HR/Documents', "$root/barch", '', 'A', 'q', 'm')
$e = Run { Invoke-OpArchive -Batch }
Assert (-not $e -and (Test-Path "$root/barch/B1/b.txt") -and (Test-Path "$root/barch/loose.txt") -and -not $global:Fake.ContainsKey('/sites/hr/documents/loose.txt')) "option 18 batch archive $e"

# --- Multi-source archive (19) with duplicate label rejected ---
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFolder '/sites/HR/Documents/2023'; Add-FakeFile '/sites/HR/Documents/2023/a.txt' 'a'
Add-FakeFolder '/sites/HR/Shared'; Add-FakeFolder '/sites/HR/Shared/2023'; Add-FakeFile '/sites/HR/Shared/2023/b.txt' 'b'
Feed @("$root/multi", '', 'https://contoso.sharepoint.com/sites/HR/Documents/2023', '', 'https://contoso.sharepoint.com/sites/HR/Shared/2023', '', '')
$e = Run { Invoke-OpMultiArchive }
Assert (-not $e -and (Test-Path "$root/multi/2023/a.txt") -and (Test-Path "$root/multi/2023 (2)/b.txt")) "option 19 multi-source with unique folders $e"

# --- Leaver (20) by UPN ---
Reset-Fake
$global:FakeSitePath = '/personal/ann_contoso_com'
Add-FakeFolder '/personal/ann_contoso_com/Documents'; Add-FakeFile '/personal/ann_contoso_com/Documents/cv.pdf' 'cv'
Add-FakeFolder '/personal/ann_contoso_com/Apps'
Feed @('ann@contoso.com', 'contoso', "$root/leaver", '', '', 'm')
$e = Run { Invoke-OpLeaver }
Assert (-not $e -and (Test-Path "$root/leaver/cv.pdf") -and $global:FakeCalls -contains 'Connect https://contoso-my.sharepoint.com/personal/ann_contoso_com') "option 20 leaver by UPN $e"

# --- Recycle links (22) ---
Feed @('bob@contoso.com', '', 'https://contoso-my.sharepoint.com/personal/x_y_com/_layouts/15/onedrive.aspx', 'https://contoso.sharepoint.com/sites/Finance', '', '')
$e = Run { Invoke-OpRecycleLinks }
Assert (-not $e -and $global:Inputs.Count -eq 0) "option 22 links $e"

# --- Unlicensed leaver (23) with mocked Graph ---
$global:Graph = New-Object System.Collections.Generic.List[string]
$global:InGroup = $false
function Invoke-RestMethod {
    param($Method, $Uri, $Body, $ContentType, $Headers, $ErrorAction)
    $global:Graph.Add("$Method $Uri")
    if ($Uri -match 'checkMemberGroups') { return [pscustomobject]@{ value = $(if ($global:InGroup) { @('g1') } else { @() }) } }
    if ($Uri -match 'devicecode$') { return [pscustomobject]@{ device_code = 'dc'; user_code = 'ABCD'; verification_uri = 'https://microsoft.com/devicelogin'; interval = 0; expires_in = 900; message = 'm' } }
    if ($Uri -match 'oauth2/v2.0/token$') {
        $payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"upn":"admin@contoso.com"}')).TrimEnd('=').Replace('+','-').Replace('/','_')
        return [pscustomobject]@{ access_token = "h.$payload.s"; refresh_token = 'rt'; expires_in = 3600 } }
    if ($Uri -match '/users/ann%40contoso.com') { return [pscustomobject]@{ id = 'u1'; displayName = 'Ann' } }
    if ($Uri -match '/groups\?') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'g1'; displayName = 'SPO Plan 1' }, [pscustomobject]@{ id = 'g2'; displayName = 'SPO Plan 1'; description = 'old' }) } }
    if ($Method -eq 'POST' -and $Uri -match '/groups/g1/members/\$ref$') { $global:InGroup = $true; return $null }
    if ($Method -eq 'DELETE' -and $Uri -match '/groups/g1/members/u1/\$ref$') { $global:InGroup = $false; return $null }
    if ($Uri -match 'licenseAssignmentStates') { return [pscustomobject]@{ licenseAssignmentStates = @() } }
    throw "unexpected graph call $Method $Uri"
}
function Start-InterruptibleSleep { param([int]$Seconds, [string]$Label) }
function Set-ClipboardText { param($Text) return $false }
Reset-Fake
Add-FakeFolder '/personal/ann_contoso_com/Documents'; Add-FakeFile '/personal/ann_contoso_com/Documents/z.txt' 'z'
Feed @("$root/lic", '', 'SPO Plan 1', '1', "ann@contoso.com`thttps://contoso-my.sharepoint.com/personal/ann_contoso_com`tAnn Smith", '', 'y', '')
$e = Run { Invoke-OpLicensedLeaver }
Assert (-not $e -and (Test-Path "$root/lic/Ann Smith/z.txt") -and -not $global:InGroup -and ($global:Graph -match 'DELETE .*/members/u1/\$ref').Count -eq 1) "option 23 licence add -> archive -> remove ($e)"
Assert (Test-Path $script:MgTokenCacheFile) 'graph sign-in cached'

# --- Guided mode -> local verify ---
Feed @('2', '4', 'y', $src, "$root/dst", 'n', 'm')
$e = Run { Start-GuidedMode }
Assert (-not $e -and $global:Inputs.Count -eq 0) "guided mode opens the right option $e"

# --- Settings toggles persist ---
Feed @('3', '2', '1', '')
$e = Run { Show-SettingsScreen }
Import-AppSettings
Assert (-not $e -and $script:Settings.ConsoleDetail -eq 'Compact' -and -not $script:Settings.ShowGuides -and $script:Settings.Theme -eq 'Ascii') "settings saved $e"

# --- Job history screen ---
Feed @('E', 'C', 'YES', '')
$e = Run { Show-JobHistory }
Assert (-not $e) "job history screen $e"

Write-Output ""
Write-Output ("RESULT: {0} passed, {1} failed" -f $global:Pass, $global:Fail)
if ($global:Fail -gt 0) { exit 1 }
