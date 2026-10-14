<#
    Engine tests for Cross-Tenant-Cloud-FileManager.ps1.
    Runs the real copy / verify / download / archive / upload engines against temp folders
    and an in-memory fake SharePoint (MockPnP.ps1). Nothing outside the temp folder is touched.
    Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-EngineTests.ps1
#>
param([string]$Script = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Cross-Tenant-Cloud-FileManager.ps1'))
$ErrorActionPreference = 'Continue'
$root = Join-Path ([System.IO.Path]::GetTempPath()) 'rhfm-engine-tests'
if (Test-Path $root) { Remove-Item $root -Recurse -Force }
$null = New-Item -ItemType Directory -Path $root
. (Join-Path $PSScriptRoot 'MockPnP.ps1')
. $Script -WorkDir "$root/wd" -LibraryMode
$script:PnPReady = $true

$global:Pass = 0; $global:Fail = 0
function Assert { param([bool]$Cond, [string]$Name) if ($Cond) { $global:Pass++; Write-Output "PASS  $Name" } else { $global:Fail++; Write-Output "FAIL  $Name" } }
function Quiet { param([scriptblock]$B) & $B 6>$null }

# ---------------- pure helpers ----------------
$s = ConvertFrom-SelectionText -Text '1, 3-5, 10-8' -Count 10
Assert ($s.Ok -and ($s.Indices -join ',') -eq '0,2,3,4,7,8,9') 'selection ranges + reverse + sort'
$s = ConvertFrom-SelectionText -Text '1 3 7' -Count 10
Assert ($s.Ok -and ($s.Indices -join ',') -eq '0,2,6') 'selection space separated'
$s = ConvertFrom-SelectionText -Text '2 - 4' -Count 10
Assert ($s.Ok -and ($s.Indices -join ',') -eq '1,2,3') 'selection spaced range'
$s = ConvertFrom-SelectionText -Text 'a' -Count 3
Assert ($s.Ok -and $s.Indices.Count -eq 3) 'selection all'
$s = ConvertFrom-SelectionText -Text '11' -Count 10
Assert (-not $s.Ok) 'selection out of range rejected'
$s = ConvertFrom-SelectionText -Text 'x' -Count 10
Assert (-not $s.Ok) 'selection garbage rejected'

Assert ((Format-Duration 119) -eq '1m 59s') 'duration 119s'
Assert ((Format-Duration 3599) -eq '59m 59s') 'duration 3599s'
Assert ((Format-Duration 7199) -eq '1h 59m') 'duration 7199s'
Assert ((Format-Duration 59.6) -eq '59s') 'duration 59.6s'
Assert ((Format-FileSize 1536) -match '^1[.,]5 KB$') 'size KB'

Assert ((ConvertTo-CleanPath '"C:\Data\x\"') -eq 'C:\Data\x') 'clean quoted path'
Assert ((ConvertTo-CleanPath 'C:\') -eq 'C:\') 'drive root kept'
Assert ((ConvertTo-CleanPath 'C:') -eq 'C:\') 'bare drive gets slash'
Assert ((ConvertTo-CleanPath "& 'D:\A B'") -eq 'D:\A B') 'ISE drag-drop'
Assert ((Get-SafeWindowsName 'a:b?.txt') -eq 'a-b-.txt') 'safe name'
Assert ((Get-SafeWindowsName 'CON.txt') -eq '_CON.txt') 'reserved name'
$used = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$a = Get-UniqueLocalName -Name 'a-b.txt' -Used $used; $b = Get-UniqueLocalName -Name 'A-B.txt' -Used $used
Assert ($a -eq 'a-b.txt' -and $b -eq 'A-B (2).txt') 'unique names'

$p = Parse-CloudUrl -Url 'https://contoso.sharepoint.com/sites/HR/Shared%20Documents/Forms/AllItems.aspx?viewid=123&id=%2Fsites%2FHR%2FShared%20Documents%2FR%26D%20Team' -Quiet
Assert ($p.SiteUrl -eq 'https://contoso.sharepoint.com/sites/HR' -and $p.FolderPath -eq '/sites/HR/Shared Documents/R&D Team') 'url id= with & and viewid'
$p = Parse-CloudUrl -Url 'https://contoso.sharepoint.com/teams/Fin/Shared%20Documents/Forms/AllItems.aspx' -Quiet
Assert ($p.SiteUrl -eq 'https://contoso.sharepoint.com/teams/Fin' -and $p.FolderPath -eq '/teams/Fin/Shared Documents') 'url teams + view page'
$p = Parse-CloudUrl -Url 'https://contoso-my.sharepoint.com/personal/john_smith_contoso_com/_layouts/15/onedrive.aspx' -Quiet
Assert ($p.FolderPath -eq '/personal/john_smith_contoso_com/Documents' -and $p.Note) 'url onedrive page'
$p = Parse-CloudUrl -Url 'https://contoso.sharepoint.com/:f:/r/sites/HR/Shared%20Documents/A?csf=1&web=1' -Quiet
Assert ($p.FolderPath -eq '/sites/HR/Shared Documents/A') 'url /:f:/r/ link'
Assert ($null -eq (Parse-CloudUrl -Url 'C:\temp' -Quiet)) 'url local path rejected'
Assert ((Get-SiteRelativePath -SiteUrl 'https://x.sharepoint.com/sites/HR' -ServerRelativePath '/sites/HR/Documents/A') -eq 'Documents/A') 'site relative'
Assert ((Get-SiteRelativePath -SiteUrl 'https://x.sharepoint.com' -ServerRelativePath '/Shared Documents/A') -eq 'Shared Documents/A') 'site relative root'
Assert ((ConvertTo-OneDriveSiteUrl -TenantName 'https://Contoso.sharepoint.com' -Upn "Anna-Maria.O'Neil@contoso.co.uk") -eq 'https://contoso-my.sharepoint.com/personal/anna_maria_o_neil_contoso_co_uk') 'onedrive url from upn'
Assert ((Get-JwtAccountName ('x.' + ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"upn":"a@b.com","x":"???>>>"}')).Replace('+','-').Replace('/','_').TrimEnd('=')) + '.y')) -eq 'a@b.com') 'jwt base64url decode'

# ---------------- display: nothing wider than the console ----------------
$W = Get-UiWidth
$out = & {
    Write-Panel -Title 'T' -Lines @(('x' * 300), @{ Label = 'Label'; Value = ('/very/long/path/' * 20) }, '-', @{ Text = ('y' * 50); Color = 'Red' })
    Write-SummaryBox -Title ('Z' * 200) -Lines @(@{ Label = 'A'; Value = 'B' })
    Write-Banner -Title 'Title' -Right 'Right' -Subtitle ('s ' * 100)
    Show-MainMenu
} 6>&1 | Out-String
$lines = $out -split "`r?`n"
$tooWide = @($lines | Where-Object { $_.Length -gt $W })
Assert ($tooWide.Count -eq 0) "no rendered line wider than $W (found $($tooWide.Count))"

# ---------------- favourites ----------------
Assert ((Save-Favorites -Favorites @([pscustomobject]@{ Name = 'One'; Type = 'Local'; Value = 'C:\A'; Added = '2024' }))) 'save one favourite'
Assert ((Get-Favorites).Count -eq 1 -and (Get-Favorites)[0].Value -eq 'C:\A') 'load one favourite'
$null = Add-Favorite -Name 'Two' -Type 'Cloud' -Value 'https://x/sites/a'
$f = (Get-Favorites)
Assert ($f.Count -eq 2 -and $f[1].Name -eq 'Two' -and $f[0].Name -eq 'One') 'load two favourites (PS5 array quirk)'
[IO.File]::WriteAllText($FavoritesFile, '[{"Name":["A","B"],"Type":["Local","Cloud"],"Value":["C:\\X","https://y"],"Added":["1","2"]}]')
$f = (Get-Favorites)
Assert ($f.Count -eq 2 -and $f[1].Value -eq 'https://y') 'repair merged legacy favourites'

# ---------------- local engines ----------------
$src = "$root/src"; $dst = "$root/dst"
$null = New-Item -ItemType Directory -Path "$src/sub/deep", "$src/empty" -Force
[IO.File]::WriteAllText("$src/a.txt", 'hello'); [IO.File]::WriteAllText("$src/sub/b.txt", 'world!'); [IO.File]::WriteAllText("$src/sub/deep/c.bin", ('z' * 5000))
$r = Quiet { Perform-LocalTransfer -Src $src -Dst $dst }
Assert ($r.Success -eq 3 -and $r.Failed -eq 0 -and (Test-Path "$dst/sub/deep/c.bin") -and (Test-Path "$dst/empty")) 'local copy incl. empty folder'
$r = Quiet { Perform-LocalTransfer -Src $src -Dst $dst }
Assert ($r.Skipped -eq 3 -and $r.Success -eq 0) 'local copy resume skips identical'
$v = Quiet { Perform-LocalVerify -Src $src -Dst $dst }
Assert ($v.Perfect -eq 3 -and $v.Mismatch -eq 0 -and $v.MissingInDst -eq 0) 'local verify match'
$v = Quiet { Perform-LocalVerify -Src $src -Dst $dst -Hash }
Assert ($v.Perfect -eq 3) 'local verify with hash'
[IO.File]::WriteAllText("$dst/sub/b.txt", 'WORLD!')
[IO.File]::SetLastWriteTimeUtc("$dst/sub/b.txt", [IO.File]::GetLastWriteTimeUtc("$src/sub/b.txt"))
$v = Quiet { Perform-LocalVerify -Src $src -Dst $dst -Hash }
Assert ($v.Mismatch -eq 1) 'hash verify catches same-size different content'
$r = Quiet { Perform-LocalTransfer -Src $src -Dst "$src/inner" -IsMove }
Assert ($r.Failed -eq 1 -and (Test-Path "$src/a.txt")) 'move into itself refused, nothing deleted'
$r = Quiet { Perform-LocalTransfer -Src $src -Dst $src -IsMove }
Assert ((Test-Path "$src/sub/b.txt")) 'move onto itself refused'
[IO.File]::WriteAllText("$src/self.txt", 'keep me')
$r = Quiet { Perform-LocalTransfer -Src "$src/self.txt" -Dst $src -IsMove }
Assert ($r.Failed -eq 1 -and [IO.File]::ReadAllText("$src/self.txt") -eq 'keep me') 'single-file move onto itself refused'
Remove-Item "$src/self.txt"
$mv = "$root/moved"
$r = Quiet { Perform-LocalTransfer -Src $dst -Dst $mv -IsMove }
Assert ($r.Success -eq 3 -and -not (Test-Path $dst) -and (Test-Path "$mv/sub/deep/c.bin")) 'local move removes emptied source'
# search
$res = $null
function Show-SearchResults { param($Results, $Term, $Kind) $global:LastResults = $Results }
Quiet { Perform-LocalSearch -SearchPath $src -SearchTerm '*.txt' }
Assert (@($global:LastResults).Count -eq 2) 'local search *.txt'
[IO.File]::WriteAllText("$src/budget [final].xlsx", 'x'); [IO.File]::WriteAllText("$src/old.xls", 'x')
Quiet { Perform-LocalSearch -SearchPath $src -SearchTerm 'budget [final]' }
Assert (@($global:LastResults).Count -eq 1) 'search with brackets is literal'
Quiet { Perform-LocalSearch -SearchPath $src -SearchTerm '*.xls' }
Assert (@($global:LastResults).Count -eq 1 -and $global:LastResults[0].Name -eq 'old.xls') '*.xls does not match .xlsx'
# wipe guard
Assert (Test-ProtectedFolder "$root/wd") 'workdir is protected'
$st = New-JobStats
Quiet { Remove-LocalItem -Path $mv -Stats $st }
Assert (-not (Test-Path $mv) -and $st.Failed -eq 0) 'remove local tree'

# ---------------- cloud engines with the fake PnP ----------------
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'
Add-FakeFolder '/sites/HR/Documents/Forms'; Add-FakeFile '/sites/HR/Documents/Forms/AllItems.aspx' 'view'
Add-FakeFolder '/sites/HR/Documents/Proj'
Add-FakeFile '/sites/HR/Documents/Proj/a-b.txt' 'first'
Add-FakeFile '/sites/HR/Documents/Proj/a:b.txt' 'second!'
Add-FakeFolder '/sites/HR/Documents/Proj/Sub'
Add-FakeFile '/sites/HR/Documents/Proj/Sub/report.pdf' ('p' * 3000)
Add-FakeFile '/sites/HR/Documents/top.txt' 'top'
Add-FakeFolder '/sites/HR/Apps'; Add-FakeFile '/sites/HR/Apps/app.cfg' 'cfg'
$cd = @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/Proj' }
$buf = "$root/buf"; $null = New-Item -ItemType Directory $buf
$dl = Quiet { Perform-DirectDownload -CloudData $cd -LocalDest "$root/dl" -BufferDest $buf -Manifest -NoRegister }
Assert ($dl.Success -eq 3 -and $dl.Failed -eq 0 -and (Test-Path "$root/dl/a-b (2).txt") -and (Test-Path "$root/dl/Sub/report.pdf")) 'download with name-clash handling'
Assert ([IO.File]::ReadAllText("$root/dl/a-b.txt") -eq 'first' -and [IO.File]::ReadAllText("$root/dl/a-b (2).txt") -eq 'second!') 'clash keeps both files'
Assert (([IO.File]::GetLastWriteTimeUtc("$root/dl/Sub/report.pdf")).Year -eq 2024) 'download keeps cloud date'
$v = Quiet { Test-ArchiveResult -CloudData $cd -DlStats $dl -Quiet }
Assert ($v.Passed -and $v.CloudFiles -eq 3 -and $v.LocalFiles -eq 3) 'archive verify passes on disk'
[IO.File]::WriteAllText("$root/dl/Sub/report.pdf", 'tampered')
$v = Quiet { Test-ArchiveResult -CloudData $cd -DlStats $dl -Quiet }
Assert (-not $v.Passed -and $v.WrongSize -eq 1) 'archive verify catches wrong size on disk'
Remove-Item "$root/dl/a-b.txt"
$v = Quiet { Test-ArchiveResult -CloudData $cd -DlStats $dl -Quiet }
Assert (-not $v.Passed -and $v.Missing -ge 1) 'archive verify catches missing file on disk'
$global:FakeFailList['documents/proj/sub'] = $true
$dl2 = Quiet { Perform-DirectDownload -CloudData $cd -LocalDest "$root/dl2" -BufferDest $buf -Manifest -NoRegister }
Assert ($dl2.ListingErrors -eq 1) 'listing failure counted (not treated as empty)'
$v = Quiet { Test-ArchiveResult -CloudData $cd -DlStats $dl2 -Quiet }
Assert (-not $v.Passed) 'archive blocked after listing failure'
$global:FakeFailList.Clear()

# Full archive of one folder -> recycled
$res = Quiet { Invoke-ArchiveOne -CloudData $cd -LocalDest "$root/arch1" -BufferDest $buf -Compact }
Assert ($res.Deleted -and $global:FakeRecycle -contains '/sites/HR/Documents/Proj' -and -not $global:Fake.ContainsKey('/sites/hr/documents/proj')) 'archive one -> verified -> recycled'
# Library root: contents recycled, library kept, Forms untouched
$res = Quiet { Invoke-ArchiveOne -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents' } -LocalDest "$root/arch2" -BufferDest $buf -Compact }
Assert ($res.Deleted -and $global:Fake.ContainsKey('/sites/hr/documents') -and $global:Fake.ContainsKey('/sites/hr/documents/forms') -and -not $global:Fake.ContainsKey('/sites/hr/documents/top.txt')) 'library root kept, contents recycled, Forms skipped'
Assert (-not (Test-Path "$root/arch2/Forms")) 'Forms folder not downloaded'
# Single file archive uses Remove-PnPFile
Add-FakeFile '/sites/HR/Documents/one.docx' 'doc'
$res = Quiet { Invoke-ArchiveOne -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/one.docx' } -LocalDest "$root/arch3" -BufferDest $buf -Compact }
Assert ($res.Deleted -and (Test-Path "$root/arch3/one.docx") -and -not $global:Fake.ContainsKey('/sites/hr/documents/one.docx')) 'single file archive'
# Remove-CloudItem refuses library root
$threw = $false; try { Remove-CloudItem -SiteUrl $global:FakeSite -ServerRelativeUrl '/sites/HR/Documents' -IsFolder } catch { $threw = $true }
Assert $threw 'library root delete refused'

# Leaver archive (Documents into root, Apps into sub-folder, Site Pages ignored)
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFolder '/sites/HR/Documents/Apps'; Add-FakeFile '/sites/HR/Documents/Apps/x.txt' 'docs-apps'
Add-FakeFile '/sites/HR/Documents/cv.pdf' 'cv'
Add-FakeFolder '/sites/HR/Apps'; Add-FakeFile '/sites/HR/Apps/app.cfg' 'cfg'
Add-FakeFolder '/sites/HR/SitePages'; Add-FakeFile '/sites/HR/SitePages/home.aspx' 'x'
$lr = Quiet { Perform-LeaverOneDriveArchive -LeaverSiteUrl $global:FakeSite -LocalDest "$root/leaver" -BufferDest $buf }
Assert ($lr.Libraries -eq 2 -and $lr.LibOk -eq 2 -and $lr.Deleted -eq 3) "leaver archive outcomes (libs=$($lr.Libraries) ok=$($lr.LibOk) deleted=$($lr.Deleted))"
Assert ((Test-Path "$root/leaver/cv.pdf") -and (Test-Path "$root/leaver/Apps/x.txt") -and (Test-Path "$root/leaver/Apps (library)/app.cfg")) 'leaver library folder clash avoided'
Assert ($global:Fake.ContainsKey('/sites/hr/sitepages/home.aspx')) 'system library untouched'

# Cancel stops the whole leaver run
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFile '/sites/HR/Documents/a.txt' 'a'
Add-FakeFolder '/sites/HR/Apps'; Add-FakeFile '/sites/HR/Apps/b.txt' 'b'
$script:JobDepth = 1; $script:JobCancelled = $true
$lr = Quiet { Perform-LeaverOneDriveArchive -LeaverSiteUrl $global:FakeSite -LocalDest "$root/leaver2" -BufferDest $buf }
$script:JobDepth = 0; $script:JobCancelled = $false
Assert ($lr.Cancelled -and $global:FakeRecycle.Count -eq 0) 'cancel stops leaver run, nothing recycled'

# Cloud verify / batch helpers
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFolder '/sites/HR/Documents/A'; Add-FakeFile '/sites/HR/Documents/A/f.txt' '12345'
Add-FakeFolder '/sites/HR/Documents/B'; Add-FakeFile '/sites/HR/Documents/B/f.txt' '1234'; Add-FakeFile '/sites/HR/Documents/B/extra.txt' 'e'
$si = Quiet { Get-CloudInventory -SiteUrl $global:FakeSite -FolderPath '/sites/HR/Documents/A' -Quiet }
$di = Quiet { Get-CloudInventory -SiteUrl $global:FakeSite -FolderPath '/sites/HR/Documents/B' -Quiet }
$cmp = Compare-CloudInventories -SrcMap (ConvertTo-InventoryMap $si) -DstMap (ConvertTo-InventoryMap $di)
Assert ($cmp.Mismatch -eq 1 -and $cmp.MissingInSrc -eq 1 -and $cmp.Perfect -eq 0) 'cloud compare'
$global:FakeCopyJobError = $true
$threw = $false; try { Quiet { Invoke-CloudCopyMove -SourceUrl '/sites/HR/Documents/A' -TargetUrl '/sites/HR/Documents/B' -Name 'A' } } catch { $threw = $true }
Assert $threw 'copy job error is detected'
$global:FakeCopyJobError = $false

# Upload with skip on second run and junk skipping
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'
$up = "$root/up"; $null = New-Item -ItemType Directory "$up/s1/s2", "$up/emptydir" -Force
[IO.File]::WriteAllText("$up/one.txt", '1'); [IO.File]::WriteAllText("$up/s1/s2/two.txt", '22'); [IO.File]::WriteAllText("$up/~`$lock.docx", 'l'); [IO.File]::WriteAllText("$up/Thumbs.db", 't')
$u = Quiet { Perform-DirectUpload -CloudDst @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/Up' } -LocalSrc $up }
Assert ($u.Success -eq 2 -and $u.Skipped -eq 2 -and $global:Fake.ContainsKey('/sites/hr/documents/up/s1/s2/two.txt') -and $global:Fake.ContainsKey('/sites/hr/documents/up/emptydir')) "upload (ok=$($u.Success) skipped=$($u.Skipped))"
$global:FakeUploads.Clear()
$u = Quiet { Perform-DirectUpload -CloudDst @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/Up' } -LocalSrc $up }
Assert ($u.Success -eq 0 -and $u.Skipped -eq 4 -and $global:FakeUploads.Count -eq 0) 'upload resume skips existing'

# Archive resume: identical file without the SharePoint date is adopted (content check),
# a different same-size file is never overwritten and blocks the delete.
Reset-Fake
Add-FakeFolder '/sites/HR/Documents'; Add-FakeFolder '/sites/HR/Documents/R'
Add-FakeFile '/sites/HR/Documents/R/same.txt' 'abcdef'; Add-FakeFile '/sites/HR/Documents/R/other.txt' '123456'
$null = New-Item -ItemType Directory "$root/resume" -Force
[IO.File]::WriteAllText("$root/resume/same.txt", 'abcdef'); [IO.File]::WriteAllText("$root/resume/other.txt", 'XXXXXX')
$rd = Quiet { Perform-DirectDownload -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/R' } -LocalDest "$root/resume" -BufferDest $buf -Manifest -NoRegister }
Assert ($rd.Skipped -eq 1 -and $rd.Failed -eq 1 -and [IO.File]::ReadAllText("$root/resume/other.txt") -eq 'XXXXXX') 'archive resume adopts identical file, never overwrites a different one'
$res = Quiet { Invoke-ArchiveOne -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/R' } -LocalDest "$root/resume" -BufferDest $buf -Compact }
Assert (-not $res.Deleted -and $global:Fake.ContainsKey('/sites/hr/documents/r/other.txt')) 'archive with a conflicting local file does not delete'
# Plain download: same size but different date is refreshed
[IO.File]::WriteAllText("$root/plain.txt", 'x')
$null = New-Item -ItemType Directory "$root/plain" -Force; [IO.File]::WriteAllText("$root/plain/other.txt", 'XXXXXX')
$pd = Quiet { Perform-DirectDownload -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/R' } -LocalDest "$root/plain" -BufferDest $buf -NoRegister }
Assert ([IO.File]::ReadAllText("$root/plain/other.txt") -eq '123456') 'download refreshes same-size file with a different date'
$pd2 = Quiet { Perform-DirectDownload -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/R' } -LocalDest "$root/plain" -BufferDest $buf -NoRegister }
Assert ($pd2.Skipped -eq 2 -and $pd2.Success -eq 0) 'second download skips (date + size match)'
# A user folder called "Forms" deeper down is NOT treated as the system folder
Add-FakeFolder '/sites/HR/Documents/R/Forms'; Add-FakeFile '/sites/HR/Documents/R/Forms/userform.pdf' 'u'
$pd3 = Quiet { Perform-DirectDownload -CloudData @{ SiteUrl = $global:FakeSite; FolderPath = '/sites/HR/Documents/R' } -LocalDest "$root/forms" -BufferDest $buf -NoRegister }
Assert (Test-Path "$root/forms/Forms/userform.pdf") 'user folder named Forms is included'
# Sub-site library roots are protected; sharing links are rejected
$threw = $false; try { Remove-CloudItem -SiteUrl $global:FakeSite -ServerRelativeUrl '/sites/HR/sub/Shared Documents' -IsFolder } catch { $threw = $true }
Assert $threw 'sub-site library root delete refused'
Assert ($null -eq (Parse-CloudUrl -Url 'https://contoso.sharepoint.com/:f:/s/Marketing/EnAbcDef?e=x' -Quiet)) 'sharing link /:f:/s/ rejected'
# The theme table is not shadowed by a caller variable called $g
function Test-Shadow { $g = 'group name'; Write-Panel -Title 'X' -Lines @('y') }
$o = (& { Test-Shadow } 6>&1 | Out-String)
Assert ($o -match [regex]::Escape([string][char]0x2510)) 'theme not shadowed by $g'

# Leaver rows parsing
$pr = ConvertFrom-LeaverRows -Rows @("a@x.com`thttps://x-my.sharepoint.com/personal/a_x_com/_layouts/15/onedrive.aspx`tAnn: Smith", "b@x.com`thttps://x-my.sharepoint.com/personal/b_x_com`t50`t10`t20%`tBob", "c@x.com`thttps://x-my.sharepoint.com/personal/c_x_com`tBob")
Assert ($pr.Leavers.Count -eq 2 -and $pr.Leavers[0].FolderName -eq 'Ann- Smith' -and $pr.Leavers[0].SiteUrl -eq 'https://x-my.sharepoint.com/personal/a_x_com' -and $pr.Errors.Count -eq 1) 'leaver rows (sanitise, 6-col, duplicate folder)'

# Job history round-trip
$st = New-JobStats; $st.Success = 5; $st.SuccessBytes = 1000
Quiet { Register-JobResult -Operation 'Test' -Stats $st -Elapsed ([TimeSpan]::FromSeconds(3)) }
$h = (Get-JobHistory)
Assert ($h.Count -ge 1 -and $h[-1].Operation -eq 'Test' -and $h[-1].Ok -eq 5) 'job history'

Write-Output ""
Write-Output ("RESULT: {0} passed, {1} failed" -f $global:Pass, $global:Fail)
if ($global:Fail -gt 0) { exit 1 }
